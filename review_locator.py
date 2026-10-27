"""Document-wide quotation matching. Similarity produces candidates, never verified anchors."""
import re
import unicodedata
from difflib import SequenceMatcher


def normalized(text):
    chars, positions = [], []
    punctuation = str.maketrans({'“':'"', '”':'"', '‘':"'", '’':"'", '。':'.', '−':'-'})
    for index, char in enumerate(text):
        if char.isspace():
            continue
        value = unicodedata.normalize('NFKC', char).translate(punctuation)
        for item in value:
            chars.append(item); positions.append(index)
    return ''.join(chars), positions


def hits(text, quote, formatted=False):
    source, positions = normalized(text) if formatted else (text, list(range(len(text))))
    needle = normalized(quote)[0] if formatted else quote
    if not needle:
        return []
    result, cursor = [], 0
    while True:
        start = source.find(needle, cursor)
        if start < 0:
            break
        end = start + len(needle)
        # Do not locate 10000 inside 100000, or 100 yuan inside 2100 yuan.
        left_numeric = start > 0 and (source[start-1].isdigit() or source[start-1] in '+-−＋－' or
                       (source[start-1] in '.,，．' and start > 1 and source[start-2].isdigit()))
        right_numeric = end < len(source) and (source[end].isdigit() or
                        (source[end] in '.,，．' and end+1 < len(source) and source[end+1].isdigit()))
        left_ok = not (needle[0].isdigit() and left_numeric)
        right_ok = not (needle[-1].isdigit() and right_numeric)
        if left_ok and right_ok:
            result.append((positions[start], positions[end-1]+1))
        cursor = start + 1
    return result


def approximate_candidates(quote, claims, limit=3):
    needle, _ = normalized(quote)
    if len(needle) < 6:
        return []
    result = []
    for claim in claims:
        text = str(claim.get('text') or '')
        source, positions = normalized(text)
        if not source:
            continue
        matcher = SequenceMatcher(None, needle, source, autojunk=False)
        longest = matcher.find_longest_match()
        if longest.size < min(8, len(needle)):
            continue
        best = (0, 0, 0)
        # Matching blocks suggest local windows even when the paragraph is much longer.
        for block in matcher.get_matching_blocks():
            if not block.size:
                continue
            start = max(0, block.b - block.a)
            for change in (-max(2,len(needle)//8), 0, max(2,len(needle)//8)):
                end = min(len(source), start + len(needle) + change)
                score = SequenceMatcher(None, needle, source[start:end], autojunk=False).ratio()
                if score > best[0]: best = (score, start, end)
        score, start, end = best
        if score < .62 or end <= start:
            continue
        actual = text[positions[start]:positions[end-1]+1]
        qnums = re.findall(r'[-+]?\d+(?:[,.]\d+)*', needle)
        nums = re.findall(r'[-+]?\d+(?:[,.]\d+)*', normalized(actual)[0])
        result.append({'claimId':claim['id'], 'location':claim.get('location',''), 'quote':actual,
                       'modelQuote':quote, 'score':round(score,3), 'numberDifference':qnums!=nums,
                       'alignment':'approximate','start':positions[start],'end':positions[end-1]+1})
    return sorted(result, key=lambda row:row['score'], reverse=True)[:limit]


def locate(target, claims):
    quote = target.get('quote')
    if not isinstance(quote,str) or not quote.strip():
        return None, {'modelQuote':str(quote or ''),'reason':'缺少原句','candidates':[]}
    reported = target.get('claimId')
    cid = str(reported).strip() if reported is not None else ''
    # Normalize syntax only. A quote must still match the resulting paragraph.
    match = re.fullmatch(r'\[?([cCpP]?)(\d+)\]?',cid)
    normalized_id = f'C{int(match[2]):03d}' if match else cid
    preferred = next((c for c in claims if c['id']==normalized_id),None)
    for formatted in (False,True):
        found = []
        for claim in claims:
            for start,end in hits(str(claim.get('text') or ''),quote,formatted):
                found.append((claim,start,end))
        chosen = [row for row in found if row[0] is preferred]
        if len(chosen)==1:
            claim,start,end = chosen[0]
        elif len(found)==1:
            claim,start,end = found[0]
        else:
            if found:
                candidates=[{'claimId':c['id'],'location':c.get('location',''),'quote':c['text'][a:b],
                             'modelQuote':quote,'score':1.0,'alignment':'ambiguous','start':a,'end':b} for c,a,b in found[:3]]
                return None, {'modelQuote':quote,'reason':'原句在多处出现，无法唯一定位','candidates':candidates}
            continue
        actual = claim['text'][start:end]
        method = 'format' if formatted else 'exact'
        if claim['id'] != reported: method = 'global-' + method
        anchor={'claimId':claim['id'],'quote':actual,'location':claim.get('location',''),
                'alignment':method,'start':start,'end':end}
        if actual!=quote: anchor['modelQuote']=quote
        if claim['id']!=reported: anchor['reportedClaimId']=reported
        return anchor, None
    return None, {'modelQuote':quote,'reason':'未找到逐字原句，以下仅为近似候选',
                  'candidates':approximate_candidates(quote,claims)}
