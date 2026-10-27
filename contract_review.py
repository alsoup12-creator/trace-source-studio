"""Full-document review orchestration; existing model providers remain unchanged."""
import hashlib
import json
import uuid
import time
import re
from pathlib import Path
import local_review_runtime
import full_document_context
import review_locator
import copy


def fingerprint(project):
    rows = [(c['id'], c.get('text', '')) for c in project.get('claims', [])]
    return hashlib.sha256(json.dumps(rows, ensure_ascii=False).encode()).hexdigest()


def document_rows(project):
    return [{'claimId': c['id'], 'text': str(c.get('text') or '')}
            for c in project.get('claims', []) if str(c.get('text') or '').strip()]


INSTRUCTION = Path(__file__).with_name('full-document-review-prompt.md').read_text(encoding='utf-8')
LOCAL_INSTRUCTION = INSTRUCTION
API_INSTRUCTION = Path(__file__).with_name('api-legal-review-prompt.md').read_text(encoding='utf-8')


def instruction_for(provider):
    return API_INSTRUCTION if provider == 'api' else LOCAL_INSTRUCTION


def start(project, provider, stance, engine, extra_checks=0):
    if provider not in ('api', 'local'):
        raise ValueError('请选择本地小模型或 API 模型。')
    if not isinstance(extra_checks, int) or not 0 <= extra_checks <= 2:
        raise ValueError('补查次数只能选择0、1或2次。')
    config = engine.load_config()
    prefix = 'localApi' if provider == 'local' else 'api'
    if not config.get(prefix + 'Base') or not config.get(prefix + 'Model'):
        raise ValueError('请先填写模型地址和名称。')
    rows = document_rows(project)
    if not rows:
        raise ValueError('文档没有可审核文字。')
    limit, source = full_document_context.capacity(config, provider)
    output = min(1024 if provider == 'local' else 2048, max(128, limit // 4))
    content = '当前文件全文（全部编号共同构成同一份文档）：\n' + json.dumps(rows, ensure_ascii=False)
    messages = [{'role':'system','content':instruction_for(provider)}, {'role':'user','content':content}]
    needed = full_document_context.check(messages, output, limit)
    project['contractReview'] = {
        'id': uuid.uuid4().hex, 'fingerprint': fingerprint(project), 'provider': provider,
        'model': config[prefix + 'Model'], 'base': config[prefix + 'Base'], 'plan':'full-document-v1',
        'instructionProfile': 'legal-v1' if provider == 'api' else 'proofread-v1',
        # Preserve the legacy response field without passing a stance to the model.
        'stance': '中立核对',
        'status':'ready','total':1,'completed':[],'failed':{},'ranges':[['全文']],
        'sizes':[sum(len(r['text']) for r in rows)], 'issues':[],'unlocated':[],
        'warnings':(['审核范围为提取的正文；页眉、页脚、脚注和图片请对照原件。'] if project.get('artifact',{}).get('kind')=='word' else []),
        'summary':'','crossDone':False,'createdAt':engine.now(),
        'textSnapshot':{c['id']:c.get('text','') for c in project['claims']},
        'locations':{c['id']:c.get('location','') for c in project['claims']},
        'extraChecks':extra_checks,'passCompleted':{},'passStats':[],
        'contextTokens':limit,'contextSource':source,'estimatedTokens':needed,'outputTokens':output,
    }
    return project


def normalize_local(value):
    if not isinstance(value.get('issues'), list):
        raise ValueError('本地模型没有返回意见数组；当前内容未审核完成。')
    rows = []
    aliases={'文字':'text','错别字':'text','计算':'math','数字':'math','金额':'math','逻辑':'logic','前后逻辑':'logic','风险':'risk','法律':'risk','法律风险':'risk','法律疑点':'risk','遗漏':'risk'}
    for raw in value['issues']:
        if isinstance(raw,dict):
            raw=dict(raw)
            key='c' if 'c' in raw else 'category'
            label=str(raw.get(key) or '').strip().lower()
            raw[key]=aliases.get(label,label)
        if not isinstance(raw, dict) or 'c' not in raw:
            rows.append(raw); continue
        targets = [{'claimId': raw.get('p'), 'quote': raw.get('q')}]
        for target in raw.get('t') if isinstance(raw.get('t'), list) else []:
            if isinstance(target, list) and len(target) == 2:
                targets.append({'claimId': target[0], 'quote': target[1]})
        reason = str(raw.get('w') or '')
        title = re.sub(r'^(明确错误|待核对|可能不一致)[：:]\s*', '', reason)[:36]
        rows.append({'category': raw.get('c'), 'severity': 'high' if raw.get('c') in ('math', 'logic') else 'normal',
                     'title': title, 'reason': reason, 'suggestion': raw.get('s'),
                     'targets': targets, 'replacement': raw.get('r') if raw.get('c') == 'text' else None})
    return {'issues': rows, 'summary': f'本次获得 {len(rows)} 条候选意见，请对照原句核查。', 'outline': ''}


def call_model(content, review, engine, on_items=None):
    config = engine.load_config()
    prefix = 'localApi' if review['provider']=='local' else 'api'
    if config.get(prefix+'Model') != review['model'] or config.get(prefix+'Base') != review['base']:
        raise ValueError('模型配置已变更，请重新开始全文审核。')
    messages=[{'role':'system','content':instruction_for(review['provider'])},
              {'role':'user','content':content}]
    review['estimatedTokens']=full_document_context.check(messages,review['outputTokens'],review['contextTokens'])
    payload={'model':review['model'],'temperature':0.1,'max_tokens':review['outputTokens'],'messages':messages}
    if review['provider']=='local':
        payload.update(chat_template_kwargs={'enable_thinking':False},reasoning_effort='none')
    key=engine.SECRETS.get(prefix+'Key','')
    headers={'Authorization':'Bearer '+key} if key else {}
    started=time.monotonic()
    engine.web_log(f'[contract-review] mode=full-document pass={review.get("activePass",0)} input_chars={len(content)} output_limit={review["outputTokens"]} started')
    try:
        result=local_review_runtime.chat(review['base'].rstrip('/')+'/chat/completions',payload,headers,review['id'],deadline=240,on_items=on_items)
    except Exception as error:
        engine.web_log(f'[contract-review] elapsed={time.monotonic()-started:.1f}s failed={type(error).__name__}')
        raise
    choice=(result.get('choices') or [{}])[0]
    text=choice.get('message',{}).get('content') or ''
    review['lastCall']={'seconds':round(time.monotonic()-started,1),'usage':result.get('usage') or {},'finish':choice.get('finish_reason'),'replyChars':len(text)}
    engine.web_log(f'[contract-review] elapsed={review["lastCall"]["seconds"]}s finish={choice.get("finish_reason")} reply_chars={len(text)}')
    partial=result.get('_partial_error')
    if choice.get('finish_reason') in ('length','content_filter'):
        rows=local_review_runtime.complete_items(text)
        if not rows:
            raise ValueError('全文审核输出未完整返回，请重试；不会标为审核完成。')
        text=json.dumps({'issues':rows},ensure_ascii=False)
        partial='输出未完整返回，完整意见已保留，全文审核尚未完成。'
    try:
        value=normalize_local(engine.parse_json_object(text))
    except (ValueError,TypeError) as error:
        # Some small models finish the array but omit only the outer object brace.
        # Accept only if adding exactly that brace makes the entire JSON valid.
        try:
            repaired=json.loads(text.strip()+'}')
            if not isinstance(repaired,dict) or not isinstance(repaired.get('issues'),list):
                raise ValueError('invalid issue array')
            value=normalize_local(repaired)
        except (ValueError,TypeError):
            raise ValueError('模型没有返回可解析的逐条意见，全文审核尚未完成。') from error
    if partial:value['_partial']=partial
    return value


def align_whitespace(text, quote):
    if quote in text:
        return quote
    positions = [i for i, char in enumerate(text) if not char.isspace()]
    normalized = ''.join(text[i] for i in positions)
    needle = ''.join(char for char in quote if not char.isspace())
    if not needle or normalized.count(needle) != 1:
        return None
    start = normalized.index(needle)
    return text[positions[start]:positions[start + len(needle) - 1] + 1]


def validate_issues(value, project, allowed=None, return_unlocated=False):
    valid, unlocated = [], []
    for raw in value['issues']:
        if not isinstance(raw,dict):raw={}
        targets=raw.get('targets') if isinstance(raw.get('targets'),list) else []
        anchors,unresolved=[],[]
        for target in targets:
            if not isinstance(target,dict):
                unresolved.append({'modelQuote':'','reason':'无效引文结构','candidates':[]});continue
            anchor,pending=review_locator.locate(target,project['claims'])
            if anchor is not None:
                # A whole-document request can resolve wrong ids anywhere in its input.
                if allowed is not None and not any(anchor['claimId']==r['claimId'] and anchor['quote'] in r['text'] for r in allowed):
                    unresolved.append({'modelQuote':str(target.get('quote') or ''),'reason':'不在本次实际输入范围','candidates':[]})
                elif not any(a['claimId']==anchor['claimId'] and a['quote']==anchor['quote'] for a in anchors):
                    anchors.append(anchor)
            else:unresolved.append(pending)
        fields={k:str(raw.get(k) or '').strip() for k in ('title','reason','suggestion')}
        category=raw.get('category') if raw.get('category') in ('text','math','logic','risk') else 'risk'
        candidates=[c for u in unresolved for c in u.get('candidates',[])]
        if not all(fields.values()) or not targets or (not anchors and not candidates):
            unlocated.append({**fields,'title':fields['title'] or '审核疑点待核对',
                'reason':fields['reason'] or '没有可核对的原句或必要字段。',
                'suggestion':fields['suggestion'] or '请人工核查。','category':category,
                'quotes':[str(t.get('quote') or '') for t in targets if isinstance(t,dict)],
                'unresolvedTargets':unresolved,'reportedTargets':copy.deepcopy(targets)})
            continue
        replacement=raw.get('replacement')
        if not isinstance(replacement,str) or len(anchors)!=1 or unresolved:replacement=None
        valid.append({**fields,'id':raw.get('id') or uuid.uuid4().hex[:12],
            'category':category,'severity':'high' if raw.get('severity')=='high' else 'normal',
            'targets':anchors,'unresolvedTargets':unresolved,'reportedTargets':copy.deepcopy(targets),'replacement':replacement,
            'locationStatus':'partial' if anchors and unresolved else 'approximate' if not anchors else 'exact',
            'decision':raw.get('decision') or 'pending','sourcePass':raw.get('sourcePass',0),**({'variants':raw['variants']} if raw.get('variants') else {})})
    return (valid,len(unlocated),unlocated) if return_unlocated else (valid,len(unlocated))


def relocate_review(project):
    review=ensure_current(project)
    history=review.setdefault('locationHistory',[])
    history.append({'issues':copy.deepcopy(review.get('issues',[])),'unlocated':copy.deepcopy(review.get('unlocated',[]))})
    review['locationHistory']=history[-3:]
    raw=copy.deepcopy(review.get('issues',[]))
    for row in raw:
        row['targets']=row.get('reportedTargets') or row.get('targets',[])+[{'claimId':None,'quote':u.get('modelQuote','')} for u in row.get('unresolvedTargets',[])]
    for item in review.get('unlocated',[]):
        row=copy.deepcopy(item)
        row['targets']=row.get('reportedTargets') or [{'claimId':None,'quote':q} for q in row.get('quotes',[])]
        row['replacement']=None
        raw.append(row)
    normalized=normalize_local({'issues':raw})
    issues,_,unlocated=validate_issues(normalized,project,return_unlocated=True)
    review['issues']=issues
    review['unlocated']=unlocated
    review['locationSchemaVersion']=2
    exact=sum(i['locationStatus']=='exact' for i in issues)
    partial=sum(i['locationStatus']=='partial' for i in issues)
    approximate=sum(i['locationStatus']=='approximate' for i in issues)
    review['summary']=f'{exact}条完整定位、{partial}条部分定位、{approximate}条近似候选、{len(unlocated)}条未定位。'
    return project


def ensure_current(project):
    review = project.get('contractReview')
    if not review or review['fingerprint'] != fingerprint(project):
        raise ValueError('文档已修改或审核不存在，请重新审核全文。')
    return review


def completed_passes(review, index='full'):
    return set(review.get('passCompleted',{}).get('full',[]))


def followup_content(content, review, allowed_ids, pass_index):
    if review['provider']=='api':
        focus = '另查称呼、计算、前后矛盾和相关条款中遗漏的法律疑点' if pass_index == 1 else '另查履行条件、责任、解除及争议解决等法律疑点，只针对文档实际内容提出'
    else:
        focus = '另查错别字、称呼、术语和数字计算' if pass_index == 1 else '另查称呼、日期、条件和前后表述不一致'
    prior = []
    budget = 650 if review['provider'] == 'local' else 1800
    for issue in review['issues']:
        if not any(t['claimId'] in allowed_ids for t in issue['targets']):
            continue
        item = [issue['category'], [[t['claimId'], t['quote'][:70]] for t in issue['targets'][:2]], issue['reason'][:70]]
        if len(json.dumps(prior + [item], ensure_ascii=False)) > budget:
            break
        prior.append(item)
    return (content + '\n已有意见（精简，仅用于避免重复；不代表已证实）：' + json.dumps(prior, ensure_ascii=False)
            + f'\n补查{pass_index}：还有别的具体问题吗？{focus}。只补充新疑点，不重复，不为凑数编造；没有就返回空issues。')


def merge_findings(review, issues, pass_index):
    """Group exact same anchors/categories, retaining distinct observations instead of deleting them."""
    def key(issue):
        return issue['category'], tuple(sorted((t['claimId'], t['quote']) for t in issue['targets'])), tuple(sorted(u.get('modelQuote','') for u in issue.get('unresolvedTargets',[])))
    existing = {key(i): i for i in review['issues']}
    added = variants = 0
    for issue in issues:
        issue['sourcePass'] = pass_index
        found = existing.get(key(issue))
        if found is None:
            review['issues'].append(issue); existing[key(issue)] = issue; added += 1
            continue
        signature = (issue['reason'], issue['suggestion'])
        if issue.get('replacement') != found.get('replacement'):
            found['replacement'] = None
        old = [(found['reason'], found['suggestion'])] + [(v['reason'], v['suggestion']) for v in found.get('variants', [])]
        if signature not in old:
            found.setdefault('variants', []).append({k: issue[k] for k in ('reason', 'suggestion', 'sourcePass')})
            if issue.get('replacement') != found.get('replacement'):
                found['replacement'] = None
            variants += 1
    return added, variants


def step(project, index, engine, pass_index=0):
    review=ensure_current(project)
    if review.get('plan')!='full-document-v1':
        raise ValueError('旧任务使用分段流程，请点击审核全文重新开始；原文修改仍保留。')
    if review.get('provider')=='api' and review.get('instructionProfile')!='legal-v1':
        raise ValueError('此前的 API 审核使用旧范围，请点击审核全文重新开始，以免混合不同范围的结果。')
    if index not in (None,0,'full'):
        raise ValueError('当前只有全文审核，不再运行逐段或另一次跨条款审核。')
    extra=review.get('extraChecks',0)
    if not isinstance(pass_index,int) or not 0<=pass_index<=extra:
        raise ValueError('审核轮次无效。')
    done=completed_passes(review)
    if pass_index in done:return project
    if pass_index and pass_index-1 not in done:
        raise ValueError('请先完成上一轮全文审核。')
    rows=document_rows(project)
    content='当前文件全文（全部编号共同构成同一份文档）：\n'+json.dumps(rows,ensure_ascii=False)
    if pass_index:
        content=followup_content(content,review,{r['claimId'] for r in rows},pass_index)
    review.update(activeStage='full',activePass=pass_index,status='running')
    before_cards=len(review['issues'])
    before_variants=sum(len(i.get('variants',[])) for i in review['issues'])
    def ingest(raw):
        value=normalize_local({'issues':raw})
        issues,_,unlocated=validate_issues(value,project,return_unlocated=True)
        merge_findings(review,issues,pass_index)
        for item in unlocated:
            item.update(range='全文',sourcePass=pass_index)
            if not any(old['title']==item['title'] and old['quotes']==item['quotes'] for old in review['unlocated']):
                review['unlocated'].append(item)
        return {'issues':review['issues'],'unlocated':review['unlocated']}
    try:
        value=call_model(content,review,engine,on_items=ingest)
        # Mock/test providers may already return expanded issue records.
        ingest(value['issues'])
        if value.get('_partial'):
            raise ValueError(value['_partial'])
        review['failed'].pop('full',None)
        review.setdefault('passCompleted',{}).setdefault('full',[]).append(pass_index)
        review['passStats'].append({'scope':'full','pass':pass_index,'newCards':len(review['issues'])-before_cards,
            'newObservations':sum(len(i.get('variants',[])) for i in review['issues'])-before_variants})
        review['completed']=[0]
        review['crossDone']=pass_index==extra
        review['status']='complete' if review['crossDone'] else 'ready'
        review['locationSchemaVersion']=2
        review['summary']=f'已汇总 {len(review["issues"])} 条可查看意见、{len(review["unlocated"])} 条未定位疑点。'
    except Exception as error:
        message='等待模型超时，完整意见已保留；全文结果尚未完整返回。' if isinstance(error,TimeoutError) else str(error)
        review['failed']['full']=(f'补查{pass_index}：' if pass_index else '')+engine.redact_sensitive(message)
        review['status']='partial'
    return project


def decide(project, issue_id, decision, engine):
    review = ensure_current(project)
    issue = next((i for i in review['issues'] if i['id'] == issue_id), None)
    if not issue or decision not in ('apply', 'ignore', 'pending'):
        raise ValueError('审核意见或处理方式无效。')
    if decision == 'apply':
        if issue.get('replacement') is None or len(issue['targets']) != 1 or issue.get('unresolvedTargets'):
            raise ValueError('此意见需要人工编辑对应条款。')
        target = issue['targets'][0]
        claim = next(c for c in project['claims'] if c['id'] == target['claimId'])
        if claim['text'].count(target['quote']) != 1:
            raise ValueError('引文在段落内不唯一，请人工修改。')
        engine.update_claim_text(project, claim['id'], claim['text'].replace(target['quote'], issue['replacement'], 1))
        issue['decision'] = 'applied'
        review['status'] = 'stale'
    else:
        issue['decision'] = decision
    return project
