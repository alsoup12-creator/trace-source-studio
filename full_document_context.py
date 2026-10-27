"""Capacity checks only. Never starts inference or substitutes partial text."""
import json
import urllib.parse
import urllib.request


def capacity(config, provider):
    key = 'localReviewContextTokens' if provider == 'local' else 'apiReviewContextTokens'
    manual = int(config.get(key) or 0)
    if manual > 0:
        return manual, '用户填写的当前上下文容量'
    if provider == 'local':
        base = urllib.parse.urlsplit(config['localApiBase'])
        url = urllib.parse.urlunsplit((base.scheme, base.netloc, '/api/v1/models', '', ''))
        try:
            with urllib.request.urlopen(url, timeout=3) as response:
                models = json.load(response).get('models', [])
            for model in models:
                if model.get('key') != config['localApiModel']:
                    continue
                for instance in model.get('loaded_instances') or []:
                    settings = instance.get('config') or {}
                    size = settings.get('context_length') or instance.get('context_length')
                    if isinstance(size, int) and size > 0:
                        return size, '模型服务报告的已加载上下文容量'
        except Exception:
            pass
    raise ValueError('尚不能确认当前模型的上下文容量。请先加载模型，或在接口设置底部填写当前上下文容量；不会改用分段审核。')


def estimate(messages, output_tokens):
    # Deliberately labelled an estimate: actual tokenization belongs to the provider.
    text = ''.join(message['content'] for message in messages)
    wide = sum(ord(char) > 127 for char in text)
    return int((wide * 1.5 + (len(text) - wide) / 3) * 1.15) + 256 + output_tokens


def check(messages, output_tokens, limit):
    needed = estimate(messages, output_tokens)
    if needed > limit:
        raise ValueError(f'全文与输出预留估算需约 {needed} token，超过当前填写或检测到的 {limit}。请改用更大上下文；原文不会截断、分段或改用摘要。估算可能与模型实际分词不同。')
    return needed
