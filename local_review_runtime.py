"""Bounded local-model streaming, cancellable by review id. No document text is logged."""
import json
import copy
import threading
import time
import urllib.error
import urllib.request

_lock = threading.Lock()
_model_lock = threading.Lock()
_jobs = {}


def progress(review_id):
    with _lock:
        job = _jobs.get(review_id)
        if not job:
            return {'stage': 'idle', 'seconds': 0, 'chars': 0}
        return {k: v for k, v in job.items() if k not in ('cancel', 'started')} | {'seconds': round(time.monotonic() - job['started'], 1)}


def cancel(review_id):
    with _lock:
        job = _jobs.get(review_id)
        if job:
            job['cancel'].set()


def complete_items(text):
    """Salvage only complete objects in the issues array, never repair quoted text."""
    import re
    match = re.search(r'"issues"\s*:\s*\[', text)
    if not match:
        return []
    decoder, rows, rest = json.JSONDecoder(), [], text[match.end():]
    while rest:
        rest = rest.lstrip(' \r\n\t,')
        try:
            item, end = decoder.raw_decode(rest)
        except ValueError:
            break
        if not isinstance(item, dict):
            break
        rows.append(item); rest = rest[end:]
    return rows


def chat(url, payload, headers, review_id, deadline=120, on_items=None):
    if not _model_lock.acquire(blocking=False):
        raise ValueError('模型仍在处理另一个审核请求，请稍后继续。')
    started, text, finish, usage = time.monotonic(), '', None, {}
    published = 0
    job = {'started': started, 'cancel': threading.Event(), 'stage': 'reading', 'chars': 0}
    with _lock:
        if len(_jobs) >= 64:
            for key in list(_jobs):
                if _jobs[key]['stage'] in ('done', 'error'):
                    del _jobs[key]
                    break
        _jobs[review_id] = job
    try:
        body = json.dumps({**payload, 'stream': True, 'stream_options': {'include_usage': True}}).encode('utf-8')
        request = urllib.request.Request(url, body, {'Content-Type': 'application/json', **headers})
        try:
            response = urllib.request.urlopen(request, timeout=40)
        except urllib.error.HTTPError as error:
            detail = error.read(4096).decode('utf-8', errors='replace')
            # Only identify well-known operational errors; never echo server request bodies.
            if any(x in detail.lower() for x in ('context length','context window','too many tokens','exceed')):
                cause = '全文超过模型实际上下文；请提高当前容量或换用更大上下文模型，原文不会截断'
            else:
                cause = '模型被卸载或不可用' if any(x in detail.lower() for x in ('unload', 'not loaded', 'model not found')) else '请检查模型服务错误记录'
            raise ValueError(f'模型服务 HTTP {error.code}：{cause}。') from error
        with response:
            if 'text/event-stream' not in response.headers.get('Content-Type', ''):
                result = json.load(response)
                message = (result.get('choices') or [{}])[0].get('message') or {}
                if message.get('reasoning_content') or message.get('reasoning'):
                    raise ValueError('模型服务仍返回思考内容，当前内容未计为完成。')
                job['stage'] = 'done'
                return result
            for line in response:
                if job['cancel'].is_set():
                    raise ValueError('审核已停止；已生成的完整意见会保留，当前原文未计为完成。')
                if time.monotonic() - started >= deadline:
                    raise TimeoutError('本次全文审核已达到等待上限，已停止请求。')
                if not line.startswith(b'data:'):
                    continue
                data = line[5:].strip()
                if data == b'[DONE]':
                    break
                event = json.loads(data)
                usage = event.get('usage') or usage
                for choice in event.get('choices') or []:
                    delta = choice.get('delta') or {}
                    if delta.get('reasoning_content') or delta.get('reasoning'):
                        raise ValueError('模型服务仍在输出思考内容，已停止等待；请检查模型端的思考开关。')
                    text += delta.get('content') or ''
                    if text.lstrip().lower().startswith(('<think>', '<analysis>')):
                        raise ValueError('模型仍在输出思考标签，已停止等待。')
                    finish = choice.get('finish_reason') or finish
                job.update(stage='generating', chars=len(text))
                if on_items:
                    rows = complete_items(text)
                    if len(rows) > published:
                        snapshot = on_items(rows)
                        with _lock:
                            job.update(copy.deepcopy(snapshot))
                        published = len(rows)
        if finish is None:
            raise ValueError('模型响应中断，未收到完成标记。')
        job['stage'] = 'done'
        return {'choices': [{'finish_reason': finish, 'message': {'content': text}}], 'usage': usage}
    except Exception as error:
        job['stage'] = 'error'
        rows = complete_items(text)
        if rows:
            return {'choices': [{'finish_reason': 'partial', 'message': {'content': json.dumps({'issues': rows}, ensure_ascii=False)}}],
                    '_partial_error': str(error), 'usage': usage}
        raise
    finally:
        _model_lock.release()
