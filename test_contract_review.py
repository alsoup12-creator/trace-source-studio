"""Run with: python -m unittest discover -s trace-source-studio -p test_contract_review.py -v"""
import json
import tempfile
import threading
import unittest
import urllib.request
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from unittest.mock import patch
import app
import contract_review as review
import contract_app
import local_review_runtime
import io
import full_document_context


def project():
    blocks = [{'id': 'p1', 'location': '第1段', 'text': '合同总价100000元。乙方指定帐户收款。', 'paragraphIndex': 1},
              {'id': 'p2', 'location': '第2段', 'text': '验收后付款。交付后立即付款。', 'paragraphIndex': 2}]
    return {'id': 'fictional-review-test', 'artifact': {'kind': 'word', 'path': 'fictional.docx', 'blocks': blocks},
            'claims': app.split_claims(blocks)}


def response(issues=None):
    return {'summary': '请核对文字与支付条件。', 'outline': 'C001：乙方指定帐户收款。C002：验收后付款。', 'issues': issues or []}


def issue(quote='帐户', cid='C001', replacement='账户'):
    return {'category': 'text', 'severity': 'normal', 'title': '收款用词核对', 'reason': '建议统一用词。',
            'suggestion': '改为账户。', 'targets': [{'claimId': cid, 'quote': quote}], 'replacement': replacement}


class ReviewTests(unittest.TestCase):
    def setUp(self):
        self.p=project()
        self.config={'apiBase':'http://127.0.0.1:1/v1','apiModel':'test-api','apiReviewContextTokens':32768,
                     'localApiBase':'http://127.0.0.1:1/v1','localApiModel':'test-small','localReviewContextTokens':8192}
        for name,value in [('load_config',self.config),('web_log',None)]:
            patcher=patch.object(app,name,return_value=value);patcher.start();self.addCleanup(patcher.stop)
        review.start(self.p,'api','中立核对',app)

    def test_numeric_fragment_does_not_verify_a_different_amount(self):
        import review_locator
        for source,quote in [('金额100000元','10000'),('金额2100元','100元'),('金额1.50元','50元'),('金额-100元','100元')]:
            anchor,_=review_locator.locate({'quote':quote},[{'id':'C001','text':source}])
            self.assertIsNone(anchor,(source,quote))
        self.assertTrue(review_locator.hits('金额，100元。','100元'))
        self.assertTrue(review_locator.hits('金额100。','100'))

    def test_identical_reason_with_conflicting_replacement_cannot_apply(self):
        a,_=review.validate_issues(response([issue()]),self.p)
        review.merge_findings(self.p['contractReview'],a,0)
        b,_=review.validate_issues(response([issue(replacement='其他文字')]),self.p)
        review.merge_findings(self.p['contractReview'],b,1)
        self.assertIsNone(self.p['contractReview']['issues'][0]['replacement'])

    def test_wrong_id_resolves_full_document(self):
        rows,_,_=review.validate_issues(response([issue(cid='C999')]),self.p,return_unlocated=True)
        self.assertEqual(rows[0]['targets'][0]['claimId'],'C001')
        self.assertEqual(rows[0]['targets'][0]['alignment'],'global-exact')

    def test_partial_anchor_is_retained_and_cannot_apply(self):
        row=issue();row['targets'].append({'claimId':'C999','quote':'验收后付钱。交付后立即付款。'})
        rows,_,_=review.validate_issues(response([row]),self.p,return_unlocated=True)
        self.assertEqual(rows[0]['locationStatus'],'partial')
        self.assertIsNone(rows[0]['replacement'])
        self.assertTrue(rows[0]['unresolvedTargets'][0]['candidates'])
        self.p['contractReview']['issues']=rows
        with self.assertRaises(ValueError):review.decide(self.p,rows[0]['id'],'apply',app)

    def test_ambiguous_quote_is_candidate_only(self):
        self.p['claims'].append({'id':'C003','text':'帐户','location':'第3段'})
        rows,_,_=review.validate_issues(response([issue(cid='C999')]),self.p,return_unlocated=True)
        self.assertEqual(rows[0]['locationStatus'],'approximate')
        self.assertEqual(rows[0]['targets'],[])
        self.assertEqual(len(rows[0]['unresolvedTargets'][0]['candidates']),2)

    def test_format_offsets_and_numeric_difference(self):
        import review_locator
        claims=[{'id':'C001','text':'费用为１００ 元，验收完成后支付。','location':'第1段'}]
        anchor,_=review_locator.locate({'claimId':'C999','quote':'费用为100元,验收完成后支付。'},claims)
        self.assertEqual(anchor['quote'],claims[0]['text'])
        anchor,pending=review_locator.locate({'quote':'费用为200元,验收完成后支付。'},claims)
        self.assertIsNone(anchor)
        self.assertTrue(pending['candidates'][0]['numberDifference'])

    def test_relocate_retains_partial_and_decision_without_model(self):
        self.p['contractReview']['unlocated']=[{**issue(),'quotes':['帐户','验收后付钱。交付后立即付款。'],'targets':[]}]
        snapshot=json.dumps(self.p['claims'],ensure_ascii=False)
        with patch.object(review,'call_model') as model:
            review.relocate_review(self.p)
            self.p['contractReview']['issues'][0]['decision']='ignore'
            review.relocate_review(self.p)
        model.assert_not_called()
        self.assertEqual(snapshot,json.dumps(self.p['claims'],ensure_ascii=False))
        row=self.p['contractReview']['issues'][0]
        self.assertEqual(row['locationStatus'],'partial')
        self.assertEqual(row['decision'],'ignore')

    def test_chinese_category(self):
        self.assertEqual(review.normalize_local({'issues':[{'c':'计算','q':'100','w':'核对','s':'检查'}]})['issues'][0]['category'],'math')

    def test_one_full_document_call_includes_distant_last_paragraph(self):
        self.p['claims'][0]['text']='甲'*11000
        review.start(self.p,'api','中立核对',app)
        with patch.object(review,'call_model',return_value=response([issue(cid='C002',quote='验收后付款')])) as model:
            review.step(self.p,'full',app)
        content=model.call_args[0][0]
        self.assertIn('甲'*11000,content)
        self.assertIn('交付后立即付款',content)
        self.assertEqual(model.call_count,1)
        self.assertEqual(self.p['contractReview']['status'],'complete')

    def test_no_separate_cross_stage(self):
        with self.assertRaises(ValueError):review.step(self.p,'cross',app)

    def test_oversize_rejected_before_any_inference(self):
        self.p['claims'][0]['text']='甲'*30000
        with patch.object(local_review_runtime,'chat') as model:
            with self.assertRaisesRegex(ValueError,'超过'):review.start(self.p,'local','中立核对',app)
        model.assert_not_called()

    def test_unknown_capacity_does_not_infer_or_silently_split(self):
        self.config.pop('apiReviewContextTokens')
        with patch.object(local_review_runtime,'chat') as model:
            with self.assertRaisesRegex(ValueError,'上下文容量'):review.start(self.p,'api','中立核对',app)
        model.assert_not_called()

    def test_followups_resend_entire_original_and_finish_only_last_pass(self):
        review.start(self.p,'api','中立核对',app,2)
        with patch.object(review,'call_model',return_value=response([issue()])) as model:
            for i in range(3):
                review.step(self.p,'full',app,i)
                self.assertEqual(self.p['contractReview']['status']=='complete',i==2)
        self.assertEqual(model.call_count,3)
        for call in model.call_args_list:
            self.assertIn('验收后付款。交付后立即付款。',call[0][0])
        self.assertEqual(self.p['contractReview']['passCompleted']['full'],[0,1,2])
        self.assertEqual(len(self.p['contractReview']['issues']),1)

    def test_cannot_skip_or_exceed_rounds(self):
        review.start(self.p,'api','中立核对',app,2)
        with self.assertRaises(ValueError):review.step(self.p,'full',app,1)
        with self.assertRaises(ValueError):review.step(self.p,'full',app,3)

    def test_retry_skips_completed_round_and_keeps_findings(self):
        review.start(self.p,'api','中立核对',app,1)
        with patch.object(review,'call_model',return_value=response([issue()])):review.step(self.p,'full',app)
        with patch.object(review,'call_model',side_effect=TimeoutError()):review.step(self.p,'full',app,1)
        self.assertIn('full',self.p['contractReview']['failed'])
        with patch.object(review,'call_model',return_value=response()) as model:
            review.step(self.p,'full',app);review.step(self.p,'full',app,1)
        self.assertEqual(model.call_count,1)
        self.assertEqual(len(self.p['contractReview']['issues']),1)

    def test_live_item_is_visible_before_completion(self):
        def model(content,r,engine,on_items=None):
            on_items([issue()])
            self.assertEqual(len(r['issues']),1)
            self.assertNotEqual(r['status'],'complete')
            return response([issue()])
        with patch.object(review,'call_model',side_effect=model):review.step(self.p,'full',app)
        self.assertEqual(len(self.p['contractReview']['issues']),1)

    def test_partial_keeps_items_but_never_claims_complete(self):
        v=response([issue()]);v['_partial']='截断'
        with patch.object(review,'call_model',return_value=v):review.step(self.p,'full',app)
        self.assertEqual(len(self.p['contractReview']['issues']),1)
        self.assertEqual(self.p['contractReview']['status'],'partial')
        self.assertFalse(self.p['contractReview']['completed'])

    def test_old_segmented_task_requires_restart(self):
        self.p['contractReview'].pop('plan')
        with self.assertRaisesRegex(ValueError,'旧任务'):review.step(self.p,0,app)

    def test_local_thinking_disabled_and_no_summary_generation_required(self):
        review.start(self.p,'local','乙方',app)
        def model(url,payload,headers,review_id,**kwargs):
            self.assertFalse(payload['chat_template_kwargs']['enable_thinking'])
            self.assertEqual(payload['reasoning_effort'],'none')
            self.assertIn('全部原文',payload['messages'][0]['content'])
            self.assertNotIn('审核立场：',payload['messages'][0]['content'])
            self.assertIn('不要进行法律评价',payload['messages'][0]['content'])
            return {'choices':[{'finish_reason':'stop','message':{'content':'{"issues":[]}'}}]}
        with patch.object(local_review_runtime,'chat',side_effect=model):review.step(self.p,'full',app)
        self.assertEqual(self.p['contractReview']['status'],'complete')

    def test_api_uses_legal_review_and_local_remains_proofreading(self):
        captured=[]
        def model(url,payload,headers,review_id,**kwargs):
            captured.append(payload)
            return {'choices':[{'finish_reason':'stop','message':{'content':'{"issues":[]}'}}]}
        with patch.object(local_review_runtime,'chat',side_effect=model):
            review.start(self.p,'local','甲方',app,1)
            review.step(self.p,'full',app)
            review.step(self.p,'full',app,1)
            review.start(self.p,'api','乙方',app,1)
            review.step(self.p,'full',app)
            review.step(self.p,'full',app,1)
        self.assertIn('不要进行法律评价',captured[0]['messages'][0]['content'])
        self.assertIn('法律疑点',captured[2]['messages'][0]['content'])
        self.assertIn('不得声称已检索法条',captured[2]['messages'][0]['content'])
        self.assertNotIn('审核立场：',captured[2]['messages'][0]['content'])
        self.assertNotIn('法律疑点',captured[1]['messages'][1]['content'])
        self.assertIn('法律疑点',captured[3]['messages'][1]['content'])
        self.assertEqual(self.p['contractReview']['instructionProfile'],'legal-v1')

    def test_old_api_review_does_not_mix_new_scope_with_old_results(self):
        self.p['contractReview'].pop('instructionProfile')
        self.p['contractReview']['issues']=[{'id':'preserved'}]
        with patch.object(review,'call_model') as model:
            with self.assertRaisesRegex(ValueError,'旧范围'):review.step(self.p,'full',app)
        model.assert_not_called()
        self.assertEqual(self.p['contractReview']['issues'][0]['id'],'preserved')

    def test_missing_outer_json_brace_is_repaired_without_changing_issue(self):
        raw=issue()
        content=json.dumps({'issues':[raw]},ensure_ascii=False)[:-1]
        reply={'choices':[{'finish_reason':'stop','message':{'content':content}}]}
        with patch.object(local_review_runtime,'chat',return_value=reply):review.step(self.p,'full',app)
        self.assertEqual(self.p['contractReview']['status'],'complete')
        self.assertEqual(self.p['contractReview']['issues'][0]['reason'],raw['reason'])

    def test_missing_inner_json_does_not_mark_completed(self):
        reply={'choices':[{'finish_reason':'stop','message':{'content':'{"issues":[{"c":"text"'} }]}
        with patch.object(local_review_runtime,'chat',return_value=reply):review.step(self.p,'full',app)
        self.assertEqual(self.p['contractReview']['status'],'partial')

    def test_unlocated_is_retained_but_cannot_apply(self):
        with patch.object(review,'call_model',return_value=response([issue('不存在的文字'),issue()])):review.step(self.p,'full',app)
        self.assertEqual(len(self.p['contractReview']['unlocated']),1)
        self.assertEqual(len(self.p['contractReview']['issues']),1)
        with self.assertRaises(ValueError):review.decide(self.p,'unlocated','apply',app)

    def test_whitespace_alignment_does_not_rewrite_source(self):
        self.assertEqual(review.align_whitespace('合同总价100000元','合同总价 100000 元'),'合同总价100000元')
        self.assertIsNone(review.align_whitespace('甲方付款，甲方付款','甲方 付款'))

    def test_apply_reuses_existing_revision_path(self):
        with patch.object(review,'call_model',return_value=response([issue()])):review.step(self.p,'full',app)
        review.decide(self.p,self.p['contractReview']['issues'][0]['id'],'apply',app)
        self.assertIn('账户',self.p['claims'][0]['text'])
        self.assertEqual(self.p['contractReview']['status'],'stale')
        self.assertTrue(app.word_revision_groups(self.p))

    def test_changed_document_or_model_invalidates_review(self):
        self.p['claims'][0]['text']+='修改'
        with self.assertRaises(ValueError):review.step(self.p,'full',app)
        self.config['apiModel']='changed'
        with self.assertRaises(ValueError):review.call_model('全文',self.p['contractReview'],app)

    def test_same_anchor_preserves_different_reasons_and_disables_ambiguous_replace(self):
        a,_=review.validate_issues(response([issue()]),self.p)
        review.merge_findings(self.p['contractReview'],a,0)
        raw=issue();raw.update(reason='另一理由',replacement='另一改法')
        b,_=review.validate_issues(response([raw]),self.p)
        review.merge_findings(self.p['contractReview'],b,1)
        self.assertEqual(len(a[0]['variants']),1);self.assertIsNone(a[0]['replacement'])

    def test_short_followup_memory_does_not_truncate_original(self):
        a,_=review.validate_issues(response([issue()]),self.p)
        self.p['contractReview']['issues']=a*100
        text=review.followup_content('完整原文'*1000,self.p['contractReview'],{'C001'},2)
        self.assertTrue(text.startswith('完整原文'*1000))
        self.assertLess(len(text)-4000,2000)

    def test_uses_loaded_capacity_not_advertised_maximum(self):
        self.config.pop('localReviewContextTokens')
        model={'key':'test-small','max_context_length':262144,'loaded_instances':[{'config':{'context_length':4096}}]}
        with patch.object(full_document_context.urllib.request,'urlopen',return_value=io.BytesIO(json.dumps({'models':[model]}).encode())):
            limit,source=full_document_context.capacity(self.config,'local')
        self.assertEqual(limit,4096)
        self.assertIn('已加载',source)

    def test_unloaded_model_maximum_is_not_used_as_current_capacity(self):
        self.config.pop('localReviewContextTokens')
        model={'key':'test-small','max_context_length':262144,'loaded_instances':[]}
        with patch.object(full_document_context.urllib.request,'urlopen',return_value=io.BytesIO(json.dumps({'models':[model]}).encode())):
            with self.assertRaises(ValueError):full_document_context.capacity(self.config,'local')


class HandlerIntegrationTests(unittest.TestCase):
    def test_real_http_model_transport_and_legacy_handlers(self):
        captured = []
        class Model(BaseHTTPRequestHandler):
            def log_message(self, *_): pass
            def do_POST(self):
                data = json.loads(self.rfile.read(int(self.headers['Content-Length'])))
                captured.append(data)
                content = json.dumps(response([issue()]), ensure_ascii=False)
                body = json.dumps({'choices': [{'finish_reason': 'stop', 'message': {'content': content}}]}).encode()
                self.send_response(200); self.send_header('Content-Length', str(len(body))); self.end_headers(); self.wfile.write(body)
        with tempfile.TemporaryDirectory() as folder:
            model = ThreadingHTTPServer(('127.0.0.1', 0), Model)
            server = ThreadingHTTPServer(('127.0.0.1', 0), contract_app.ContractHandler)
            for item in (model, server): threading.Thread(target=item.serve_forever, daemon=True).start()
            config = {'apiBase': f'http://127.0.0.1:{model.server_port}/v1', 'apiModel': 'fixture','apiReviewContextTokens':32768}
            try:
                with patch.object(app, 'PROJECTS', Path(folder)), patch.object(app, 'load_config', return_value=config), patch.object(app,'web_log'):
                    base = f'http://127.0.0.1:{server.server_port}'
                    def post(path, data):
                        request = urllib.request.Request(base + path, json.dumps(data).encode(), {'Content-Type': 'application/json'})
                        return json.loads(urllib.request.urlopen(request).read())
                    page = urllib.request.urlopen(base).read().decode()
                    self.assertIn('/contract-review.js', page)
                    self.assertIn('<title>审核溯源台</title>', page)
                    self.assertIn('/audit-theme.css', page)
                    self.assertIn('pdf-reader-dialog', page)
                    self.assertEqual(post('/api/contract-review/cancel',{'reviewId':'test-no-active-request'}),{'ok':True})
                    progress=json.loads(urllib.request.urlopen(base+'/api/contract-review/progress?id=test-no-active-request').read())
                    self.assertEqual(progress['stage'],'idle')
                    p = post('/api/contract-review/start', {'project': project(), 'provider': 'api'})
                    p = post('/api/contract-review/step', {'project': p, 'index': 0})
                    self.assertEqual(p['contractReview']['completed'], [0]); self.assertEqual(len(captured), 1)
                    self.assertEqual(p['contractReview']['status'], 'complete')
                    result = post('/api/update-claim', {'project': p, 'claimId': 'C001', 'text': '人工修改后的合同。'})
                    self.assertEqual(result['project']['claims'][0]['revisionStatus'], 'pending')
                    self.assertIn('contractReview', json.loads(urllib.request.urlopen(base + '/api/project/fictional-review-test').read()))
            finally:
                for item in (server, model): item.shutdown(); item.server_close()


class LocalStreamingTests(unittest.TestCase):
    def stream(self, events):
        class Response(io.BytesIO):
            headers={'Content-Type':'text/event-stream'}
        return Response(b''.join(b'data: '+json.dumps(event).encode()+b'\n\n' for event in events)+b'data: [DONE]\n\n')

    def test_stream_reply_and_progress(self):
        events=[{'choices':[{'delta':{'content':'{"issues":[]}'},'finish_reason':None}]},
                {'choices':[{'delta':{},'finish_reason':'stop'}],'usage':{'completion_tokens':8}}]
        with patch.object(local_review_runtime.urllib.request,'urlopen',return_value=self.stream(events)) as transport:
            v=local_review_runtime.chat('http://fixture/v1/chat/completions',{}, {},'stream-test')
        self.assertEqual(v['choices'][0]['finish_reason'],'stop')
        self.assertEqual(local_review_runtime.progress('stream-test')['stage'],'done')
        sent=json.loads(transport.call_args[0][0].data)
        self.assertTrue(sent['stream'])

    def test_complete_opinion_published_before_final_stop(self):
        seen=[]
        class Response:
            headers={'Content-Type':'text/event-stream'}
            def __enter__(self):return self
            def __exit__(self,*args):pass
            def __iter__(self):
                first={'choices':[{'delta':{'content':'{"issues":[{"c":"text","p":"C001","q":"帐户","w":"核对","s":"统一"}]'},'finish_reason':None}]}
                yield b'data: '+json.dumps(first).encode()+b'\n'
                seen.append(local_review_runtime.progress('live-test').get('issues'))
                yield b'data: '+json.dumps({'choices':[{'delta':{'content':'}'},'finish_reason':'stop'}]}).encode()+b'\n'
        with patch.object(local_review_runtime.urllib.request,'urlopen',return_value=Response()):
            local_review_runtime.chat('http://fixture',{}, {},'live-test',on_items=lambda rows:{'issues':rows})
        self.assertEqual(seen[0][0]['q'],'帐户')

    def test_thinking_is_stopped_not_waited_out(self):
        events=[{'choices':[{'delta':{'reasoning_content':'internal reasoning'}}]}]
        with patch.object(local_review_runtime.urllib.request,'urlopen',return_value=self.stream(events)):
            with self.assertRaisesRegex(ValueError,'思考'):
                local_review_runtime.chat('http://fixture',{}, {},'thinking-test')

    def test_incomplete_json_salvages_only_complete_objects(self):
        text='{"issues":[{"c":"text","p":"C001","q":"帐户","w":"待核对用词","s":"统一用词"},{"c":"logic","q":"未完成'
        rows=local_review_runtime.complete_items(text)
        self.assertEqual(len(rows),1)
        self.assertEqual(rows[0]['q'],'帐户')

    def test_cancel_preserves_complete_findings(self):
        class Response:
            headers={'Content-Type':'text/event-stream'}
            def __enter__(self): return self
            def __exit__(self,*args): pass
            def __iter__(self):
                first={'choices':[{'delta':{'content':'{"issues":[{"c":"text","p":"C001","q":"帐户","w":"核对用词","s":"统一"}]'}}]}
                yield b'data: '+json.dumps(first).encode()+b'\n'
                local_review_runtime.cancel('cancel-test')
                yield b'data: {}\n'
        with patch.object(local_review_runtime.urllib.request,'urlopen',return_value=Response()):
            value=local_review_runtime.chat('http://fixture',{}, {},'cancel-test')
        self.assertEqual(value['choices'][0]['finish_reason'],'partial')
        self.assertIn('停止',value['_partial_error'])

    def test_deadline_stops_stream(self):
        with patch.object(local_review_runtime.urllib.request,'urlopen',return_value=self.stream([{'choices':[]}])):
            with self.assertRaises(TimeoutError): local_review_runtime.chat('http://fixture',{}, {},'deadline-test',deadline=0)


if __name__ == '__main__': unittest.main()
