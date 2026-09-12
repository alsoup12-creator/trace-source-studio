#!/usr/bin/env python3
import importlib.util, json, re, tempfile, zipfile
from pathlib import Path
from pypdf import PdfWriter

ROOT=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location("trace_source_app",ROOT/"app.py"); app=importlib.util.module_from_spec(spec); spec.loader.exec_module(app)

def make_docx(path,text):
    xml=f'''<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:body></w:document>'''
    with zipfile.ZipFile(path,"w") as archive:archive.writestr("word/document.xml",xml)

def make_pptx(path,text):
    xml=f'''<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>'''
    with zipfile.ZipFile(path,"w") as archive:archive.writestr("ppt/slides/slide1.xml",xml)

def make_xlsx(path):
    workbook='''<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="车辆数据" sheetId="1" r:id="rId1"/></sheets></workbook>'''
    rels='''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'''
    strings='''<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>品牌</t></si><si><t>金额</t></si><si><t>丰田</t></si></sst>'''
    sheet='''<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row><row r="2"><c r="A2" t="s"><v>2</v></c><c r="B2"><v>100</v></c><c r="C2"><f>SUM(B2:B2)</f><v>100</v></c></row></sheetData></worksheet>'''
    with zipfile.ZipFile(path,"w") as archive:
        archive.writestr("xl/workbook.xml",workbook);archive.writestr("xl/_rels/workbook.xml.rels",rels);archive.writestr("xl/sharedStrings.xml",strings);archive.writestr("xl/worksheets/sheet1.xml",sheet)

def main():
    with tempfile.TemporaryDirectory() as folder:
        root=Path(folder); word=root/"result.docx"; ppt=root/"result.pptx"; excel=root/"source.xlsx"; source=root/"source.txt"
        app.DATA=root/"data"; app.PROJECTS=app.DATA/"projects"; app.CONFIG=app.DATA/"config.json"; app.LOG=app.DATA/"trace-source-studio.log"
        make_docx(word,"合同金额为100万元，应在验收后十日内支付。");make_pptx(ppt,"项目收入增长20%。");make_xlsx(excel);source.write_text("合同约定金额为100万元，付款期限为验收完成后十日。",encoding="utf-8")
        word_artifact=app.parse_artifact(str(word));ppt_artifact=app.parse_artifact(str(ppt));excel_artifact=app.parse_artifact(str(excel))
        assert word_artifact["kind"]=="word" and ppt_artifact["kind"]=="ppt" and excel_artifact["kind"]=="excel"
        assert excel_artifact["blocks"][1]["excel"]=={"sheet":"车辆数据","row":"2","cells":[{"ref":"A2","column":"A","value":"丰田"},{"ref":"B2","column":"B","value":"100"},{"ref":"C2","column":"C","value":"100；公式=SUM(B2:B2)"}],"continuation":False}
        excel_rows=app.file_blocks(excel);assert excel_rows and "工作表“车辆数据”!A2:C2"==excel_rows[1]["location"] and "公式=SUM(B2:B2)" in excel_rows[1]["text"]
        image=root/"scan.png";image.write_bytes(b"test-image")
        original_image_blocks=app.image_blocks
        try:
            app.image_blocks=lambda path:[{"id":"image-ocr-1-1","location":"图片OCR第1行","text":"合同金额为100万元","sourceType":"图片 OCR（CUDA）"}]
            image_artifact=app.parse_artifact(str(image))
            assert image_artifact["kind"]=="image" and image_artifact["blocks"][0]["text"]=="合同金额为100万元"
        finally:
            app.image_blocks=original_image_blocks
        image.unlink()
        claims=app.split_claims(word_artifact["blocks"]);assert claims and claims[0]["blockId"]=="p1"
        assert word_artifact["sourceSnapshot"]["sha256"]==app.file_sha256(word)
        assert word_artifact["blocks"][0]["paragraphIndex"]==1 and claims[0]["sourceStart"]==0 and claims[0]["sourceEnd"]==len(claims[0]["text"])
        revision_project={"artifact":word_artifact,"claims":json.loads(json.dumps(claims,ensure_ascii=False))}
        revision_claim=revision_project["claims"][0];revision_claim["candidates"]=[{"label":"修改前依据"}];revision_claim["questions"]=[{"role":"user","content":"旧问题"}]
        app.update_claim_text(revision_project,revision_claim["id"],"合同总金额为100万元，应在验收后十日内支付。")
        assert revision_claim["revisionStatus"]=="pending" and revision_claim["status"]=="待重新溯源" and not revision_claim["candidates"]
        assert revision_claim["evidenceHistory"][0]["candidates"]==[{"label":"修改前依据"}]
        revision_groups=app.word_revision_groups(revision_project)
        assert len(revision_groups)==1 and revision_groups[0]["paragraphIndex"]==1 and revision_groups[0]["operations"][0]["original"].startswith("合同金额")
        assert app.character_revision_edits("甲方应在十日内付款。","甲方应在十五日内付款。") == [{"start":5,"end":5,"original":"","replacement":"五"}]
        class FakeCount:
            def __init__(self,value=0): self.Count=value
        class FakeRange:
            def __init__(self,document,start=0,end=0,paragraph=False): self.document=document;self.Start=start;self.end=end;self.paragraph=paragraph
            @property
            def Duplicate(self): return self
            @property
            def Text(self): return self.document.text+"\r" if self.paragraph else self.document.text[self.Start:self.end]
            @Text.setter
            def Text(self,value):
                self.document.writes.append((self.Start,self.end,value))
                self.document.text=self.document.text[:self.Start]+value+self.document.text[self.end:];self.document.Revisions.Count+=2
        class FakeParagraphs:
            Count=1
            def __init__(self,document): self.document=document
            def Item(self,index): return type("Paragraph",(),{"Range":FakeRange(self.document,paragraph=True)})()
        class FakeDocument:
            def __init__(self,text): self.text=text;self.writes=[];self.Revisions=FakeCount();self.Paragraphs=FakeParagraphs(self);self.TrackRevisions=False
            def Range(self,start,end): return FakeRange(self,start,end)
        fake_document=FakeDocument(word_artifact["blocks"][0]["text"])
        assert app.apply_word_revision_groups(fake_document,revision_groups)==2 and fake_document.TrackRevisions is True
        assert fake_document.text.startswith("合同总金额为100万元")
        assert fake_document.writes==[(2,2,"总")]
        app.mark_claim_retraced(revision_claim,[{"label":"新依据"}]);assert revision_claim["revisionStatus"]=="retraced"
        try:
            app.update_claim_text(revision_project,revision_claim["id"],"不允许\n新增段落")
            raise AssertionError("newline revision should fail")
        except ValueError as error:
            assert "不支持" in str(error)
        assert len(app.split_claims([{"id":f"p{index}","location":f"段落{index}","text":f"第{index}段"} for index in range(1,131)]))==130
        docs=app.scan_folder(str(root));rows=app.local_candidates(claims[0]["text"],docs)
        assert any(item["fileName"]=="source.xlsx" and item["blocks"][1]["sourceType"]=="Excel 工作表" for item in docs)
        assert rows and rows[0]["path"]
        first_index=app.build_local_index(str(root))
        indexed_rows=app.indexed_local_candidates(claims[0]["text"],first_index)
        second_index=app.build_local_index(str(root))
        assert indexed_rows and indexed_rows[0]["method"]=="本地索引关键词召回"
        assert second_index["rebuilt"]==0 and second_index["reused"]>=3
        source.write_text("合同金额为100万元，付款期限调整为验收完成后十五日。",encoding="utf-8")
        changed_index=app.build_local_index(str(root))
        assert changed_index["rebuilt"]==1 and changed_index["reused"]>=2
        diverse_docs=[
            {"path":"a.txt","fileName":"a.txt","blocks":[{"id":"1","location":"1","text":"共同还款责任应由共同还款人承担。"},{"id":"2","location":"2","text":"共同还款人承担共同还款责任。"},{"id":"3","location":"3","text":"共同还款义务属于合同责任。"}]},
            {"path":"b.txt","fileName":"b.txt","blocks":[{"id":"1","location":"1","text":"当事人约定共同承担还款义务。"}]},
        ]
        diverse=app.local_candidates("共同还款人承担还款责任",diverse_docs,limit=3)
        assert len({item["path"] for item in diverse})==2 and all(item.get("matchType") for item in diverse)
        semantic=app.local_candidates("第三方向债权人作出承诺",[{"path":"c.txt","fileName":"c.txt","blocks":[{"id":"1","location":"1","text":"民法典规定债务加入的构成要件。"}]}],extra_queries=["债务加入"])
        assert semantic[0]["matchType"]=="观点近似"
        single=app.trace_claim_candidates(claims[0],"",{"local":True,"web":False,"mcp":False},{},docs)
        assert single and single[0]["path"]
        original_plan=app.plan_claim_query
        try:
            app.plan_claim_query=lambda *args,**kwargs: (_ for _ in ()).throw(AssertionError("local-only mode must skip online understanding"))
            local_only_claim={"text":claims[0]["text"],"selected":True,"status":"待查","candidates":[]}
            local_only=app.trace_claim_candidates(local_only_claim,"",{"local":True,"localAi":False,"understanding":True,"web":False,"mcp":False,"pkulaw":False},{},docs)
            assert local_only and "tracePlan" not in local_only_claim
        finally:
            app.plan_claim_query=original_plan
        assert app.relation_label("收入增长20%","收入下降20%") in {"同向候选","反向候选"}
        assert app.parse_json_object('```json\n{"coreIssue":"共同还款责任"}\n```')["coreIssue"]=="共同还款责任"
        original_chat=app.metaso_chat
        try:
            app.metaso_chat=lambda messages,config:'{"coreIssue":"共同还款人的责任性质","reviewQuestion":"共同还款约定产生何种法律关系","viewpointLabels":["共同债务","保证担保","债务加入"],"subclaims":["共同承担还款责任"],"webQuery":"共同还款人 法院认为","pkulawKeywordQuery":"共同还款人 共同债务人","pkulawSemanticQuery":"共同还款人的法律性质是什么"}'
            plan=app.plan_claim_query("共同还款人共同承担还款责任","",{})
            assert plan["webQuery"]=="共同还款人 法院认为" and plan["viewpointLabels"]==["共同债务","保证担保","债务加入"]
            app.metaso_chat=lambda messages,config:'{"sourceType":"司法案例","coreView":"法院认定共同承担责任。","viewpointLabels":[{"label":"共同债务","summary":"法院如此认定","quote":"应承担共同还款责任"},{"label":"保证担保","summary":"没有原文","quote":"正文不存在"}],"keyQuotes":[{"role":"法院认为","quote":"应承担共同还款责任"},{"role":"无效引用","quote":"正文不存在"}],"alignment":[{"subclaim":"共同承担还款责任","status":"明确涉及","explanation":"来源明确表述","quote":"应承担共同还款责任"}]}'
            reading=app.deep_read_candidate({"text":"共同承担还款责任","tracePlan":plan},{"rawText":"案件事实部分用于说明合同签订经过。原审法院认为，应承担共同还款责任。该判断属于法院对责任性质的具体认定。"},"",{})
            assert len(reading["keyQuotes"])==1 and reading["viewpointLabels"]==[{"label":"共同债务","summary":"法院如此认定","quote":"应承担共同还款责任"}]
            app.metaso_chat=lambda messages,config:'{"overallView":"来源主要从共同债务角度表述。","groups":[{"label":"共同债务","summary":"来源明确提及共同承担责任。","sourceIndexes":[0],"keyQuotes":[{"sourceIndex":0,"quote":"法院认为，应承担共同还款责任"}]}],"gaps":["未说明全部责任条件"]}'
            summary=app.summarize_claim_sources({"text":"共同承担还款责任"},[{"label":"案例网页","url":"https://example.com","method":"秘塔 Search API","quote":"法院认为，应承担共同还款责任"}],"",{})
            assert summary["status"]=="ok" and summary["groups"][0]["sourceIndexes"]==[0]
        finally:
            app.metaso_chat=original_chat
        original_post_json=app.post_json
        try:
            local_calls=[]
            def fake_local_post(url,payload,headers=None,timeout=30):
                local_calls.append((url,payload,headers))
                return {"choices":[{"message":{"content":'{"groups":[{"unitIds":["p1-u1"],"shortView":"合同金额","type":"数据"},{"unitIds":["p1-u2"],"shortView":"付款期限","type":"合同约定"},{"unitIds":["不存在-u9"],"shortView":"无效","type":"事实"}]}'}}]}
            app.post_json=fake_local_post
            local_claims,structure_stats=app.local_semantic_structure(word_artifact["blocks"],"重点拆分金额",{"localApiBase":"http://127.0.0.1:1234/v1","localApiModel":"test-4b","localApiKey":""},return_stats=True)
            assert len(local_claims)==2 and local_claims[0]["text"]=="合同金额为100万元，" and local_claims[1]["text"]=="应在验收后十日内支付。"
            assert local_claims[0]["claimType"]=="数据" and local_claims[0]["structureSource"]=="local-model" and local_calls[0][2]=={}
            assert local_claims[0]["originalText"]==local_claims[0]["text"] and local_claims[0]["sourceStart"]==0 and local_claims[0]["sourceEnd"]==11
            assert structure_stats["modelAcceptedGroups"]==2 and structure_stats["invalidGroups"]==1 and structure_stats["fallbackRate"]==0
            assert local_calls[0][1]["reasoning_effort"]=="none" and local_calls[0][1]["chat_template_kwargs"]["enable_thinking"] is False
            assert "只选择和组合编号" in local_calls[0][1]["messages"][0]["content"]
        finally:
            app.post_json=original_post_json
        original_post_json=app.post_json
        try:
            app.post_json=lambda *args,**kwargs:{"choices":[{"message":{"content":'{"groups":[{"unitIds":["错误编号"],"shortView":"错误","type":"事实"}]}'}}]}
            fallback_claims,fallback_stats=app.local_semantic_structure(word_artifact["blocks"],"",{"localApiBase":"http://127.0.0.1:1234/v1","localApiModel":"test-1b"},return_stats=True)
            assert "".join(item["text"] for item in fallback_claims)==word_artifact["blocks"][0]["text"]
            assert fallback_stats["fallbackRate"]==1 and all(item["structureSource"]=="rule-fallback" for item in fallback_claims)
        finally:
            app.post_json=original_post_json
        original_post_json=app.post_json
        try:
            many_text="".join(f"第{index}项义务。" for index in range(1,46))
            many_blocks=[{"id":"p9","location":"长段落","text":many_text}]
            batch_calls=[]
            def fake_many_post(url,payload,headers=None,timeout=30):
                batch_calls.append(payload)
                ids=re.findall(r"\[([^\]]+-u\d+)\]",payload["messages"][1]["content"])
                groups=[{"unitIds":[value],"shortView":value,"type":"合同约定"} for value in ids]
                return {"choices":[{"message":{"content":json.dumps({"groups":groups},ensure_ascii=False)}}]}
            app.post_json=fake_many_post
            many_claims,many_stats=app.local_semantic_structure(many_blocks,"",{"localApiBase":"http://127.0.0.1:1234/v1","localApiModel":"test-3b"},return_stats=True)
            assert len(batch_calls)==2 and many_stats["batches"]==2 and many_stats["fallbackRate"]==0
            assert "".join(item["text"] for item in many_claims)==many_text
        finally:
            app.post_json=original_post_json
        original_post_json=app.post_json
        try:
            local_calls=[]
            def fake_expand_post(url,payload,headers=None,timeout=30):
                local_calls.append((url,payload,headers,timeout))
                return {"choices":[{"message":{"content":'{"queries":["共同债务","债务加入","共同还款责任"]}'}}]}
            app.post_json=fake_expand_post
            expanded=app.local_ai_expand_query("共同还款人承担还款责任","优先法律术语",{"localApiBase":"http://127.0.0.1:1234/v1","localApiModel":"test-4b","localApiKey":""})
            assert expanded==["共同债务","债务加入","共同还款责任"]
            assert local_calls[0][1]["reasoning_effort"]=="none" and local_calls[0][1]["chat_template_kwargs"]["enable_thinking"] is False
            assert "短的同义词、法律术语或固定表达" in local_calls[0][1]["messages"][0]["content"]
            assert "不要返回自然语言问题或长句" in local_calls[0][1]["messages"][0]["content"]
            assert app.local_candidates(claims[0]["text"],docs,extra_queries=expanded)[0]["method"]=="本地AI增强关键词召回"
        finally:
            app.post_json=original_post_json
        original_post_json=app.post_json
        try:
            refine_rows=[
                {"path":"exact.txt","url":"","label":"逐字","location":"1","quote":"合同金额为100万元","relation":"同向候选","matchType":"逐字命中"},
                {"path":"conflict.txt","url":"","label":"冲突","location":"1","quote":"合同金额为200万元","relation":"反向候选","matchType":"关键词组合"},
                {"path":"upgrade.txt","url":"","label":"近义","location":"1","quote":"价款合计壹佰万元","relation":"同向候选","matchType":"观点近似"},
                {"path":"noise.txt","url":"","label":"噪声","location":"1","quote":"其他项目合同金额统计","relation":"同向候选","matchType":"关键词组合"},
            ]
            app.apply_rule_candidate_tiers(refine_rows)
            app.post_json=lambda *args,**kwargs:{"choices":[{"message":{"content":'{"results":[{"candidateId":"L1","label":"unrelated","reason":"模型误判逐字项"},{"candidateId":"L2","label":"direct","reason":"模型误判冲突项"},{"candidateId":"L3","label":"direct","reason":"语义直接对应"},{"candidateId":"L4","label":"unrelated","reason":"只涉及统计主题"}]}'}}]}
            refine_stats=app.local_ai_refine_candidates("合同金额为100万元",refine_rows,{"localApiBase":"http://127.0.0.1:1234/v1","localApiModel":"test-3b"})
            assert [row["reviewTier"] for row in refine_rows]==["primary","conflict","primary","weak"]
            assert len(refine_rows)==4 and refine_stats["classified"]==4 and refine_stats["counts"]=={"primary":2,"conflict":1,"weak":1}
            assert refine_rows[0]["reviewTierSource"]=="rule-guarded" and refine_rows[0]["aiLabel"]=="硬规则保留"
            assert refine_rows[2]["aiReason"]=="语义直接对应"
        finally:
            app.post_json=original_post_json
        original_expand=app.local_ai_expand_query
        try:
            expand_calls=[]
            app.local_ai_expand_query=lambda claim,prompt,config: (expand_calls.append(claim), ["付款期限"])[1]
            probe_claim={"text":"合同金额为100万元", "selected":True, "status":"待查", "candidates":[]}
            probe_docs=[
                {"path":"rule.txt", "fileName":"rule.txt", "blocks":[{"location":"文本片段1", "text":"合同金额为100万元。"}]},
                {"path":"supplement.txt", "fileName":"supplement.txt", "blocks":[{"location":"文本片段1", "text":"付款期限为验收后十日。"}]},
            ]
            merged=app.trace_claim_candidates(probe_claim,"",{"local":True,"localAi":True,"web":False,"mcp":False,"pkulaw":False},{},probe_docs)
            assert expand_calls==["合同金额为100万元"] and len(merged)==2
            assert merged[0]["method"]=="本地关键词召回" and merged[1]["method"]=="本地AI增强关键词召回"
        finally:
            app.local_ai_expand_query=original_expand
        fragment=app.text_fragment_url("https://example.com/page#section","前文<mark>《民法典》第五百五十二条规定债务加入</mark>后文")
        assert "#section:~:text=" in fragment and "%E6%B0%91%E6%B3%95%E5%85%B8" in fragment
        original_urlopen=app.urllib.request.urlopen; original_log=app.web_log
        try:
            captured={}
            class FakeResponse:
                status=200
                def __enter__(self): return self
                def __exit__(self,*args): pass
                def read(self): return b'{"webpages":[{"title":"Metaso result","link":"https://example.com/source","snippet":"Matched text","score":"high"}]}'
            def fake_urlopen(request,timeout=30):
                captured["payload"]=app.json.loads(request.data.decode("utf-8")); return FakeResponse()
            app.urllib.request.urlopen=fake_urlopen; app.web_log=lambda message:None
            provider=app.MetasoSearchProvider("test-key",size=3,concise_snippet=True,include_summary=True,include_raw_content=True)
            metaso=provider.search("合同金额为100万元")
            assert metaso[0]["method"]=="秘塔 Search API" and metaso[0]["url"]=="https://example.com/source"
            assert captured["payload"]["conciseSnippet"] is True
            assert captured["payload"]["includeSummary"] is True
            assert captured["payload"]["includeRawContent"] is True
            assert "page" not in captured["payload"]
            class ChatResponse(FakeResponse):
                def read(self): return '{"data":{"answer":"秘塔问答返回的来源概括"}}'.encode("utf-8")
            def fake_chat_urlopen(request,timeout=90):
                captured["chat_payload"]=app.json.loads(request.data.decode("utf-8")); return ChatResponse()
            app.urllib.request.urlopen=fake_chat_urlopen
            original_secret=app.SECRETS.get("metasoApiKey")
            app.SECRETS["metasoApiKey"]="test-key"
            answer=app.metaso_chat([{"role":"system","content":"只返回答案"},{"role":"user","content":"测试问题"}],{"metasoChatModel":"fast"})
            assert answer=="秘塔问答返回的来源概括"
            assert captured["chat_payload"]["q"].endswith("测试问题") and captured["chat_payload"]["format"]=="chat_completions"
            if original_secret is None: app.SECRETS.pop("metasoApiKey", None)
            else: app.SECRETS["metasoApiKey"]=original_secret
            class ErrorResponse(FakeResponse):
                def read(self): return b'{"errCode":401,"errMsg":"invalid credential"}'
            app.urllib.request.urlopen=lambda request,timeout=30:ErrorResponse()
            try:
                provider.search("普通测试查询")
                raise AssertionError("Metaso business errors must not be treated as zero results")
            except ValueError as error:
                assert "errCode=401" in str(error)
            mcp_calls=[]
            class PkulawResponse:
                status=200
                headers={"Mcp-Session-Id":"test-session"}
                def __init__(self,body): self.body=body
                def __enter__(self): return self
                def __exit__(self,*args): pass
                def read(self): return self.body.encode("utf-8")
            def fake_pkulaw_urlopen(request,timeout=45):
                payload=app.json.loads(request.data.decode("utf-8"));mcp_calls.append(payload)
                captured["pkulaw_auth"]=request.headers.get("Authorization")
                if payload["method"]=="initialize": return PkulawResponse('{"jsonrpc":"2.0","id":1,"result":{"protocolVersion":"2024-11-05","capabilities":{}}}')
                if payload["method"]=="notifications/initialized": return PkulawResponse('{}')
                if payload["method"]=="tools/list": return PkulawResponse('{"jsonrpc":"2.0","id":2,"result":{"tools":[{"name":"get_law_list","inputSchema":{"type":"object","properties":{"title":{"type":"string"},"fulltext":{"type":"string"},"Timeliness":{"type":"array"}}}}]}}')
                return PkulawResponse('{"jsonrpc":"2.0","id":3,"result":{"structuredContent":{},"content":[{"type":"text","text":"{\\"Message\\":\\"成功\\",\\"Data\\":[{\\"Title\\":\\"中华人民共和国民法典\\",\\"Url\\":\\"[北大法宝](https://www.pkulaw.com/lar/example.html?way=mcp)\\",\\"DocumentNO\\":\\"主席令第四十五号\\",\\"IssueDepartment\\":[\\"全国人民代表大会\\"],\\"TimelinessDic\\":[\\"现行有效\\"],\\"EffectivenessDic\\":[\\"法律\\"]}]}"}]}}')
            app.urllib.request.urlopen=fake_pkulaw_urlopen
            pkulaw=app.PkulawMcpProvider("Bearer test-token",mode="keyword",size=6).search("合同解除")
            assert mcp_calls[-1]["method"]=="tools/call" and mcp_calls[-1]["params"]["name"]=="get_law_list"
            assert mcp_calls[-1]["params"]["arguments"]=={"fulltext":"合同解除","Timeliness":["现行有效"]}
            assert captured["pkulaw_auth"]=="Bearer test-token"
            assert app.redact_sensitive('token: Bearer 12345678-1234-1234-1234-123456789012')=='token: Bearer ***'
            assert pkulaw[0]["method"].startswith("北大法宝 MCP") and pkulaw[0]["url"].startswith("https://www.pkulaw.com/")
            assert pkulaw[0]["sourceType"]=="法律" and "现行有效" in pkulaw[0]["location"]
        finally:
            app.urllib.request.urlopen=original_urlopen; app.web_log=original_log
        app.DATA=root/"data"; app.PROJECTS=app.DATA/"projects"; app.CONFIG=app.DATA/"config.json"; app.LOG=app.DATA/"trace-source-studio.log"
        server,url=app.start_server()
        try:
            project=app.post_json(url+"/api/import",{"path":str(word)})
            assert project["structureMethod"]=="rule"
            original_structure=app.local_semantic_structure
            try:
                app.local_semantic_structure=lambda blocks,prompt,config,return_stats=False: ([{"id":"C001","blockId":"p1","location":"第1个非空段落","text":"合同金额为100万元，","selected":True,"status":"待查","candidates":[]}],{"model":"test-3b","originalBlocks":1,"claims":1,"averageClaimChars":11,"maxClaimChars":11,"fallbackRate":0})
                structured=app.post_json(url+"/api/structure-local",{"project":project,"prompt":"详细细分"})
                assert structured["structureMethod"]=="local-model" and structured["structureStats"]["model"]=="test-3b"
            finally:
                app.local_semantic_structure=original_structure
            revision_http_project=app.post_json(url+"/api/import",{"path":str(word)})
            revision_http_project["claims"][0]["candidates"]=[{"label":"旧依据"}]
            edited_response=app.post_json(url+"/api/update-claim",{"project":revision_http_project,"claimId":"C001","text":"合同金额为120万元，应在验收后十日内支付。"})
            revision_http_project=edited_response["project"]
            assert revision_http_project["claims"][0]["text"].startswith("合同金额为120万元") and revision_http_project["claims"][0]["revisionStatus"]=="pending"
            original_revision_copy=app.create_tracked_revision_copy
            try:
                fake_copy=root/"result_溯源修订_test.docx"
                app.create_tracked_revision_copy=lambda value:{"path":str(fake_copy),"fileName":fake_copy.name,"changedClaims":1,"changedParagraphs":1,"newRevisions":2}
                exported=app.post_json(url+"/api/export-revision",{"project":revision_http_project})
                assert exported["export"]["changedClaims"]==1 and exported["project"]["lastRevisionExport"]["path"]==str(fake_copy)
            finally:
                app.create_tracked_revision_copy=original_revision_copy
            project["structureMethod"]="rule"
            untouched={"label":"保留的旧候选"}; project["claims"].append({"id":"KEEP","blockId":"p1","location":"测试","text":"不应被本次更新的观点","selected":False,"status":"已找到候选","candidates":[untouched]})
            updated=app.post_json(url+"/api/run-claim",{"project":project,"claimId":project["claims"][0]["id"],"prompt":"","modes":{"local":True,"web":False,"mcp":False},"folder":str(root)})
            assert updated["claims"][0]["candidates"]
            assert next(item for item in updated["claims"] if item["id"]=="KEEP")["candidates"]==[untouched]
            local_pdf=root/"local.pdf"; writer=PdfWriter(); writer.add_blank_page(width=288,height=432); writer.add_blank_page(width=288,height=432)
            with local_pdf.open("wb") as stream: writer.write(stream)
            pdf_path=app.urllib.parse.quote(str(local_pdf),safe="")
            with app.urllib.request.urlopen(url+"/api/pdf-meta?path="+pdf_path) as response:
                meta=app.json.loads(response.read().decode("utf-8"))
            with app.urllib.request.urlopen(url+"/api/pdf-page?path="+pdf_path+"&page=2") as response:
                preview_bytes=response.read()
            assert meta["pageCount"]==2 and preview_bytes.startswith(b"\x89PNG")
        finally:
            server.shutdown(); server.server_close()
        original_startfile=app.os.startfile; opened=[]
        try:
            app.os.startfile=lambda target: opened.append(target)
            opened_message=app.open_target(str(local_pdf),"https://example.com/stale","PDF第2页",2)
            assert "第2页" in opened_message and opened==[str(local_pdf)]
        finally:
            app.os.startfile=original_startfile
    print("PASS: Word/PPT/Excel parsing, claim splitting and local source retrieval are valid.")

if __name__=="__main__":main()
