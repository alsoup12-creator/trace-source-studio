"""审核溯源台 entry point. Existing API/MCP handlers are delegated intact."""
import urllib.parse
import threading
import sys
import webbrowser
import app as studio
import contract_review as review
import local_review_runtime


class ContractHandler(studio.Handler):
    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == '/api/contract-review/settings':
            config=studio.load_config()
            return self.send_json(200, {key:int(config.get(key) or 0) for key in ('localReviewContextTokens','apiReviewContextTokens')})
        if parsed.path == '/api/contract-review/progress':
            review_id = (urllib.parse.parse_qs(parsed.query).get('id') or [''])[0]
            return self.send_json(200, local_review_runtime.progress(review_id))
        if urllib.parse.urlparse(self.path).path == '/audit-mark.svg':
            body = (studio.WEB / 'audit-mark.svg').read_bytes()
            self.send_response(200)
            self.send_header('Content-Type', 'image/svg+xml')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers(); self.wfile.write(body)
            return
        if urllib.parse.urlparse(self.path).path in ('/', '/index.html'):
            page = (studio.WEB / 'index.html').read_text(encoding='utf-8')
            page = page.replace('<title>溯源台</title>', '<title>审核溯源台</title>')
            page = page.replace('本地模型（可选智能细分与检索辅助）', '本地审核模型')
            page = page.replace('在线观点结构化模型', 'API 审核模型')
            page = page.replace('</head>', '<link rel="icon" href="/audit-mark.svg" type="image/svg+xml"><link rel="stylesheet" href="/contract-review.css"><link rel="stylesheet" href="/audit-theme.css"></head>')
            page = page.replace('</body>', '<script src="/contract-review.js"></script></body>')
            body = page.encode('utf-8')
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.send_header('Content-Length', str(len(body)))
            self.send_header('Cache-Control', 'no-store')
            self.end_headers(); self.wfile.write(body)
            return
        return super().do_GET()

    def do_POST(self):
        path = urllib.parse.urlparse(self.path).path
        if not path.startswith('/api/contract-review/'):
            return super().do_POST()
        try:
            data = self.body()
            if path.endswith('/settings'):
                config=studio.read_json(studio.CONFIG,{})
                for key in ('localReviewContextTokens','apiReviewContextTokens'):
                    value=int(data.get(key) or 0)
                    if not 0<=value<=2097152:
                        raise ValueError('上下文容量应为0至2097152之间的token数。')
                    config[key]=value
                studio.write_json(studio.CONFIG,config)
                return self.send_json(200,{'ok':True})
            if path.endswith('/cancel'):
                local_review_runtime.cancel(str(data.get('reviewId') or ''))
                return self.send_json(200, {'ok': True})
            project = data['project']
            action = path.rsplit('/', 1)[-1]
            if action == 'start':
                review.start(project, data.get('provider'), data.get('stance'), studio, data.get('extraChecks', 0))
            elif action == 'step':
                review.step(project, data.get('index'), studio, data.get('passIndex', 0))
            elif action == 'decide':
                review.decide(project, data.get('issueId'), data.get('decision'), studio)
            elif action == 'relocate':
                review.relocate_review(project)
            else:
                return self.send_json(404, {'error': 'not found'})
            project['updatedAt'] = studio.now()
            studio.write_json(studio.project_path(project['id']), project)
            return self.send_json(200, project)
        except Exception as error:
            return self.send_json(400, {'error': studio.redact_sensitive(str(error))})


studio.Handler = ContractHandler


def main():
    studio.PROJECTS.mkdir(parents=True, exist_ok=True)
    server, url = studio.start_server()
    studio.APP = studio.Api()
    try:
        import webview
        webview.create_window('审核溯源台', url, js_api=studio.APP, width=1460, height=920, min_size=(1180, 760))
        webview.start(debug='--debug' in sys.argv)
    except ImportError:
        webbrowser.open(url, new=2)
        print('审核溯源台已在默认浏览器打开。')
        try:
            threading.Event().wait()
        except KeyboardInterrupt:
            pass
    finally:
        server.shutdown()


if __name__ == '__main__':
    main()
