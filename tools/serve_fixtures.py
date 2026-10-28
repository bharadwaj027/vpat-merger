"""Test helper: serve fixtures/ with CORS so the browser pane can load them into the upload inputs."""
import http.server, functools, sys
class H(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        super().end_headers()
http.server.ThreadingHTTPServer(("127.0.0.1", 8765), functools.partial(H, directory=sys.argv[1])).serve_forever()
