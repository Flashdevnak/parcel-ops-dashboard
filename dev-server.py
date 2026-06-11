from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer

class NoCacheHandler(SimpleHTTPRequestHandler):
    extensions_map = {
        **SimpleHTTPRequestHandler.extensions_map,
        ".js": "application/javascript",
        ".css": "text/css",
        ".svg": "image/svg+xml",
        ".webmanifest": "application/manifest+json",
        ".ico": "image/x-icon",
    }

    def do_GET(self):
        # Python 3.12 HTTPMessage ไม่มี .pop()
        # ใช้ del แทน เพื่อกัน server ตอบ 304 ตอนทดสอบบนเครื่อง
        for key in ("If-Modified-Since", "If-None-Match"):
            try:
                if key in self.headers:
                    del self.headers[key]
            except Exception:
                pass
        return super().do_GET()

    def do_HEAD(self):
        for key in ("If-Modified-Since", "If-None-Match"):
            try:
                if key in self.headers:
                    del self.headers[key]
            except Exception:
                pass
        return super().do_HEAD()

    def end_headers(self):
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        super().end_headers()

if __name__ == "__main__":
    port = 5173
    print(f"Serving production preview no-cache at http://localhost:{port}")
    print("Open: http://localhost:5173")
    ThreadingHTTPServer(("", port), NoCacheHandler).serve_forever()
