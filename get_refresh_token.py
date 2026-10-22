import urllib.parse
import urllib.request
import json
from http.server import HTTPServer, BaseHTTPRequestHandler
import webbrowser

SCOPES = "https://www.googleapis.com/auth/drive"
REDIRECT_URI = "http://localhost:8080/"

auth_code = None

class OAuthHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        global auth_code
        parsed_url = urllib.parse.urlparse(self.path)
        params = urllib.parse.parse_qs(parsed_url.query)
        if "code" in params:
            auth_code = params["code"][0]
            self.send_response(200)
            self.send_header("Content-type", "text/html")
            self.end_headers()
            self.wfile.write(b"<h1>Success! You can close this window and return to the terminal.</h1>")
        else:
            self.send_response(400)
            self.end_headers()
            self.wfile.write(b"<h1>Failed to capture authorization code.</h1>")

    def log_message(self, format, *args):
        return  # suppress server request logs

def exchange_code_for_tokens(client_id, client_secret, code):
    token_url = "https://oauth2.googleapis.com/token"
    data = urllib.parse.urlencode({
        "client_id": client_id,
        "client_secret": client_secret,
        "code": code,
        "grant_type": "authorization_code",
        "redirect_uri": REDIRECT_URI
    }).encode("utf-8")
    
    req = urllib.request.Request(token_url, data=data, method="POST")
    with urllib.request.urlopen(req) as resp:
        result = json.loads(resp.read().decode("utf-8"))
        return result

def main():
    print("=" * 60)
    print(" Google Drive OAuth Refresh Token Generator")
    print("=" * 60)

    client_id = input("Paste your Client ID: ").strip()
    client_secret = input("Paste your Client Secret: ").strip()

    if not client_id or not client_secret:
        print("[!] Client ID and Secret cannot be empty.")
        return

    # Start local listener FIRST
    server = HTTPServer(('localhost', 8080), OAuthHandler)

    params = {
        "client_id": client_id,
        "redirect_uri": REDIRECT_URI,
        "response_type": "code",
        "scope": SCOPES,
        "access_type": "offline",
        "prompt": "consent"
    }
    auth_url = "https://accounts.google.com/o/oauth2/v2/auth?" + urllib.parse.urlencode(params)

    print("\n[*] Opening your browser to sign in...")
    print(f"If the browser doesn't open automatically, copy this link:\n{auth_url}\n")
    webbrowser.open(auth_url)

    print("[*] Waiting for browser authentication callback on port 8080...")
    while not auth_code:
        server.handle_request()

    tokens = exchange_code_for_tokens(client_id, client_secret, auth_code)
    refresh_token = tokens.get("refresh_token")

    print("\n" + "=" * 60)
    print(" SUCCESS! Copy these into your GitHub Secrets:")
    print("=" * 60)
    print(f"GDRIVE_CLIENT_ID:     {client_id}")
    print(f"GDRIVE_CLIENT_SECRET: {client_secret}")
    print(f"GDRIVE_REFRESH_TOKEN: {refresh_token}")
    print("=" * 60)

if __name__ == "__main__":
    main()
