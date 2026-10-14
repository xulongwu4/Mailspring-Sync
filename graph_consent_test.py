#!/usr/bin/env python3
"""Check whether your O365 tenant lets Mailspring's app ID send mail via Microsoft Graph.

Usage: python3 /tmp/graph_consent_test.py [--send]
  --send  also sends a one-line test mail to yourself through Graph /me/sendMail
"""
import base64, hashlib, http.server, json, secrets, sys, urllib.error, urllib.parse, urllib.request, webbrowser

CLIENT_ID = "8787a430-6eee-41e1-b914-681d90d35625"  # Mailspring's O365 app (onboarding-constants.ts)
PORT = 12141
REDIRECT = f"http://localhost:{PORT}/desktop"
SCOPES = "openid offline_access user.read Mail.Send"

b64 = lambda b: base64.urlsafe_b64encode(b).rstrip(b"=").decode()
verifier = b64(secrets.token_bytes(32))
state = b64(secrets.token_bytes(16))
url = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?" + urllib.parse.urlencode({
    "client_id": CLIENT_ID, "redirect_uri": REDIRECT, "response_type": "code", "scope": SCOPES,
    "response_mode": "query", "code_challenge": b64(hashlib.sha256(verifier.encode()).digest()),
    "code_challenge_method": "S256", "state": state, "prompt": "select_account"})

result = {}
class H(http.server.BaseHTTPRequestHandler):
    def do_GET(self):
        q = dict(urllib.parse.parse_qsl(urllib.parse.urlparse(self.path).query))
        if not q: return self.send_response(404) or self.end_headers()
        result.update(q)
        self.send_response(200); self.end_headers()
        self.wfile.write(b"Done - return to the terminal.")
    def log_message(self, *a): pass

print("Open this URL and sign in with your O365 account:\n\n" + url + "\n")
webbrowser.open(url)
srv = http.server.HTTPServer(("127.0.0.1", PORT), H)
while not result: srv.handle_request()

if "error" in result:
    sys.exit(f"FAIL at consent: {result['error']}: {result.get('error_description')}\n"
             "(AADSTS65001 / 'admin approval required' = tenant blocks user consent -> patch won't help)")
assert result.get("state") == state, "state mismatch"

def post(u, data, headers=None):
    req = urllib.request.Request(u, data=data, headers=headers or {}, method="POST")
    try:
        with urllib.request.urlopen(req) as r: return r.status, r.read().decode()
    except urllib.error.HTTPError as e: return e.code, e.read().decode()

st, body = post("https://login.microsoftonline.com/common/oauth2/v2.0/token", urllib.parse.urlencode({
    "client_id": CLIENT_ID, "code": result["code"], "code_verifier": verifier, "grant_type": "authorization_code",
    "redirect_uri": REDIRECT, "scope": SCOPES}).encode())
tok = json.loads(body)
if st != 200: sys.exit(f"FAIL at token exchange ({st}): {body}")
payload = tok["access_token"].split(".")[1]
claims = json.loads(base64.urlsafe_b64decode(payload + "=" * (-len(payload) % 4)))
scp = claims.get("scp", "").split()
print("Granted scopes:", " ".join(scp))
if "Mail.Send" not in scp: sys.exit("FAIL: Mail.Send not granted")
print("OK: Graph Mail.Send granted for", claims.get("upn") or claims.get("unique_name"))

if "--send" in sys.argv:
    me = claims.get("upn") or claims.get("unique_name")
    mime = f"From: {me}\r\nTo: {me}\r\nSubject: Mailspring Graph send test\r\n\r\nSent via Graph /me/sendMail.\r\n"
    st, body = post("https://graph.microsoft.com/v1.0/me/sendMail", base64.b64encode(mime.encode()),
                    {"Authorization": "Bearer " + tok["access_token"], "Content-Type": "text/plain"})
    print("sendMail:", "OK (202) - check your inbox" if st == 202 else f"FAIL ({st}): {body}")
