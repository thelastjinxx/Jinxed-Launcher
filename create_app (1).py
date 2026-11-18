#!/usr/bin/env python3
"""Create an Entra app registration for a Minecraft launcher, without the az CLI.

Usage:  python3 create_app.py ["Launcher Name"] [tenant-id-or-domain]
Standard library only. Signs in with a device code, then calls Microsoft Graph.
"""
import json
import sys
import time
import urllib.error
import urllib.parse
import urllib.request

# Microsoft Graph Command Line Tools: a Microsoft-owned public client that allows device-code login.
CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
SCOPE = "https://graph.microsoft.com/Application.ReadWrite.All offline_access"

name = sys.argv[1] if len(sys.argv) > 1 else "Jinxed Launcher"
tenant = sys.argv[2] if len(sys.argv) > 2 else "organizations"
BASE = "https://login.microsoftonline.com/%s/oauth2/v2.0" % tenant


def call(url, data=None, jsn=None, token=None):
    headers = {"User-Agent": "create-app/1.0"}
    body = None
    if data is not None:
        body = urllib.parse.urlencode(data).encode()
    elif jsn is not None:
        body = json.dumps(jsn).encode()
        headers["Content-Type"] = "application/json"
    if token:
        headers["Authorization"] = "Bearer " + token
    req = urllib.request.Request(url, data=body, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return r.status, json.loads(r.read() or b"{}")
    except urllib.error.HTTPError as e:
        raw = e.read()
        try:
            return e.code, json.loads(raw)
        except Exception:
            return e.code, {"raw": raw.decode(errors="replace")}


st, dc = call(BASE + "/devicecode", data={"client_id": CLIENT_ID, "scope": SCOPE})
if st != 200:
    sys.exit("Could not start sign-in: %s %s" % (st, dc))

print("\n1. Open:  %s" % dc["verification_uri"])
print("2. Enter code:  %s" % dc["user_code"])
print("3. Sign in and accept the permission prompt.\n")

token = None
deadline = time.time() + int(dc.get("expires_in", 900))
interval = int(dc.get("interval", 5))
while time.time() < deadline:
    time.sleep(interval)
    st, t = call(BASE + "/token", data={
        "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
        "client_id": CLIENT_ID, "device_code": dc["device_code"]})
    if st == 200:
        token = t["access_token"]
        break
    err = t.get("error")
    if err == "authorization_pending":
        continue
    if err == "slow_down":
        interval += 5
        continue
    sys.exit("Sign-in failed: %s" % t.get("error_description", t))
if not token:
    sys.exit("Timed out waiting for sign-in.")

st, app = call("https://graph.microsoft.com/v1.0/applications", token=token, jsn={
    "displayName": name,
    "signInAudience": "PersonalMicrosoftAccount",
    "isFallbackPublicClient": True,
})
if st not in (200, 201):
    msg = app.get("error", app)
    sys.exit("Could not create the app (%s): %s" % (st, json.dumps(msg, indent=2)))

print("Created app registration '%s'" % name)
print("Client ID (paste into the launcher):  %s" % app["appId"])
