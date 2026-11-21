import msal
import os
import sys

TENANT = os.environ.get("TENANT_ID")
CLIENT = os.environ.get("CLIENT_ID")

if not TENANT or not CLIENT:
    sys.exit("Set TENANT_ID and CLIENT_ID environment variables before running.")

app = msal.PublicClientApplication(
    CLIENT, authority=f"https://login.microsoftonline.com/{TENANT}"
)
flow = app.initiate_device_flow(scopes=["Mail.Read", "Calendars.Read"])
if "user_code" not in flow:
    sys.exit(str(flow))
print(flow["message"])          # tells you the code and the web address

r = app.acquire_token_by_device_flow(flow)
if "access_token" in r:
    print("OK:", r["id_token_claims"].get("preferred_username"))
    print("access token lifetime (seconds):", r.get("expires_in"))
    print("refresh token received:", "refresh_token" in r)
else:
    print("DENIED:", r.get("error"))
    print((r.get("error_description") or "")[:200])