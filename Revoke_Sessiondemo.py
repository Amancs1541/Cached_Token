import msal
import os
import sys

TENANT = os.environ.get("TENANT_ID")
CLIENT = os.environ.get("CLIENT_ID")

if not TENANT or not CLIENT:
    sys.exit("Set TENANT_ID and CLIENT_ID environment variables before running.")

if len(sys.argv) != 2 or sys.argv[1] not in ("login", "check"):
    sys.exit("Usage: python Revoke_Sessiondemo.py login|check")

CACHE = "leaver_cache.bin"

cache = msal.SerializableTokenCache()
if os.path.exists(CACHE):
    cache.deserialize(open(CACHE).read())
app = msal.PublicClientApplication(
    CLIENT, authority=f"https://login.microsoftonline.com/{TENANT}", token_cache=cache
)

if sys.argv[1] == "login":
    flow = app.initiate_device_flow(scopes=["Mail.Read", "Calendars.Read"])
    print(flow["message"])
    r = app.acquire_token_by_device_flow(flow)
    open(CACHE, "w").write(cache.serialize())
    print("login:", "OK" if "access_token" in r else r.get("error"))
else:
    acc = app.get_accounts()[0]
    r1 = app.acquire_token_silent(["Mail.Read"], account=acc)
    print("cached access token still usable:", bool(r1 and "access_token" in r1))
    r2 = app.acquire_token_silent(["Mail.Read"], account=acc, force_refresh=True)
    print("refresh token still works:", bool(r2 and "access_token" in r2),
          (r2 or {}).get("error"))