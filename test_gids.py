import urllib.request
import ssl
import sys

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

gids = {
    "Tipificaciones (1109198771)": "1109198771",
    "Daka (1240474880)": "1240474880",
    "Damasco (841459536)": "841459536",
    "Multimax (2089283830)": "2089283830"
}

headers = {'User-Agent': 'Mozilla/5.0'}

for name, gid in gids.items():
    url = f"https://docs.google.com/spreadsheets/d/e/2PACX-1vTfq81DhLQ_8jkbFIAs7OWaO7qkYRis350TTRz_BbbsVucVw4K87Ai0YgiynRIQG1CqRJv9i1V6oEDo/pub?gid={gid}&single=true&output=csv"
    req = urllib.request.Request(url, headers=headers)
    print(f"Testing {name}...")
    try:
        with urllib.request.urlopen(req, context=ctx) as r:
            body = r.read(100).decode('utf-8')
            print(f"  SUCCESS! Header: {body[:60]}")
    except Exception as e:
        print(f"  FAILED: {e}")
        # Try to read error body if available
        if hasattr(e, 'read'):
            try:
                err_body = e.read().decode('utf-8')
                print(f"  Error Body (first 200 chars): {err_body[:200]}")
            except Exception:
                pass
    print("-" * 50)
