import urllib.request
import io
import pandas as pd
import ssl

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTfq81DhLQ_8jkbFIAs7OWaO7qkYRis350TTRz_BbbsVucVw4K87Ai0YgiynRIQG1CqRJv9i1V6oEDo/pub?gid=841459536&single=true&output=csv"
req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})

try:
    with urllib.request.urlopen(req, context=ctx) as r:
        df = pd.read_csv(io.StringIO(r.read().decode('utf-8')))
        print("Damasco GID (841459536) Master DB:")
        print("Shape:", df.shape)
        print("Columns:", list(df.columns))
        print("Unique brands in Damasco DB:")
        print(df['MARCA'].dropna().astype(str).str.upper().str.strip().unique())
        
        # Check if Hyundai is in Damasco DB
        hyundai_df = df[df['MARCA'].astype(str).str.contains('hyundai', case=False, na=False)]
        print(f"\nFound {len(hyundai_df)} Hyundai products in Damasco DB:")
        print(hyundai_df.to_string())
        
        # Check if Daewoo is in Damasco DB
        daewoo_df = df[df['MARCA'].astype(str).str.contains('daewoo', case=False, na=False)]
        print(f"\nFound {len(daewoo_df)} Daewoo products in Damasco DB:")
        print(daewoo_df.to_string())

except Exception as e:
    print("Error:", e)
