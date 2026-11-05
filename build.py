"""Build site/index.html from app_template.html by embedding the dataset (base64) and the RFx JSON.
Usage: python3 build.py
"""
import json, base64, os
d = "dataset"
files = ["A_ShreeBalaji_quotation.xlsx", "A_ShreeBalaji_ISO9001.pdf", "B_SiamKraft_offer.pdf", "C_Vardhman_quote.docx",
         "D_Kaveri_ratecard_photo.jpg", "D_Kaveri_ISO9001.pdf", "E_Lakshmi_email.eml"]
emb = {f: base64.b64encode(open(os.path.join(d, f), "rb").read()).decode() for f in files}
s = open("app_template.html").read()
out = s.replace("__EMBED__", json.dumps(emb)).replace("__RFX__", open(os.path.join(d, "00_RFx_RFQ-PKG-2026-114.json")).read())
os.makedirs("site", exist_ok=True); open("site/index.html", "w").write(out)
print("wrote site/index.html", len(out), "bytes")
