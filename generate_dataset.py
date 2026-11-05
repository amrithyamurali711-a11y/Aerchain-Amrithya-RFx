import json, random, math, os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from docx import Document
from docx.shared import Pt
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from PIL import Image, ImageDraw, ImageFont, ImageFilter
random.seed(7)
OUT="out"; os.makedirs(OUT,exist_ok=True)

# ---------------- RFx: 30 lines ----------------
types=[("3-ply",150,"RSC"),("5-ply",200,"RSC"),("7-ply",250,"RSC")]
lines=[]
sizes=[(200,150,100),(250,200,150),(300,200,150),(300,250,200),(350,250,200),(400,300,200),(400,300,250),(450,350,250),(500,350,300),(500,400,300)]
i=0
for ply,gsm,style in types:
    for (l,w,h) in sizes:
        i+=1
        # surface area m2 approx for RSC: 2*(l+w)*(h+w) mm2
        area=2*(l+w)*(h+w)/1e6
        plyf={"3-ply":3,"5-ply":5,"7-ply":7}[ply]
        weight_g=round(area*gsm*plyf*0.62)  # rough
        qty=random.choice([2000,3000,5000,8000,10000,15000,20000])
        if ply=="7-ply": qty=random.choice([1000,1500,2000,3000])
        lines.append(dict(line=i,item=f"{ply} corrugated box {l}x{w}x{h} mm",ply=ply,length_mm=l,width_mm=w,height_mm=h,
            board_gsm=gsm,flute="BC" if ply!="3-ply" else "B",print="1-colour flexo" if i%3 else "Plain",
            approx_weight_g=weight_g,qty=qty,uom="pieces"))
rfx=dict(rfx_id="RFQ-PKG-2026-114",buyer="Anvaya Foods Pvt Ltd (FMCG snacks)",category="Corrugated packaging",
    issued="2026-10-20",due="2026-10-29",currency="INR",delivery="Ex-works pricing not accepted. Delivered to Anvaya DC, Hosur, Tamil Nadu",
    contract="12-month rate contract, monthly call-offs",payment="Net 45 preferred",lines=lines,
    questionnaire=[
      dict(id="Q1",q="Do you hold a valid ISO 9001 certificate? Attach copy.",critical=True),
      dict(id="Q2",q="Do you perform BCT (box compression) and burst-strength testing on every batch?",critical=True),
      dict(id="Q3",q="Standard lead time from PO to delivery (days)?",critical=False),
      dict(id="Q4",q="Minimum order quantity per SKU?",critical=False),
      dict(id="Q5",q="Payment terms offered?",critical=False),
      dict(id="Q6",q="Is the kraft paper FSC-certified or recycled-content? State %.",critical=False)])
json.dump(rfx,open(f"{OUT}/00_RFx_RFQ-PKG-2026-114.json","w"),indent=2)

# readable RFx doc
d=Document(); d.add_heading("Request for Quotation RFQ-PKG-2026-114",0)
d.add_paragraph(f"Buyer: {rfx['buyer']}\nCategory: Corrugated packaging\nIssued: 20 Oct 2026   Response due: 29 Oct 2026\nDelivery: {rfx['delivery']}\nContract: {rfx['contract']}\nPayment: {rfx['payment']}\nCurrency: INR, price per piece, delivered, GST extra")
d.add_heading("Line items",1)
t=d.add_table(rows=1,cols=7); t.style="Light Grid Accent 1"
for c,h in zip(t.rows[0].cells,["#","Item","Board GSM","Flute","Print","Approx wt (g)","Qty (pcs)"]): c.text=h
for L in lines:
    r=t.add_row().cells
    for c,v in zip(r,[L['line'],L['item'],L['board_gsm'],L['flute'],L['print'],L['approx_weight_g'],L['qty']]): c.text=str(v)
d.add_heading("Supplier questionnaire",1)
for q in rfx["questionnaire"]: d.add_paragraph(f"{q['id']}. {q['q']}"+("  (mandatory)" if q['critical'] else ""))
d.add_paragraph("Please quote in the attached template (RFQ-PKG-2026-114_template.xlsx).")
d.save(f"{OUT}/00_RFx_RFQ-PKG-2026-114.docx")

# template xlsx
wb=Workbook(); ws=wb.active; ws.title="Quote"
ws.append(["Line #","Item","Qty (pcs)","Unit price INR / piece","Total INR","Remarks"])
for L in lines: ws.append([L['line'],L['item'],L['qty'],None,None,None])
wb.save(f"{OUT}/00_RFQ-PKG-2026-114_template.xlsx")

# ---------------- true prices (hidden ground truth) ----------------
def base(L): return L['approx_weight_g']/1000*40 + (1.2 if L['print']!="Plain" else 0)
vend={"A":1.00,"B":0.95,"C":0.92,"D":1.04,"E":0.98}
truth={}
for v,f in vend.items():
    truth[v]={L['line']:round(base(L)*f*random.uniform(0.95,1.06),2) for L in lines}
USD=84.0

# ---------- A: Excel ignoring template ----------
wb=Workbook(); ws=wb.active; ws.title="Our Offer"
ws["A1"]="SHREE BALAJI PACKAGING INDUSTRIES — Quotation"; ws["A1"].font=Font(bold=True,size=14)
ws["A2"]="Ref: your RFQ 114 | Valid 30 days | Prices per pc, delivered Hosur, GST 18% extra"
ws.append([]); hdr=["Sl","Description","Size (LxWxH)","Ply","Rate/pc (Rs.)","Order Qty","Amount (Rs.)","Lead time"]
ws.append(hdr)
for c in ws[4]: c.font=Font(bold=True,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="2F5597")
order=lines[:]; random.shuffle(order)  # reordered rows
for k,L in enumerate(order,1):
    p=truth["A"][L['line']]
    ws.append([k,f"Corr. box {L['ply'].upper()} {'printed' if L['print']!='Plain' else 'plain'}",f"{L['length_mm']}x{L['width_mm']}x{L['height_mm']}",L['ply'][0],p,L['qty'],round(p*L['qty'],2),"12 days"])
ws.append([]); ws.append(["","","","","","TOTAL",f"=SUM(G5:G{4+len(order)})"])
ws2=wb.create_sheet("Compliance")
for r in [["Question","Our answer"],["ISO 9001","Yes, certificate no. QMS/IN/22/4471, valid to Mar 2027 (attached)"],
          ["BCT / burst testing","Every batch, in-house lab, reports shared with dispatch"],["Lead time","12 working days"],
          ["MOQ","1,000 pcs per SKU"],["Payment","Net 30. Net 45 possible at +1.5%"],["Kraft","70% recycled content, not FSC"]]: ws2.append(r)
wb.save(f"{OUT}/A_ShreeBalaji_quotation.xlsx")

# ---------- B: PDF letterhead, USD, discount in footnote ----------
c=canvas.Canvas(f"{OUT}/B_SiamKraft_offer.pdf",pagesize=A4); W,H=A4
c.setFillColorRGB(0.05,0.3,0.25); c.rect(0,H-35*mm,W,35*mm,fill=1,stroke=0)
c.setFillColorRGB(1,1,1); c.setFont("Helvetica-Bold",20); c.drawString(18*mm,H-18*mm,"SIAM KRAFT PACKAGING CO., LTD.")
c.setFont("Helvetica",9); c.drawString(18*mm,H-25*mm,"Amata City, Chonburi, Thailand  |  India office: Guindy, Chennai  |  sales@siamkraft.co.th")
c.setFillColorRGB(0,0,0); c.setFont("Helvetica-Bold",12); c.drawString(18*mm,H-48*mm,"Commercial Offer — RFQ-PKG-2026-114 (Anvaya Foods)")
c.setFont("Helvetica",9); c.drawString(18*mm,H-55*mm,"All prices in USD per piece, CIF Chennai port. Inland freight to Hosur by buyer.*")
y=H-65*mm; c.setFont("Helvetica-Bold",8)
for x,h in zip([18,30,100,130,160],["No.","Specification","Qty","USD/pc","Ext. USD"]): c.drawString(x*mm,y,h)
c.setFont("Helvetica",8)
for L in lines:
    y-=5*mm
    if y<40*mm:
        c.drawString(18*mm,25*mm,"continued..."); c.showPage(); y=H-25*mm; c.setFont("Helvetica",8)
    u=round(truth["B"][L['line']]/USD,4)
    c.drawString(18*mm,y,str(L['line'])); c.drawString(30*mm,y,f"{L['ply']} RSC {L['length_mm']}x{L['width_mm']}x{L['height_mm']} {L['board_gsm']}gsm")
    c.drawRightString(115*mm,y,f"{L['qty']:,}"); c.drawRightString(148*mm,y,f"{u:.4f}"); c.drawRightString(185*mm,y,f"{u*L['qty']:,.2f}")
y-=12*mm; c.setFont("Helvetica-Bold",9); c.drawString(18*mm,y,"Compliance: ISO 9001:2015 (TÜV), BCT + burst every lot. Lead time 28 days incl. sea freight. MOQ 5,000 pcs.")
y-=5*mm; c.drawString(18*mm,y,"Payment: 30% advance, 70% against B/L. Kraft: FSC Mix, 100%.")
c.setFont("Helvetica",6.5)
c.drawString(18*mm,18*mm,"* Prices valid at USD/THB 33.5. A volume rebate of 4% applies to the full order value where the annual contract exceeds USD 150,000,")
c.drawString(18*mm,14*mm,"  settled as a credit note at contract end. Customs duty and IGST payable by importer.")
c.save()

# ---------- C: Word paragraph, 27 of 30, fails quality ----------
skip={4,17,28}
d=Document(); d.add_paragraph("Vardhman Box & Carton Works\nPeenya Industrial Area, Bengaluru")
d.add_paragraph("Dear Sir / Madam,")
d.add_paragraph("Thank you for the opportunity. Please find our best rates below. We are a family run unit with 22 years in corrugation and supply to several regional snack brands.")
paras=[]
for L in lines:
    if L['line'] in skip: continue
    p=truth["C"][L['line']]
    paras.append(f"for the {L['ply']} {L['length_mm']} by {L['width_mm']} by {L['height_mm']} box we can do Rs {p:.2f} each")
d.add_paragraph("Coming to pricing, "+"; ".join(paras[:9])+".")
d.add_paragraph("Similarly "+", ".join(paras[9:18])+".")
d.add_paragraph("And finally "+", ".join(paras[18:])+". All rates are inclusive of delivery to Hosur. GST extra as applicable.")
d.add_paragraph("We are sorry we cannot quote for a few of the sizes this time as our die is not available for them. We do not currently hold ISO certification but our quality is checked visually by our supervisor before every dispatch. Lead time is around 7 days, MOQ 500 pieces, and we accept payment in 15 days.")
d.add_paragraph("Regards,\nR. Vardhman\n98450 xxxxx")
d.save(f"{OUT}/C_Vardhman_quote.docx")

# ---------- D: angled photo of rate card, per 100 pcs ----------
img=Image.new("RGB",(1400,2000),(250,248,240)); dr=ImageDraw.Draw(img)
fb=ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",44)
fr=ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",27)
fs=ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",24)
dr.text((60,50),"KAVERI CORRUGATORS",font=fb,fill=(20,20,90))
dr.text((60,110),"RATE CARD 2026-27  (Hosur / Krishnagiri belt)",font=fr,fill=(40,40,40))
dr.text((60,150),"Rates per 100 pcs, delivered. GST extra.",font=fr,fill=(150,0,0))
y=210; dr.text((60,y),"Sl  Box (mm)            Ply    Rs/100 pcs",font=fr,fill=(0,0,0)); y+=12
dr.line((55,y+30,1340,y+30),fill=(0,0,0),width=2); y+=45
for L in lines:
    p=truth["D"][L['line']]*100
    dr.text((60,y),f"{L['line']:>2}  {L['length_mm']}x{L['width_mm']}x{L['height_mm']:<6}   {L['ply']:<6} {p:>9,.0f}",font=fs,fill=(0,0,0)); y+=50
y+=20
for t_ in ["ISO 9001 certified - cert attached","Lab test (BCT/burst) every batch","Lead time 10 days | MOQ 2000/SKU","Payment 45 days","Kraft 100% recycled"]:
    dr.text((60,y),t_,font=fs,fill=(30,30,30)); y+=38
# perspective + rotate + noise = phone photo
img=img.rotate(-7,expand=True,fillcolor=(90,80,70))
w,h=img.size
img=img.transform((w,h),Image.PERSPECTIVE,(1,0.06,-60, 0.02,1,-30, 0.00004,0.00002),fillcolor=(90,80,70))
img=img.filter(ImageFilter.GaussianBlur(1.2))
px=img.load()
for _ in range(60000):
    x=random.randrange(w); yy=random.randrange(h); r,g,b=px[x,yy]; n=random.randint(-18,18); px[x,yy]=(max(0,min(255,r+n)),)*1 + (max(0,min(255,g+n)),max(0,min(255,b+n)))
# shadow gradient
sh=Image.new("L",(w,h)); sd=ImageDraw.Draw(sh)
for k in range(w): sd.line((k,0,k,h),fill=int(70*k/w))
img=Image.composite(Image.new("RGB",(w,h),(0,0,0)),img,sh.point(lambda v:v//2))
img=img.resize((w*2//3,h*2//3)); img.save(f"{OUT}/D_Kaveri_ratecard_photo.jpg",quality=72)

# ---------- E: one-line email ----------
open(f"{OUT}/E_Lakshmi_email.eml","w").write("""From: Suresh <suresh.lakshmipack@gmail.com>
To: procurement@anvayafoods.in
Subject: Re: RFQ-PKG-2026-114
Date: Tue, 28 Oct 2026 22:41:07 +0530

Sir, 42/kg for the 5-ply, 38 for the 3-ply, 7-ply 46. rest same as last year. freight extra. ISO yes.

Sent from my iPhone
""")

# ---------- attachments ----------
for fn,name,cert in [("A_ShreeBalaji_ISO9001.pdf","Shree Balaji Packaging Industries","QMS/IN/22/4471  valid until 31 Mar 2027"),
                     ("D_Kaveri_ISO9001.pdf","Kaveri Corrugators","QMS/IN/19/2210  valid until 30 Jun 2025")]:
    c=canvas.Canvas(f"{OUT}/{fn}",pagesize=A4)
    c.setFont("Helvetica-Bold",22); c.drawCentredString(W/2,H-60*mm,"CERTIFICATE OF REGISTRATION")
    c.setFont("Helvetica",13); c.drawCentredString(W/2,H-80*mm,"This is to certify that the Quality Management System of")
    c.setFont("Helvetica-Bold",16); c.drawCentredString(W/2,H-95*mm,name)
    c.setFont("Helvetica",13); c.drawCentredString(W/2,H-110*mm,"conforms to ISO 9001:2015 for the manufacture of corrugated boxes")
    c.drawCentredString(W/2,H-130*mm,"Certificate no. "+cert); c.save()

# ground truth for testing only
json.dump(dict(note="TEST ORACLE ONLY - never shown to the app",usd_inr=USD,truth_inr_per_piece=truth,
  traps=dict(A="rows shuffled, renamed columns, answers on 2nd sheet",
             B="USD per piece, CIF port so inland freight extra, 4% rebate in footnote conditional on >USD150k",
             C="27/30 lines (missing 4,17,28), prices inside prose, no ISO and visual-only QC -> fails mandatory",
             D="photo, rates per 100 pcs, ISO cert attached is EXPIRED (Jun 2025) despite claiming certified",
             E="per kg pricing needs weight conversion, 'rest same as last year' is unknowable, freight extra, ISO claimed without proof")),
  open(f"{OUT}/zz_test_oracle.json","w"),indent=2)
print("done", sorted(os.listdir(OUT)))
