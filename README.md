# QuoteLens · Kill the Quote Spreadsheet

**Live app:** https://aerchain-quotelens-amrithya.netlify.app
**Built for:** Aerchain product assignment, by Amrithya Muralikrishnan
**Demo video:** see the submission email

QuoteLens takes an RFx from draft to award without a spreadsheet. Vendors reply however they like (Excel, PDF in USD, Word letter, phone photo, one-line email), AI reads each reply as it lands, code normalises every price to ₹ per piece, and a buyer gets one comparison where every number shows how sure it is. A VP can then ask questions in plain words and get answers computed from the data, not generated.

## The idea in one line

Reading messy files is no longer the hard part. **The real risk is a confident wrong number**: a price per 100 pieces read as per piece, dollars compared with rupees, a missing line counted as ₹0. QuoteLens is built so that every number is either trustworthy or clearly says it isn't.

## What is real and what is stubbed

The brief allowed plumbing to be stubbed but required the AI loops to be real. This is how that line was drawn.

| Part | Status |
|---|---|
| RFx co-pilot (draft and edit the RFx from plain words) | **Real.** Gemini, with the buyer's purchase history as context |
| Reading vendor replies (Excel, PDF, Word, photo, email) | **Real.** Gemini reads each file. PDFs and photos go in as inline data, Excel and Word are converted to text in the browser |
| Normalisation (÷100, USD→₹, per-kg → per-piece) | **Code.** The model never does the maths |
| Grounding check (does the vendor's own text contain that price, that size, that ply?) | **Code.** Prices that fail are marked red and excluded from totals |
| Analyst (plain-language questions) | **Real AI, computed numbers.** Gemini picks a tool, code computes the result, Gemini narrates only those numbers |
| Vendor score | **Code.** Weighted, explainable, buyer rating included |
| Award options and memo | **Code** for the options, **Gemini** for the memo text, from the computed facts only |
| Email send and receive, vendor pool, approvals, the RFx home list | **Stubbed.** Replies "arrive" from embedded files, three earlier events are static |
| Rules & settings toggles, AI prompt editor | **Display only** in this prototype, except custom analyst instructions, which are live |

Nothing in the comparison is hardcoded. Change a vendor file in `dataset/`, rebuild, and the numbers change.

## How it works

```
Draft RFx ──► Invite vendors ──► Replies land ──► AI reads each reply ──► Code normalises + grounds
                                                                              │
                 Award memo ◄── Award options ◄── Vendor score ◄── Compare ◄──┘
                                       ▲
                           Ask the analyst (AI picks tool → code computes → AI narrates)
```

**Confidence tiers on every price**

| Colour | Meaning | In totals? |
|---|---|---|
| Green | Read with confidence, found verbatim in the vendor's document | Yes |
| Amber | Converted or estimated (currency, per-100, per-kg from RFx box weight) | Yes, labelled |
| Red | Could not be verified in the vendor's own text | **No**, until the buyer confirms |
| Grey | Not quoted | Never treated as ₹0 |

**Vendor status from the mandatory questionnaire**

- Fails a mandatory question (no ISO, expired certificate): **Disqualified**, score capped at 40
- Claims without proof, or not answered: **Conditional**, can win lines once proof arrives
- Passes with evidence: **Qualified**

**Vendor score** = price 35% · compliance 25% · data confidence 15% · completeness 10% · buyer rating 15%. Weights are editable under Settings → Setup → Evaluation Criteria, with templates and a live preview.

## Reliability

Gemini on the free tier is throttled under load, so the app has: retries with backoff, a model fallback chain (`gemini-flash-latest` → `gemini-2.5-flash` → lite), a 90-second timeout per call, one reply read at a time, automatic re-queue of a failed read, and a cache of successful reads per RFx (labelled "Saved AI read" on screen when reused). In production this would run server-side with queuing and a paid tier.

## Run it

1. Open the live link, or open `site/index.html` locally in Chrome.
2. Click **Model & API key** (bottom left), paste a Gemini API key from Google AI Studio, save.
3. **+ New RFx** → type what you need (for example: "Need corrugated boxes for our Hosur plant again, same sizes as last quarter. ISO 9001 and batch test reports are mandatory.") → Next → Send to vendors → watch the replies land.

Everything runs in the browser. The key is stored in `localStorage` only and sent to Google's API directly.

## Repository layout

```
app_template.html      the app source (single file, with __EMBED__ / __RFX__ placeholders)
build.py               embeds the dataset and writes site/index.html
site/index.html        the deployed file
generate_dataset.py    generates the RFx and the five deliberately messy vendor replies
dataset/               the RFx (JSON, DOCX, quote template) and the vendor replies
tests/answer_key.json  the true ₹/piece for every vendor and line, used to check extraction
docs/Video_Script.md   the demo video script
```

## Test data

`generate_dataset.py` builds a 30-line corrugated-box RFx for Anvaya Foods and five replies with the traps a buyer actually meets:

| Vendor | Format | Trap |
|---|---|---|
| Shree Balaji | Excel, own layout | Rows shuffled, serial numbers that are not RFx line numbers |
| Siam Kraft (TH) | PDF | USD, CIF Chennai port only, conditional 4% rebate |
| Vardhman | Word letter | 27 of 30 lines, no ISO |
| Kaveri | Phone photo of a rate card | Prices per 100 pieces, ISO certificate expired |
| Lakshmi | One-line email | Prices per kg, freight extra, "ISO yes" with no proof |

`tests/answer_key.json` holds the correct ₹ per piece for every line. The app never loads it; it exists so extraction can be checked. On the runs behind the demo, four vendors matched exactly and the USD vendor was within ₹47 on an ₹18.4 lakh total, from rounding.

## What I would build next

1. Award memo to approval workflow, with the approver's view
2. Auto-reminders and a vendor portal for late or incomplete replies, including a one-click "send us your ISO certificate"
3. Lead time, payment terms and past performance as scoring criteria (needs vendor history data)
4. Server-side extraction with per-document caching and a human review queue for red prices
5. Multi-category profiles with their own prompts, and an audit trail for every buyer override
