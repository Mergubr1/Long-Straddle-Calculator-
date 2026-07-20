# Long Straddle Calculator

Model a long straddle (buy one call + buy one put, same strike & expiration):
combined premium, breakevens, P/L at any expiration price, and a full payoff
table/chart across a price range.

This repo has two independent ways to use it:

| | What it is | Where |
|---|---|---|
| 🌐 **Web app** | Single-page, client-side calculator — no install, no server | [`index.html`](./index.html) |
| 📊 **Excel generator** | Python script that builds a fully formula-driven `.xlsx` | [`generator/build_xlsx.py`](./generator/build_xlsx.py) |

Both implement the same math and will always agree on results for the same inputs.

---

## 1. Web app

`index.html` is fully self-contained (HTML/CSS/JS, one Google Fonts link, no
build step). Every input recalculates the results, payoff table, and chart
live in the browser — nothing is sent anywhere.

**Run it locally:**
```bash
open index.html          # macOS
# or just double-click the file, or serve it:
python3 -m http.server    # then visit http://localhost:8000
```

**Deploy for free with GitHub Pages:**
1. Push this repo to GitHub.
2. Repo → **Settings → Pages** → Source: `Deploy from a branch` → Branch: `main` / root.
3. Your calculator will be live at `https://<your-username>.github.io/<repo-name>/`.

## 2. Excel generator

`generator/build_xlsx.py` writes a `.xlsx` workbook where every result cell
is a real Excel formula (not a hardcoded number), so it recalculates in
Excel/LibreOffice/Google Sheets exactly like the web app does in-browser.

**Setup:**
```bash
cd generator
pip install -r requirements.txt
```

**Generate a workbook:**
```bash
python build_xlsx.py                          # -> Long_Straddle_Calculator.xlsx, symbol XYZ
python build_xlsx.py --symbol AAPL -o AAPL_straddle.xlsx
```

The workbook includes:
- Editable trade inputs (blue cells) and automatic results (green cells)
- A dynamic expiration payoff table (scales with Strike, Scenario Low/High %, and Scenario Intervals — up to 40 intervals)
- A payoff line chart
- Conditional formatting (profit/loss coloring)
- An assumptions/notes section documenting what the model does and doesn't account for

## What the model does — and doesn't — account for

- Values each leg **at expiration only**, using intrinsic value (`MAX(price − strike, 0)`, etc.).
  It does **not** price the options before expiration — no Black-Scholes / time-value estimate.
- "Number of Straddles" scales both legs equally (1 call + 1 put per straddle), matching a standard long straddle.
- Fees & Commissions are a single total applied once to the position, not per contract.

## License

MIT — do whatever you like with it.
