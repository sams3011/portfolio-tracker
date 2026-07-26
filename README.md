# Portfolio Tracker

Fetches daily stock data from Yahoo Finance, computes metrics, and publishes a mobile-friendly dashboard via GitHub Pages. Runs automatically every weekday at **6 PM IST**.

## Live dashboard

`https://<your-username>.github.io/portfolio-tracker/`

---

## Adding / removing stocks

Edit `my_portfolio.txt` directly on GitHub (pencil icon) and commit. The next scheduled run picks up the change.

Format: `Display Name | YAHOO_TICKER`

Indian stocks use `.NS` (NSE) or `.BO` (BSE) suffix.

---

## Metrics computed

**Portfolio stocks:** Prev Price, Today Price, Change %, Weekly %, Monthly %, Prev Vol, Today Vol, RelVol, Z Score, RSI, 200 DMA, 50 DMA, 52W High, 52W Low, % from High, % from Low

**Benchmarks:** Prev Price, Today Price, Change %, Weekly %, Monthly %, Z Score

**Expiry banner:** Shows today's NSE F&O expiry (Nifty / Bank Nifty — weekly and monthly), or days until next expiry.

---

## Outputs

| File | What |
|------|------|
| `docs/dashboard.html` | Live web dashboard (GitHub Pages) |
| `data/portfolio_tracker.xlsx` | Downloadable Excel — Dashboard sheet + per-stock History sheet |

---

## One-time GitHub setup

1. Push this repo to GitHub (public)
2. **Settings → Actions → General → Workflow permissions → Read and write → Save**
3. **Settings → Pages → Source → Deploy from branch → `main` / `docs` → Save**
4. Dashboard will be live at `https://<username>.github.io/portfolio-tracker/`

---

## Run locally

```bash
pip install -r requirements.txt
python portfolio.py
```
