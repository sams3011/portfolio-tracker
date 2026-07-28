# Metric Calculation Reference

Explains exactly how every value in the dashboard is computed —
what is pulled from the API, what field is used, and what formula is applied.

---

## Data Source

All price and volume data is fetched from **Yahoo Finance** via the `yfinance`
Python library. One year of daily OHLCV data is pulled per ticker on every run.

```
yf.Ticker(ticker).history(period="1y", auto_adjust=True)
```

**`auto_adjust=True`** means all prices are **split-adjusted and
dividend-adjusted**. This ensures that historical prices are comparable
across stock splits and dividends — the prices will differ slightly from
the raw unadjusted prices shown on some other platforms.

**Fields available from the API per day:**
| Field  | Description |
|--------|-------------|
| Open   | Opening price of the trading session |
| High   | Highest price during the session |
| Low    | Lowest price during the session |
| Close  | Closing price of the trading session |
| Volume | Total shares traded during the session |

High and Low are fetched but not used in any metric currently.

---

## Portfolio Metrics

### Prev Price (Yesterday Price)
- **Field used:** `Close`
- **Value:** Closing price of the second-most-recent trading day in the dataset
- **Code:** `closes.iloc[-2]`
- No formula — raw API value.

---

### Today Price
- **Field used:** `Close`
- **Value:** Closing price of the most recent trading day in the dataset
- **Code:** `closes.iloc[-1]`
- No formula — raw API value.

> Note: "Today" means the last available trading day. On weekends or holidays
> this will be Friday's or the last market day's close, not a live price.

---

### Change %
- **Fields used:** Today Close, Yesterday Close
- **Formula:** `(Today Close - Yesterday Close) / Yesterday Close × 100`
- **Code:** `safe_pct(today_px, yest_px)`
- Rounded to 2 decimal places.

---

### Weekly Return %
- **Fields used:** Today Close, Close from 5 trading days ago
- **Formula:** `(Today Close - Close[-6]) / Close[-6] × 100`
- **Code:** `safe_pct(today_px, closes.iloc[-6])`
- Uses the 6th-to-last row in the dataset (5 trading days back from today).
- Rounded to 2 decimal places.
- Shows `—` if fewer than 6 days of data exist.

---

### Monthly Return %
- **Fields used:** Today Close, Close from 21 trading days ago
- **Formula:** `(Today Close - Close[-22]) / Close[-22] × 100`
- **Code:** `safe_pct(today_px, closes.iloc[-22])`
- Uses the 22nd-to-last row (approximately 1 calendar month of trading days).
- Rounded to 2 decimal places.
- Shows `—` if fewer than 22 days of data exist.

---

### Prev Vol (Yesterday Volume)
- **Field used:** `Volume`
- **Value:** Volume of the second-most-recent trading day
- **Code:** `int(volumes.iloc[-2])`
- No formula — raw API value.

---

### Today Vol
- **Field used:** `Volume`
- **Value:** Volume of the most recent trading day
- **Code:** `int(volumes.iloc[-1])`
- No formula — raw API value.

---

### RelVol (Relative Volume)
- **Field used:** `Volume`
- **Formula:** `Today Volume / Average Volume over the previous 20 trading days`
- **Code:**
  ```python
  avg = volumes.shift(1).rolling(20).mean()
  relvol = volumes / avg
  ```
- The `shift(1)` excludes today from the average — the average is built from
  the 20 days *before* today, not including today itself.
- A RelVol of **1.0** means today's volume exactly matches the 20-day average.
- A RelVol of **2.0** means today traded at double the normal volume.
- Shows `—` if fewer than 21 days of data exist.
- Rounded to 2 decimal places.

---

### Z Score
- **Fields used:** `Open`, `Close`
- **What it measures:** How unusual today's overnight price move is relative
  to recent history. A large absolute Z score means today's gap open was
  abnormally large.
- **Step 1 — Overnight Log Return for each day:**
  ```
  OvernightLR(t) = LN( Open(t) / Close(t-1) )
  ```
  This measures the log return from yesterday's close to today's open
  (the "overnight gap").
- **Step 2 — Rolling standard deviation (sigma):**
  ```
  sigma(t) = StdDev of OvernightLR over the previous 20 days
             (excluding today — shift(1) is applied)
  ```
- **Step 3 — Z Score:**
  ```
  Z(t) = OvernightLR(t) / sigma(t)
  ```
- **Code:**
  ```python
  ovn_lr = np.log(df['Open'] / df['Close'].shift(1))
  sigma  = ovn_lr.shift(1).rolling(20).std()
  z      = ovn_lr / sigma
  ```
- Interpretation:
  - Z near 0 → normal day, gap in line with recent history
  - Z > 2 → unusually large positive gap open
  - Z < -2 → unusually large negative gap open
- Shows `—` if fewer than 21 days of data exist.
- Rounded to 2 decimal places.

---

### RSI (Relative Strength Index)
- **Field used:** `Close`
- **Period:** 14 trading days (standard)
- **Formula — Wilder's smoothing method:**
  ```
  Delta(t)    = Close(t) - Close(t-1)
  Gain(t)     = Delta(t) if Delta(t) > 0, else 0
  Loss(t)     = |Delta(t)| if Delta(t) < 0, else 0

  AvgGain     = Exponential weighted mean of Gain (span = 13, min 14 periods)
  AvgLoss     = Exponential weighted mean of Loss (span = 13, min 14 periods)

  RS          = AvgGain / AvgLoss
  RSI         = 100 - (100 / (1 + RS))
  ```
- **Code:**
  ```python
  delta    = closes.diff()
  gain     = delta.clip(lower=0)
  loss     = -delta.clip(upper=0)
  avg_gain = gain.ewm(com=13, min_periods=14).mean()
  avg_loss = loss.ewm(com=13, min_periods=14).mean()
  rs       = avg_gain / avg_loss
  rsi      = 100 - (100 / (1 + rs))
  ```
- Interpretation:
  - RSI > 70 → overbought (stock may have risen too fast)
  - RSI < 30 → oversold (stock may have fallen too fast)
  - RSI 40–60 → neutral
- Shows `—` if fewer than 14 days of data exist.
- Rounded to 1 decimal place.

---

### 200 DMA (200-Day Moving Average)
- **Field used:** `Close`
- **Formula:** Simple arithmetic mean of the last 200 closing prices
  ```
  200 DMA = (Close[-200] + Close[-199] + ... + Close[-1]) / 200
  ```
- **Code:** `closes.rolling(200).mean().iloc[-1]`
- Shows `—` if fewer than 200 days of data exist (requires ~10 months of history;
  since 1 year is fetched this will populate for most stocks).
- Rounded to 2 decimal places.

---

### 50 DMA (50-Day Moving Average)
- **Field used:** `Close`
- **Formula:** Simple arithmetic mean of the last 50 closing prices
  ```
  50 DMA = (Close[-50] + Close[-49] + ... + Close[-1]) / 50
  ```
- **Code:** `closes.rolling(50).mean().iloc[-1]`
- Shows `—` if fewer than 50 days of data exist.
- Rounded to 2 decimal places.

---

### 52W High (52-Week High)
- **Field used:** `Close`
- **Value:** Maximum closing price over the last 252 trading days
- **Code:** `closes.tail(252).max()`
- 252 is the standard number of trading days in a year.
- Uses **closing prices only** — not intraday highs.
- Rounded to 2 decimal places.

---

### 52W Low (52-Week Low)
- **Field used:** `Close`
- **Value:** Minimum closing price over the last 252 trading days
- **Code:** `closes.tail(252).min()`
- Uses **closing prices only** — not intraday lows.
- Rounded to 2 decimal places.

---

### From High % (% below 52-Week High)
- **Fields used:** Today Close, 52W High
- **Formula:** `(Today Close - 52W High) / 52W High × 100`
- **Code:** `safe_pct(today_px, high_52w)`
- This will almost always be **negative or zero** — it tells you how far the
  stock is below its 52-week high.
- Example: -15.2% means the stock is 15.2% below its yearly peak.
- Rounded to 2 decimal places.

---

### From Low % (% above 52-Week Low)
- **Fields used:** Today Close, 52W Low
- **Formula:** `(Today Close - 52W Low) / 52W Low × 100`
- **Code:** `safe_pct(today_px, low_52w)`
- This will almost always be **positive or zero** — it tells you how far the
  stock has recovered from its 52-week low.
- Example: +42.0% means the stock is 42% above its yearly trough.
- Rounded to 2 decimal places.

---

## Benchmark Metrics

Benchmarks (Nifty 50, Bank Nifty, Brent Crude, Gold, Dow, India VIX) use a
smaller set of the same calculations:

| Metric | Same formula as portfolio? |
|--------|---------------------------|
| Prev Price | Yes — `Close.iloc[-2]` |
| Today Price | Yes — `Close.iloc[-1]` |
| Change % | Yes |
| Weekly Return % | Yes |
| Monthly Return % | Yes |
| Z Score | Yes — same overnight log return formula |

Volume, RSI, DMA, and 52W levels are not shown for benchmarks.

---

## Expiry Dates

Computed entirely from Python date logic — no API call is made.
NSE F&O expiry rules are fixed by the exchange:

| Instrument | Rule | Code logic |
|------------|------|------------|
| Stock F&O (individual stocks) | Last **Tuesday** of each month | `last_weekday_of_month(year, month, weekday=1)` |
| Nifty 50 monthly | Last **Thursday** of each month | `last_weekday_of_month(year, month, weekday=3)` |
| Nifty 50 weekly | Every **Thursday** | Next Thursday from today |
| Bank Nifty monthly | Last **Wednesday** of each month | `last_weekday_of_month(year, month, weekday=2)` |
| Bank Nifty weekly | Every **Wednesday** | Next Wednesday from today |

**Days left** = `expiry_date - today` in calendar days.

If today is the last Tuesday/Thursday/Wednesday of the month, the monthly and
weekly labels merge into "Monthly + Weekly Expiry".

If the stock F&O expiry (last Tuesday) has already passed for this month,
the countdown rolls to the next month's last Tuesday.

---

## Important Caveats

1. **Prices are end-of-day, not live.** The script runs at 6 PM IST after
   market close. All values reflect that day's final closing prices.

2. **Auto-adjusted prices.** Yahoo Finance adjusts historical prices for
   splits and dividends. This is standard practice for return calculations
   but means the historical Close values may not match what you saw on-screen
   on that date.

3. **"Today" on weekends/holidays.** If run on a non-trading day, `Close[-1]`
   will be the last trading day's close (e.g. Friday's close on Saturday).

4. **52W High/Low uses closing prices.** Intraday highs/lows are not used.
   A stock may have touched a higher intraday price that is not reflected here.

5. **Minimum data requirements.** Some metrics show `—` until enough history
   accumulates: Z Score needs 21 days, RSI needs 14 days, 200 DMA needs 200 days.
