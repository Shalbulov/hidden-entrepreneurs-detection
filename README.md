# Hidden Entrepreneurs Detection

**Mastercard × AIESEC Hackathon — May 2026**

Detecting individuals who run a commercial business through ordinary consumer
cards, by training a binary classifier on the behavioural signal of 25K verified
business cards and applying it to 80K consumer cards.

---

## Problem

Mastercard's MDQ dataset contains two card populations:

| Segment | Cards | Transactions |
|---|---|---|
| Business cardholders | 25 K | ~3 M |
| Consumer cardholders | 80 K | ~10 M |

A non-trivial fraction of the *consumer* population are actually
micro-entrepreneurs (resellers, freelancers, repair shops, tutors, etc.) using
a personal card for commercial flows. Finding them lets the bank pitch a
proper business product (acquiring, business card, line of credit).

## Approach

1. **Aggregate** transactions to one row per card (~30 behavioural features:
   amount stats, MCC mix, time-of-day mix, recurrence, channel, geography).
2. **Train** a binary classifier: business (1) vs consumer (0).
3. **Treat label noise** — consumer ≠ negative; some consumers *are*
   entrepreneurs. Two techniques:
   - **Iterative cleaning** — drop suspicious consumer rows from train, retrain.
   - **PU Learning** (2 iterations) — only the bottom 70 % of consumer scores
     are treated as reliable negatives. Iterate to remove circular bias.
4. **Boost** the best PU labelling with **XGBoost + early stopping**.
5. **Score** all 80K consumer cards and bucket them into priority tiers.

## Results

| Model | Test AUC |
|---|---|
| Random Forest (baseline) | ~0.92 |
| XGBoost (baseline) | ~0.93 |
| Random Forest (iterative cleaning) | ~0.94 |
| Random Forest PU (2 iter) | ~0.95 |
| **XGBoost PU + early stop** | **~0.96** |

Final scoring tiers across the 80K consumer base:

- **Critical** (P ≥ 0.50) — pitch business card + acquiring.
- **High risk** (P ≥ 0.30) — targeted outreach via relationship manager.
- **Monitoring** (P ≥ 0.10) — observe 2–3 months before contact.

## Repo layout

```
hidden_entrepreneurs_final.ipynb   main notebook (load → features → models → scoring)
README.md                          this file
.gitignore                         excludes the large parquet inputs
```

The three input parquets are **not** committed (the consumer file is over
GitHub's 100 MB limit). To run the notebook, place these files next to it:

```
business_cards_MDQ.parquet     ~52 MB
consumer_cards_MDQ.parquet    ~155 MB
merchants_reference.parquet    ~45 KB
```

## Running

```bash
pip install pandas numpy scikit-learn xgboost matplotlib pyarrow
jupyter notebook hidden_entrepreneurs_final.ipynb
```

The notebook is self-contained: top-to-bottom execution produces
`hidden_entrepreneurs.csv` (scored consumer cards) and `results.png`
(feature importance, ROC curves, score distribution, entrepreneur profile).

## Key engineering fixes (vs. the first-pass code)

- **No data leakage** — train/val/test split by `card_number`; every retraining
  step (cleaning, PU iterations) uses *only* train cards.
- **`is_night` bug fixed** — `Series.between(22, 6)` is always False (22 > 6);
  replaced with `(hour >= 22) | (hour <= 6)`.
- **PU Learning iterated twice** — the second pass uses the cleaner decision
  boundary from the first pass, removing the circular bias.
- **Early stopping for XGBoost** — 1000 max trees, halts when val AUC plateaus.
