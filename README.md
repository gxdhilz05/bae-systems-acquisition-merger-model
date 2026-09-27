# BAE Systems Acquisition Screening and Merger Model

## Project objective

This project assesses a hypothetical acquisition by BAE Systems of one of three UK-listed defence and aerospace companies:

- Chemring Group
- Cohort plc
- Avon Technologies plc

The analysis will identify the strongest strategic target, establish a defensible offer range, test alternative financing structures, and evaluate the transaction's effect on earnings per share and leverage.

## Central question

> Should BAE Systems acquire Chemring, Cohort or Avon Technologies, and would the selected transaction create value for BAE shareholders at a realistic purchase price?

## Learning-first build process

This repository is being built in stages. A stage is not treated as complete until its purpose, calculations, assumptions and limitations can be explained without relying on the finished model.

1. Understand the transaction and screening logic.
2. Screen the three targets strategically and financially.
3. Select and value the preferred target.
4. Set the offer price and financing structure.
5. Build purchase accounting and the combined income statement.
6. Calculate EPS accretion/dilution and post-deal leverage.
7. Run sensitivities and form a recommendation.
8. Produce the final memo, presentation and interview guide.

## Planned outputs

- Target-screening scorecard
- Standalone operating forecasts
- DCF, trading-comparables and precedent-transactions valuation
- Offer-price and control-premium analysis
- Sources-and-uses schedule
- Purchase-price allocation
- Accretion/dilution model
- Post-transaction leverage analysis
- Sensitivity tables
- Two-page transaction memo
- Ten-minute presentation
- Interview preparation guide

## Repository structure

```text
data/       Source data and data notes
learning/   Plain-English explanations and knowledge checks
model/      Screening and merger-model workbooks
research/   Company research, assumptions and source log
outputs/    Final memo and presentation
```

## Status

**Stage 2B complete — Cohort valued and purchase accounting added**

Cohort is the preferred target in the evidence-backed screen, scoring 85 versus 74 for Avon Technologies and 70 for Chemring.

The third Excel model includes a simple five-year Cohort forecast, DCF, synergy-supported maximum price, purchase-price allocation and offer-price sensitivity. The 30% premium case is approximately 0.1% EPS accretive after purchase-accounting amortisation but approximately 0.5% dilutive in Year 1 after integration costs.

The next stage is to consolidate the findings into the final recommendation and interview explanation.

Key files:

- `research/04_cohort_target_profile.md`
- `research/05_avon_target_profile.md`
- `research/06_target_screen.md`
- `research/07_cohort_forecast_and_valuation.md`
- `research/08_purchase_accounting_and_sensitivity.md`
- `model/build_model.mjs`

## Disclaimer

This is an independent educational project based on a hypothetical transaction. It is not investment advice and is not affiliated with BAE Systems or any target company.
