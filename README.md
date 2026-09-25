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

**Stage 1 complete — Cohort selected for detailed modelling**

Cohort is the preferred target in the evidence-backed screen, scoring 85 versus 74 for Avon Technologies and 70 for Chemring.

The first Excel model has been built. It includes editable transaction assumptions, the target screen, offer and sources & uses, illustrative earnings impact, leverage and source documentation. It remains an initial transaction calculator: the next stage is a full Cohort forecast, valuation, purchase accounting and sensitivity analysis.

Key files:

- `research/04_cohort_target_profile.md`
- `research/05_avon_target_profile.md`
- `research/06_target_screen.md`
- `model/build_model.mjs`

## Disclaimer

This is an independent educational project based on a hypothetical transaction. It is not investment advice and is not affiliated with BAE Systems or any target company.
