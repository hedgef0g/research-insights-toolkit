# Decisions

## 2026-04-28

The first significance MVP assumes:
- input values are proportions/percentages;
- bases are below values;
- selection is a 2x2 range;
- significance is calculated using pooled z-test;
- confidence level is fixed at 95%;
- Excel wrapper is only responsible for reading selection and displaying result;
- statistical logic lives in src/core/significance.js.