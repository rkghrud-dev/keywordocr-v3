# Project instructions

This project contains a Korean e-commerce keyword and title generator.

Before changing any keyword generation or market-facing title logic, read:
- docs/coupang_title_rule_prompt.md

When working on this feature, follow these rules:

1. Treat the analyzed Coupang H-column style as the target market-facing title style.
2. When a source/internal title is available, classify the input as `KEEP` or `EXPAND` first.
3. `KEEP` when the original title is already a clear consumer-facing product title; in that case, mostly remove internal codes and obvious noise only.
4. `EXPAND` when the original title is compressed, internal, ambiguous, or too short for search exposure.
5. For `EXPAND`, build titles in this order when relevant:
   - use scene / category / place
   - unpacked core product term
   - close synonyms / spacing variants / alternate search terms
   - function / benefit / problem-solving keywords
   - target user / usage context
   - material / shape / spec / quantity only when justified
   - option at the end only
6. Remove internal codes such as `GS0600731A`.
7. Never invent materials, functions, quantities, or components that are not supported by the input.
8. Output must be exactly one Korean product-title line, not bullets, JSON, or explanation.
9. Match the existing Coupang-style long search-oriented naming pattern described in `docs/coupang_title_rule_prompt.md`.
10. Do not over-expand with unrelated categories or emotional marketing words.
11. When editing code, also add or update tests/examples for `KEEP` vs `EXPAND` behavior.
