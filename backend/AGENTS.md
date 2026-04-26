# Project instructions

This project contains a Korean e-commerce keyword and title generator.

Current source-of-truth files for keyword/title behavior:
- backend/app/services/legacy_core.py
- backend/app/services/keyword_builder.py
- backend/app/services/market_keywords.py
- backend/app/services/pipeline.py

Core principles for this patch series:
1. `evidence-first`: use product name first, then OCR/Vision, then search-data hints.
2. `no global filler`: do not inject category-crossing filler such as shared segment pools or slot hints.
3. `no typo-expansion`: do not create intentional typos, spelling variants, or coverage-only noise.
4. `base_name-centered topic filtering`: anchors and baseline must come from the base product identity, not expanded final titles.
5. If evidence is weak, shorter output is better than unrelated expansion.
6. Never invent unsupported material, function, target user, quantity, or use-case.
7. Update regression tests when changing keyword or market-title behavior.

Practical guardrails:
- Keep the main identity tokens at the front of titles.
- Do not add new product groups or category labels just to increase length.
- Treat stage2 as prune/reorder/refine, not as a second expansion pass.
- Treat B-market outputs as subsets/refinements of validated A-market evidence.
