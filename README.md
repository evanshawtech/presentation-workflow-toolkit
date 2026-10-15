# Presentation Workflow Toolkit

A small local workflow layer for presentation planning, renderer selection and evidence checks. It contains selected project-specific helper code; it is **not** the upstream rendering engines or a complete installed presentation environment.

## Problem and approach

A presentation can pass a structural check while an exported deck remains stale or unreadable. This layer keeps a page plan, chooses one primary renderer, performs static PPTX checks, optionally renders with Microsoft PowerPoint, and checks that the final visual-review record refers to the current file hash.

The included helpers implement page-plan validation, hash-bound QA gates, grouped-text/CJK reference measurements, a process-local guard for selected image-size v1 format handlers, and a PowerPoint renderer that serializes concurrent requests and avoids closing presentations it does not own.

## Reading and using the original helpers

```sh
python router/toolkit.py validate-plan examples/page-plan.json
python router/toolkit.py --help
python -B tests/test_gates.py
```

Plan validation uses only the standard library. PPTX analysis/measurement additionally needs packages in `requirements.txt`. Actual PowerPoint rendering requires Windows and a licensed Microsoft PowerPoint installation. No installer or automatic upstream checkout is included.

`router/toolkit.py route` expects independently installed, correctly licensed engines beneath `skills-src/`. Its expected paths are in `RENDERERS`; missing engines cause a clear failure or documented fallback. Consult the upstream projects linked in [THIRD_PARTY.md](THIRD_PARTY.md), including their current workflow and licensing constraints, before using them. This snapshot does not redistribute those engines or override their required approval gates.

`extract-design-dna.py` measures an approved local reference PPTX; it does not upload that deck. The image guard is an opt-in Node preload meant for compatible installed dependency layouts, not a security scanner or a promise that arbitrary images are safe. `render-powerpoint.ps1` may launch PowerPoint when invoked; do not treat static validation as render proof.

## Evidence

The original installed-workflow report records 12 routing/plan/final-gate checks and five rendered example decks. Those decks, screenshots, engine installations and machine-specific reports are **not** shipped here; the report does not prove this compact snapshot can reproduce them without external setup. R5's tests cover plan validity and stale-evidence rejection on synthetic fixtures only. Actual rendering and provider/LLM calls: **NOT_RUN**.

## Limitations

This is a selected workflow-layer snapshot. It is not an end-to-end slide generator, a vendor-neutral installer, or a claim to have authored third-party renderers. Hash-bound review cannot replace viewing every slide. Font substitution, PowerPoint COM failures and upstream API changes remain possible. The original reference-measurement helper is described as an extension of upstream methodology; its contribution boundary still requires confirmation before publication.

## License

Original helper software and associated documentation are available under the [MIT License](LICENSE). Upstream engines, methods, images, fonts and commercial integration code are excluded and retain their own rights; see [THIRD_PARTY.md](THIRD_PARTY.md).
