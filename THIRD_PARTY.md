# Upstream and dependency boundaries

No upstream source tree, commercial exporter, renderer binary, fonts, provider credentials, real reference deck or generated customer presentation is included. The expected optional engines are independently maintained:

- [ppt-master](https://github.com/hugohe3/ppt-master)
- [Dashi](https://github.com/chuspeeism/dashi-ppt-skill)
- [Academic Spark](https://github.com/wycmochi/cn-academic-spark)
- [presentation-skill](https://github.com/siril9/presentation-skill)
- [deck-dna](https://github.com/opitaru-sys/deck-dna)

These links identify integration and method sources, not endorsement or a new license grant. Dashi's integration-restricted export core and patches involving it are deliberately excluded. The helper for design measurements states that it extends the upstream counting method for grouped text and CJK; review the relevant contribution boundary rather than asserting ownership of upstream work.

The Python helpers import python-pptx, Pillow and lxml for selected functions. They are external packages, not vendored source. PowerPoint and image-size are external runtimes/dependencies; the helper does not relicense either. Install and use dependencies under their own current licenses.
