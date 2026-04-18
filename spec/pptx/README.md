# PPTX Spec

Per [`prompt.md`](../../prompt.md). Mirrors the structure of [`spec/docx/`](../docx).

| Doc                                                  | What it answers                                              |
| ---------------------------------------------------- | ------------------------------------------------------------ |
| [`analysis.md`](./analysis.md)                       | Step A — clean-room study of references and OOXML            |
| [`feature-scope.md`](./feature-scope.md)             | P0/P1/P2/OUT for every prompt-listed feature                 |
| [`document-model.md`](./document-model.md)           | Typed `PptxPresentation`, `Slide`, `Shape` union, `TextBody` |
| [`ooxml-mapping.md`](./ooxml-mapping.md)             | Every typed model node ↔ OOXML element                       |
| [`parser.md`](./parser.md)                           | Algorithm, errors, determinism                               |
| [`serializer.md`](./serializer.md)                   | Dirty-flag-driven, byte-preserving                           |
| [`renderer.md`](./renderer.md)                       | Hybrid SVG + HTML overlay; pure layout fns                   |
| [`agent-commands.md`](./agent-commands.md)           | Ten commands, full payloads + diffs + errors                 |
| [`edge-cases.md`](./edge-cases.md)                   | Slide-id accounting, theme color, placeholder inheritance    |
| [`acceptance-criteria.md`](./acceptance-criteria.md) | Measurable bar for declaring the phase done                  |

The shared specs in [`../shared/`](../shared) cover PPTX-applicable
infrastructure (command bus, OOXML utils, agent API, plugin system).
