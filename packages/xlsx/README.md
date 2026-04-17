# @officeai/xlsx

**Status: deferred.** XLSX support (incl. ~150-function formula engine) is scoped
in [`spec/xlsx/`](../../spec/xlsx/) and will be implemented in a follow-up
session per [`prompt.md`](../../prompt.md). The core architecture
(command bus, agent API, OOXML utilities) is built format-agnostically in
[`@officeai/core`](../core) so this package can be filled in without
disturbing DOCX.
