# PDF — Keyboard Shortcuts

> Full catalogue. Every action reachable from the keyboard. Vim-mode
> opt-in via cookie (`pdf-vim-mode=1`).

Cross-references: accessibility contract in
[`accessibility.md`](./accessibility.md);
text-layer selection in [`text-layer.md`](./text-layer.md);
search in [`search.md`](./search.md);
dark mode in [`dark-mode.md`](./dark-mode.md).

## Conventions

- `Ctrl` on Linux/Windows ≡ `Cmd` on macOS unless noted.
- `Alt` on Linux/Windows ≡ `Option` on macOS.
- Shortcuts not active when a text input has focus (Find, GotoDialog,
  free-text annotation, comment composer) — typing flows to the input.
- Vim-mode bindings are listed in the last column; they override
  default bindings on conflict.

## Navigation

| Action                          | Default              | Vim mode      |
| ------------------------------- | -------------------- | ------------- |
| Next page                       | `PageDown`, `Space`  | `j`           |
| Previous page                   | `PageUp`, `Shift+Space` | `k`         |
| Scroll up (line)                | `Up`                 | `Ctrl+u`      |
| Scroll down (line)              | `Down`               | `Ctrl+d`      |
| Scroll left                     | `Left`               | `h`           |
| Scroll right                    | `Right`              | `l`           |
| First page                      | `Home`               | `gg`          |
| Last page                       | `End`                | `G`           |
| Jump to page                    | `Ctrl+G`             | `:N` then `Enter` |
| Back in history                 | `Alt+Left`           | `Ctrl+o`      |
| Forward in history              | `Alt+Right`          | `Ctrl+i`      |
| Next bookmark / outline entry   | `Alt+Down`           | `]b`          |
| Previous bookmark / outline     | `Alt+Up`             | `[b`          |

## Zoom and view

| Action                          | Default              | Vim mode      |
| ------------------------------- | -------------------- | ------------- |
| Zoom in                         | `Ctrl++`             | `+`           |
| Zoom out                        | `Ctrl+-`             | `-`           |
| Reset zoom (100%)               | `Ctrl+0`             | `=`           |
| Fit width                       | `Ctrl+1`             | `zw`          |
| Fit page                        | `Ctrl+2`             | `zp`          |
| Fit actual size                 | `Ctrl+3`             | `za`          |
| Single-page view                | `Ctrl+Alt+1`         |               |
| Continuous view                 | `Ctrl+Alt+2`         |               |
| Two-up view (book)              | `Ctrl+Alt+3`         |               |
| Two-up view (cover)             | `Ctrl+Alt+4`         |               |
| Rotate page clockwise           | `Ctrl+Alt+R`         | `]r`          |
| Rotate page counter-clockwise   | `Ctrl+Alt+Shift+R`   | `[r`          |
| Toggle reflow mode              | `Ctrl+Shift+R`       | `zr`          |
| Toggle dark mode                | `Ctrl+Alt+D`         | `zd`          |
| Toggle distraction-free / chrome| `Esc`                | `zz`          |

## Search

| Action                          | Default              | Vim mode      |
| ------------------------------- | -------------------- | ------------- |
| Open find bar                   | `Ctrl+F`             | `/`           |
| Reverse-search prompt           |                      | `?`           |
| Find next                       | `Ctrl+G`, `F3`       | `n`           |
| Find previous                   | `Ctrl+Shift+G`, `Shift+F3` | `N`     |
| Toggle case sensitivity         | `Alt+C` (with find open) |           |
| Toggle whole word               | `Alt+W` (with find open) |           |
| Toggle regex                    | `Alt+R` (with find open) |           |
| Close find bar                  | `Esc`                |               |

## Sidebar

| Action                          | Default              | Vim mode      |
| ------------------------------- | -------------------- | ------------- |
| Toggle sidebar                  | `Ctrl+\`             | `zs`          |
| Outline tab                     | `Ctrl+Shift+O`       | `go`          |
| Thumbnails tab                  | `Ctrl+Shift+T`       | `gt`          |
| Annotations tab                 | `Ctrl+Shift+A`       | `ga`          |
| Search tab                      | `Ctrl+Shift+F`       | `gf`          |
| Comments tab                    | `Ctrl+Shift+C`       | `gc`          |
| Attachments tab                 | `Ctrl+Shift+L`       | `gl`          |

## Annotations

| Action                          | Default              | Vim mode      |
| ------------------------------- | -------------------- | ------------- |
| Highlight selected text         | `Ctrl+Shift+H`       | `mh`          |
| Underline selected text         | `Ctrl+Shift+U`       | `mu`          |
| Strikethrough selected text     | `Ctrl+Shift+X`       | `mx`          |
| Squiggly underline selected     | `Ctrl+Shift+Q`       | `mq`          |
| Add sticky note (at click point)| `Ctrl+Alt+N`         | `mn`          |
| Add free-text annotation        | `Ctrl+Alt+T`         | `mt`          |
| Activate ink (free-hand) tool   | `Ctrl+Alt+I`         | `mi`          |
| Activate rectangle tool         | `Ctrl+Alt+S`         | `ms`          |
| Activate redaction tool (mark)  | `Ctrl+Alt+X`         | `mx`          |
| Apply redactions                | `Ctrl+Alt+Shift+X`   |               |
| Add comment at focused page     | `Ctrl+Alt+C`         | `mc`          |
| Reply to focused comment        | `R`                  | `R`           |
| Resolve focused comment         | `E` (with comment focused) | `E`     |
| Delete focused annotation       | `Delete`, `Backspace`| `dd`          |
| Open hyperlink popover (focused link annotation) | `Enter` |       |

## Forms

| Action                          | Default              |
| ------------------------------- | -------------------- |
| Next field                      | `Tab`                |
| Previous field                  | `Shift+Tab`          |
| Toggle checkbox / radio         | `Space`              |
| Open combo / list dropdown      | `Alt+Down`           |
| Reset form                      | `Ctrl+Alt+Shift+R`   |
| Flatten form                    | `Ctrl+Alt+Shift+F`   |

## Page operations

| Action                          | Default              |
| ------------------------------- | -------------------- |
| Insert page (open dialog)       | `Ctrl+Alt+P`         |
| Delete current page             | `Ctrl+Alt+Backspace` |
| Move current page up            | `Ctrl+Alt+Up`        |
| Move current page down          | `Ctrl+Alt+Down`      |
| Duplicate current page          | `Ctrl+Alt+J`         |
| Open page properties            | `Ctrl+Alt+E`         |

## Editing & history

| Action                          | Default              |
| ------------------------------- | -------------------- |
| Undo                            | `Ctrl+Z`             |
| Redo                            | `Ctrl+Shift+Z`, `Ctrl+Y` |
| Open command palette            | `Ctrl+K`             |
| Approve focused pending mutation| `Ctrl+Enter`         |
| Reject focused pending mutation | `Ctrl+Shift+Enter`   |

## Selection & copy

| Action                          | Default              |
| ------------------------------- | -------------------- |
| Select all (current page)       | `Ctrl+A`             |
| Select all (full document)      | `Ctrl+Shift+A`       |
| Copy as plain text              | `Ctrl+C`             |
| Copy as Markdown                | `Ctrl+Shift+C`       |
| Copy as HTML                    | `Ctrl+Alt+C`         |
| Lasso (rectangular) selection   | hold `Alt` while drag |
| Translate selection             | `Ctrl+Shift+T` *(when selected)* |

## File / sharing

| Action                          | Default              |
| ------------------------------- | -------------------- |
| Open file                       | `Ctrl+O`             |
| Save (download modified)        | `Ctrl+S`             |
| Save as (download original)     | `Ctrl+Shift+S`       |
| Print                           | `Ctrl+P`             |
| Document properties             | `Ctrl+Alt+I`         |
| Share link                      | `Ctrl+Alt+H`         |

## Accessibility shortcuts

| Action                          | Default              |
| ------------------------------- | -------------------- |
| Skip to toolbar                 | `Ctrl+,`             |
| Skip to sidebar                 | `Ctrl+1` *(when sidebar visible)* |
| Skip to page content            | `Ctrl+2`             |
| Cycle focus across landmarks    | `F6`                 |
| Reverse-cycle landmarks         | `Shift+F6`           |
| Show keyboard shortcut help     | `Ctrl+/`             |

## Vim-mode specifics

Activated by setting cookie `pdf-vim-mode=1` (or via the page menu →
"Vim mode"). On activation:

- A small `vim` indicator appears in the status bar.
- Pending command buffer (`gg`, `mh`, `:42`, etc.) is shown in the
  bottom-left like Vim's mode line.
- `:N` followed by `Enter` jumps to page N.
- `:q` closes the document (returns to home).
- `:set nu` toggles a per-page line-number gutter (text-layer-line
  numbers).
- `Esc` always cancels the pending command.

## Conflict resolution

Where defaults overlap with browser shortcuts (e.g. `Ctrl+G` is
"Find next" in Chrome), we **only** intercept when the viewer has
focus. The browser's own search (`Ctrl+F` opening Chrome's bar) is
shadowed by ours when the page focus is on the viewer surface.
`Ctrl+P` is intercepted to surface our print dialog (with PDF-aware
options) and falls through to browser print if the user dismisses
ours.

## Help discoverability

`Ctrl+/` opens the in-product **Keyboard shortcut cheatsheet** — a
modal listing every shortcut above, searchable, printable. The
cheatsheet is generated from this spec so it can never drift.
