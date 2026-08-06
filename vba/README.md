# VBA modules

Source lives here as importable `.bas` text files — never only inside a binary workbook. GitHub can render, diff and review text; it can't see inside an `.xlsm`.

## Import

1. Open the VBA editor (`Alt+F11`)
2. File → Import File… → pick the `.bas` module
3. Save the workbook as `.xlsm`

Both modules are self-contained: no library references to add (`modReconCompare` late-binds `Scripting.Dictionary`). That binding also pins `modReconCompare` to Windows Excel — Mac Excel has no `Scripting.Dictionary` and no reference can supply it. `modWorkpaperFormat` runs on both.

## Modules

| Module | What it does |
|---|---|
| `modWorkpaperFormat` | Standard workpaper header block, reviewer sign-off line, accounting number format, freeze panes |
| `modReconCompare` | Keyed two-way reconciliation between two (key, amount) ranges with tolerance — subledger vs GL pattern. It replaces only a prior result sheet that it marked itself; a user-created sheet named `Recon Result` is left untouched and the macro stops with instructions |

## Contributing your own

Export a module as text before committing: right-click the module in the VBE → Export File… Keep `Option Explicit` on and note any required references in the module header comment.
