# excel-engine-mcp

[한국어 안내 → README.ko.md](README.ko.md)

A local MCP server that drives the **real Excel desktop application** (COM automation) to do the things file-manipulation libraries (openpyxl etc.) fundamentally cannot: **recalculate formulas, read computed cell values, and render true-fidelity PDF**.

> Design philosophy: this server **complements** library workflows instead of replacing them.
> Writing static cells and bulk data is faster with openpyxl — but openpyxl only stores formula
> *strings* (reading them back gives `None`), and no library has Excel's render engine.
> This MCP handles only the engine-exclusive part, keeping its tool surface tiny (5 tools).

## Requirements

- Windows 10+ with a **logged-in interactive desktop session** (Excel has no true headless mode)
- Microsoft Office (Excel) installed and licensed — verified on **Office 2016+ (Excel 16.0)**
- Python 3.10+ — verified on **3.12**
- [Claude Code](https://claude.com/claude-code) or any MCP client

## Install

```powershell
git clone https://github.com/Feynman520/excel-engine-mcp.git
cd excel-engine-mcp
py -3.12 -m venv .venv          # or: python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

### Register with Claude Code

Run this in the cloned folder (uses absolute paths, so it works from anywhere afterwards):

```powershell
claude mcp add excel-automation --scope user -- "$PWD\.venv\Scripts\python.exe" "$PWD\server.py"
```

`--scope user` makes it available in every project. Use `--scope project` to limit it to one project.

### Verify

```powershell
$py = ".\.venv\Scripts\python.exe"; $env:PYTHONUTF8 = "1"
& $py tests\smoke_com.py      # COM recalc / computed values / PDF / zombie cleanup
& $py tests\server_tools.py   # MCP tool registration (does not launch Excel)
```

## Tools (3 core + 2 diagnostics)

| # | Tool | Input → Output | Why engine-only |
|---|---|---|---|
| ① | `excel_recalculate` | `src_path, out_path` → `{out_path}` | Full dependency recalculation — libraries write formulas but cannot compute them |
| ② | `excel_read_values` | `path, sheet?, cell_range?, recalc?` → `{sheet,address,rows,cols,values,text}` | Reads **computed** cell values (openpyxl only has the formula strings) |
| ③ | `excel_export_pdf` | `src_path, out_path, sheet?, recalc?` → `{out_path}` | WYSIWYG PDF via Excel's render engine (print areas, fit-to-page, headers/footers) |
| — | `excel_health` | → `{alive, excel_version}` | Session check (launches Excel on first call) |
| — | `excel_restart` | → `{alive}` | Recovery from COM errors |

Typical flow: write inputs with openpyxl → `excel_recalculate` to bake real values →
`excel_read_values` to verify / `excel_export_pdf` for the final PDF.
Originals are never modified; results are always written to `out_path`.

## Architecture notes

- **Single STA worker thread** (`engine/session.py`): win32com COM objects are bound to the
  thread that created them, while FastMCP may hop threads — so every Excel call is serialized
  onto one dedicated thread.
- **Lazy session**: Excel starts on the first tool call, is reused across calls, and is closed
  when the server stops.
- **`DispatchEx` + early binding (`gencache.EnsureDispatch`)**: a dedicated Excel instance
  (never attaches to the user's open Excel), called through the type library.
  ⚠️ `ExportAsFixedFormat` **fails with 1004 (0x800A03EC) under late binding** due to optional
  argument marshalling — early binding is required.
- **Zombie prevention**: after `Quit()`, any surviving dedicated Excel process is force-killed
  by PID (identified via the main window handle).
- **RPC-rejection retry**: `RPC_E_CALL_REJECTED` right after startup is retried with backoff.
- **stdout protection**: the one-time makepy cache generation output is redirected to stderr
  (stdout is the JSON-RPC channel).

## Limitations

- Not suitable for unattended/service sessions (needs an interactive desktop).
- Bulk static writes are faster with libraries — that is by design.

## License

[MIT](LICENSE)
