# n8n-nodes-exceljs

This is an [n8n](https://n8n.io) community node. It lets you read, create and manipulate
Excel (`.xlsx`) files directly in your workflows, powered by the
[ExcelJS](https://github.com/exceljs/exceljs) library.

Excel files are passed between nodes as binary data, so the ExcelJS node composes well
with HTTP Request, Read/Write Files from Disk, Email (IMAP/SMTP), Google Drive, and any
other node that handles binaries.

## Installation

### Community Nodes (recommended)

1. In your n8n instance open **Settings → Community nodes**.
2. Select **Install a community node**.
3. Enter `n8n-nodes-exceljs` and confirm.

### Manual installation (self-hosted)

```bash
cd ~/.n8n/nodes
npm install n8n-nodes-exceljs
```

Restart n8n afterwards. For Docker, mount the package into the container and set
`N8N_CUSTOM_EXTENSIONS`, or build a custom image.

## Operations

The node has one **Resource** selector with the following operations:

### Workbook
| Operation | Description |
|---|---|
| Create | Create a new empty workbook |
| Get Info | Get workbook properties and the sheet list |
| Set Properties | Set author, title, subject, company |

### Worksheet
| Operation | Description |
|---|---|
| Add / Remove / Rename | Manage sheets |
| Get All | List all worksheets |
| Protect / Unprotect | Sheet protection with a password |
| Freeze Panes | Freeze rows and/or columns |
| Set Auto Filter | Enable auto filter on a range |
| Set Properties | Tab color, visibility (incl. very hidden), zoom |
| Set Page Setup | Orientation, paper size, fit to page |

### Cell
| Operation | Description |
|---|---|
| Read Value / Read Hyperlink | Read a cell value or hyperlink |
| Write Value | Write string, number, boolean or date |
| Write Formula | Write a formula (without leading `=`) |
| Set Style | Font, fill, borders, alignment, wrap text |
| Merge Cells / Unmerge Cells | Merge or unmerge a range |
| Add Hyperlink | Add a hyperlink with optional display text |
| Set Number Format | e.g. `#,##0.00`, `dd.mm.yyyy`, `0.00%` |

### Row
| Operation | Description |
|---|---|
| Add Row | Append a single row from a JSON array |
| Add Rows | Append many rows from input items (optional header) |
| Get Row | Read a row by number |
| Set Height / Hide / Show | Row appearance |

### Column
| Operation | Description |
|---|---|
| Set Width / Hide / Show | Column appearance |

### Range
| Operation | Description |
|---|---|
| Write Data | Write input items to a range (optional header) |
| Read Data | Read a range into JSON items (first row as header, or `col_N` keys) |
| Read Hyperlinks | Extract hyperlinks from a range |
| Set Style | Apply bold / fill / borders to a range |

Ranges accept a closed form `A1:D10` or an open form `A1:D` (down to the last row)
for the read operations.

### Image
| Operation | Description |
|---|---|
| Add Image | Embed a PNG/JPEG/GIF from binary data, anchored by cells |

### Conditional Formatting
| Operation | Description |
|---|---|
| Add Rule | `cellIs` (greater than, less than, equal, between) or `containsText` rules with a highlight fill |

## Formula hyperlinks

When writing rows from items (Row → Add Rows, Range → Write Data), a string value that
looks like `=HYPERLINK("https://example.com", "Example")` is converted into a real Excel
hyperlink with link styling automatically.

## Compatibility

- Requires n8n 1.x or later (community nodes API v1).
- The package bundles ExcelJS — no additional runtime dependencies are installed.

## Development

```bash
npm install
npm run build      # typecheck + esbuild bundle into dist/
```

To test locally, link or copy the package into `~/.n8n/nodes` and restart n8n.

Releasing: bump the version with `npm version patch|minor|major` and push the tag —
GitHub Actions (.github/workflows/publish.yml) builds and publishes the package to npm
with a provenance statement.

## License

[MIT](LICENSE)
