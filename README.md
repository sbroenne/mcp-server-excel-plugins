# ExcelMcp Copilot CLI Plugins

Windows-only GitHub Copilot CLI plugins for ExcelMcp.

This repository is the publish target for plugin artifacts from [`sbroenne/mcp-server-excel`](https://github.com/sbroenne/mcp-server-excel).

## Plugins

- **excel-mcp** — MCP server plugin for conversational Excel automation
- **excel-cli** — CLI skill plugin for scripting and batch workflows

## Repository Layout

```text
plugins/
├── excel-mcp/
└── excel-cli/
marketplace.json
```

`marketplace.json` is the marketplace manifest. The `plugins/` directory contains the distributable plugin templates and published content used by the source repo's `publish-plugins.yml` workflow.

## Install

```powershell
# Register this marketplace
copilot plugin marketplace add sbroenne/mcp-server-excel-plugins

# Install one or both plugins
copilot plugin install excel-mcp@mcp-server-excel
copilot plugin install excel-cli@mcp-server-excel
```

## Notes

- **Windows only** — ExcelMcp depends on Microsoft Excel COM automation.
- **excel-mcp** downloads the MCP server binary separately via its bundled helper scripts.
- **excel-cli** expects `excelcli.exe` to be installed separately.

## Source and Support

- Source repo: [sbroenne/mcp-server-excel](https://github.com/sbroenne/mcp-server-excel)
- Issues: [sbroenne/mcp-server-excel/issues](https://github.com/sbroenne/mcp-server-excel/issues)
- Plugin docs: [`plugins/excel-mcp/README.md`](./plugins/excel-mcp/README.md), [`plugins/excel-cli/README.md`](./plugins/excel-cli/README.md)

## License

MIT
