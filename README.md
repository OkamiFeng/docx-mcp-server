# Docx MCP Server 📄

[![Python](https://img.shields.io/badge/python-3.10%2B-blue)](https://www.python.org/)
[![MCP](https://img.shields.io/badge/MCP-Protocol-green)](https://modelcontextprotocol.io/)

A Model Context Protocol (MCP) server that enables AI agents to read, edit, and create Microsoft Word documents (`.docx`). It supports rich text, tables, and images, and provides flexible deployment options (Standalone Executable, Python/UVX, SSE/Remote).

## ✨ Features

- **Document Management**: Create new documents or load existing ones.
- **Content Reading**: Read document structure, paragraphs, tables, and **images**.
- **Rich Editing**:
  - Add stylized paragraphs (Bold, Italic handled via styles).
  - Add Headings (Levels 1-9).
  - **Insert Images** (from local path or Base64).
  - **Create Tables** with custom data.
- **Image Extraction**: Extract images from docx files to local disk.
- **Multi-Mode Support**:
  - **StdIO**: Standard integration for local MCP clients (e.g., Claude Desktop).
  - **SSE**: Server-Sent Events support for remote or URL-based connections.

## 🚀 Quick Start

### Option 1: Standalone Explorer (No Python Required)

Download the latest release and run the executable directly.

1. Get `docx-mcp-server.exe` from the [Releases](#) page.
2. Run in terminal or configure in your MCP client.

### Option 2: Using UVX (Requires PyPI Publication)

> ⚠️ **Note**: This package is **not** currently published to PyPI by the author. 
> To use `uvx`, you must fork and publish it yourself, or install locally.

If published to PyPI, you could run:
```bash
uvx docx-mcp-server
```

### Option 3: Development / Local Source

```bash
# Clone the repository
git clone https://github.com/OkamiFeng/docx-mcp-server.git
cd docx-mcp-server

# Run with uv
uv run docx-mcp-server
```

---

## 🛠️ Configuration

### Claude Desktop

#### Using UVX (If published to PyPI)
```json
{
  "mcpServers": {
    "docx": {
      "command": "uvx",
      "args": ["docx-mcp-server"]
    }
  }
}
```

#### Windows (using .exe)
```json
{
  "mcpServers": {
    "docx": {
      "command": "C:/Path/To/docx-mcp-server.exe",
      "args": []
    }
  }
}
```

#### Windows (using Python/UV)
```json
{
  "mcpServers": {
    "docx": {
      "command": "uv",
      "args": ["run", "--directory", "C:/Projects/docx-mcp-server", "docx-mcp-server"]
    }
  }
}
```

### Remote / SSE Configuration

To run the server in HTTP mode (accessible via URL):

```bash
# Using Exe
docx-mcp-server.exe --transport sse --port 8000

# Using Python
uv run server.py --transport sse --port 8000
```

**SSE Endpoint**: `http://localhost:8000/sse`

---

## 📚 Tools Available

| Tool Name | Description |
| :--- | :--- |
| `create_new_document` | Create a fresh empty document. |
| `load_document` | Load an existing `.docx` file from a path. |
| `get_document_structure` | Get a simplified JSON-like view of paragraphs and content counts. |
| `add_paragraph` | Add text paragraph with optional style. |
| `add_heading` | Add a heading (Level 1-9). |
| `add_table` | Insert a table with specified rows, cols, and data. |
| `add_image` | Insert an image from a file path or Base64 string. |
| `extract_images` | Extract all images from the loaded document to a folder. |
| `save_document` | Save changes (overwrite or new path). |

## 🏗️ Development

1. **Clone**: `git clone https://github.com/OkamiFeng/docx-mcp-server.git`
2. **Install**: `uv sync` or `pip install -r requirements.txt`
3. **Build Exe**:
   ```bash
   pyinstaller --name docx-mcp-server --onefile server.py --icon="icon.png" --collect-all docx
   ```

## 📄 License

MIT
