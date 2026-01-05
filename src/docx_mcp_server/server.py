import sys
import argparse
import uvicorn
from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.server.sse import SseServerTransport
from mcp.types import Tool, TextContent, ImageContent, EmbeddedResource

from starlette.applications import Starlette
from starlette.routing import Route, Mount
import mcp.types as types
from .docx_manager import DocxManager
import traceback

# Initialize Manager
manager = DocxManager()

# Create Server instance
server = Server("docx-mcp-server")

async def run_stdio():
    async with stdio_server() as (read_stream, write_stream):
        await server.run(read_stream, write_stream, server.create_initialization_options())

# ... (tools omitted, they remain the same) ...

# ... (rest of tools) ...

# At the bottom, update main execution
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--transport", default="stdio", choices=["stdio", "sse"])
    parser.add_argument("--port", type=int, default=8000)
    args = parser.parse_args()

    if args.transport == "sse":
        from starlette.applications import Starlette
        from starlette.routing import Route
        from mcp.server.sse import SseServerTransport
        import uvicorn
        
        sse = SseServerTransport("/messages")
        
        async def sse_endpoint(request):
            async with sse.connect_sse(request.scope, request.receive, request._send) as streams:
                await server.run(streams[0], streams[1], server.create_initialization_options())

        async def messages_endpoint(request):
            await sse.handle_post_message(request.scope, request.receive, request._send)

        app = Starlette(
            routes=[
                Route("/sse", endpoint=sse_endpoint, methods=["GET"]),
                Route("/messages", endpoint=messages_endpoint, methods=["POST"]),
            ]
        )
        
        print(f"Starting SSE server on port {args.port}...")
        uvicorn.run(app, host="0.0.0.0", port=args.port)
        
    else:
        # Run STDIO
        import asyncio
        asyncio.run(run_stdio())

if __name__ == "__main__":
    main()

# --- Tool Definitions ---

@server.list_tools()
async def handle_list_tools() -> list[types.Tool]:
    return [
        types.Tool(
            name="create_new_document",
            description="Create a new empty Docx document.",
            inputSchema={"type": "object", "properties": {}}
        ),
        types.Tool(
            name="load_document",
            description="Load an existing Docx document.",
            inputSchema={
                "type": "object", 
                "properties": {
                    "path": {"type": "string", "description": "Absolute path to the .docx file"}
                },
                "required": ["path"]
            }
        ),
        types.Tool(
            name="save_document",
            description="Save the current document.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Optional save path (Save As)"}
                }
            }
        ),
        types.Tool(
            name="get_document_structure",
            description="Get the structure of the current document.",
            inputSchema={"type": "object", "properties": {}}
        ),
        types.Tool(
            name="add_paragraph",
            description="Add a text paragraph.",
            inputSchema={
                "type": "object",
                "properties": {
                    "text": {"type": "string"},
                    "style": {"type": "string", "description": "Style name, e.g. 'Normal', 'Heading 1'"}
                },
                "required": ["text"]
            }
        ),
        types.Tool(
            name="add_heading",
            description="Add a heading.",
            inputSchema={
                "type": "object",
                "properties": {
                    "text": {"type": "string"},
                    "level": {"type": "integer", "default": 1}
                },
                "required": ["text"]
            }
        ),
        types.Tool(
            name="add_table",
            description="Add a table.",
            inputSchema={
                "type": "object",
                "properties": {
                    "rows": {"type": "integer"},
                    "cols": {"type": "integer"},
                    "data": {
                        "type": "array", 
                        "items": {"type": "array", "items": {"type": "string"}},
                        "description": "2D list of strings"
                    }
                },
                "required": ["rows", "cols"]
            }
        ),
        types.Tool(
            name="add_image",
            description="Add an image.",
            inputSchema={
                "type": "object",
                "properties": {
                    "image_source": {"type": "string", "description": "Path or Base64 string"},
                    "width_inches": {"type": "number"}
                },
                "required": ["image_source"]
            }
        ),
        types.Tool(
            name="extract_images",
            description="Extract images from the current document to a directory.",
            inputSchema={
                "type": "object",
                "properties": {
                    "output_dir": {"type": "string", "description": "Directory to save extracted images"}
                },
                "required": ["output_dir"]
            }
        )
    ]

@server.call_tool()
async def handle_call_tool(name: str, arguments: dict | None) -> list[types.TextContent | types.ImageContent | types.EmbeddedResource]:
    if arguments is None:
        arguments = {}
    
    try:
        result = ""
        if name == "create_new_document":
            result = manager.create_new()
        elif name == "load_document":
            result = manager.load_document(arguments.get("path"))
        elif name == "save_document":
            result = manager.save_document(arguments.get("path"))
        elif name == "get_document_structure":
            result = str(manager.get_structure())
        elif name == "add_paragraph":
            result = manager.add_paragraph(arguments.get("text"), arguments.get("style"))
        elif name == "add_heading":
            result = manager.add_heading(arguments.get("text"), arguments.get("level", 1))
        elif name == "add_table":
            result = manager.add_table(arguments.get("rows"), arguments.get("cols"), arguments.get("data"))
        elif name == "add_image":
            result = manager.add_image(arguments.get("image_source"), arguments.get("width_inches"))
        elif name == "extract_images":
            paths = manager.extract_images(arguments.get("output_dir"))
            result = f"Extracted {len(paths)} images: {paths}"
        else:
            raise ValueError(f"Unknown tool: {name}")
            
        return [types.TextContent(type="text", text=str(result))]
    except Exception as e:
        return [types.TextContent(type="text", text=f"Error: {str(e)}\n{traceback.format_exc()}")]


