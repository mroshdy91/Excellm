"""Main entry point for excellm package."""

from .server import mcp

if __name__ == "__main__":
    mcp.run(transport="stdio")
