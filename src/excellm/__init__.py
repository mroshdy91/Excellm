"""Excel Live MCP Server - Real-time Excel automation for LLMs.

This package provides MCP (Model Context Protocol) tools for interacting
with open Excel files using Windows COM automation.

Version: 1.0.0
Author: Excel Live Injector Team
"""

__version__ = "1.0.0"
__author__ = "Excel Live Injector Team"

from .server import create_server

__all__ = ["create_server", "__version__"]
