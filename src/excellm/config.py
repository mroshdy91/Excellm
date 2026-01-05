"""Configuration settings for ExceLLM MCP server.

Centralized environment-based configuration to control features and security.
"""

import os

# =============================================================================
# SECURITY SETTINGS
# =============================================================================

# VBA Execution: DISABLED by default for security
# VBA code runs with full Excel/COM privileges and can modify files,
# access system resources, and execute arbitrary code.
# Set EXCELLM_ENABLE_VBA=true in environment to enable.
VBA_ENABLED = os.getenv("EXCELLM_ENABLE_VBA", "").lower() in ("true", "1", "yes")


# =============================================================================
# OPERATIONAL SETTINGS
# =============================================================================

# Logging level (DEBUG, INFO, WARNING, ERROR)
LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO")

# Maximum cells per write operation (hallucination prevention)
MAX_CELLS_LIMIT = int(os.getenv("EXCELLM_MAX_CELLS", "250"))

# Session expiry time in hours
SESSION_EXPIRY_HOURS = int(os.getenv("EXCELLM_SESSION_EXPIRY_HOURS", "1"))
