import os
import sys
import functools
from .tools import mcp


def _enhanced_run_wrapper(original_run):
    """Wrapper around FastMCP's run method to add custom error handling."""
    
    @functools.wraps(original_run)
    def wrapper(*args, **kwargs):
        # Add any pre-run setup or configuration here
        try:
            return original_run(*args, **kwargs)
        except Exception as e:
            # Log the error for debugging
            print(f"FastMCP server error: {e}", file=sys.stderr)
            # Re-raise to maintain original behavior
            raise
    
    return wrapper


def main() -> None:
    if not os.getenv("MICROSOFT_MCP_CLIENT_ID"):
        print(
            "Error: MICROSOFT_MCP_CLIENT_ID environment variable is required",
            file=sys.stderr,
        )
        sys.exit(1)

    # Wrap the run method with our enhanced error handling
    mcp.run = _enhanced_run_wrapper(mcp.run)
    
    # Run the server
    mcp.run()


if __name__ == "__main__":
    main()
