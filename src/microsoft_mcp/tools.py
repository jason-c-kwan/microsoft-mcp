import base64
import datetime as dt
import pathlib as pl
import urllib.parse
from typing import Any, Optional
from fastmcp import FastMCP
from . import graph, auth

# Color mapping for natural language color names to Outlook preset codes
COLOR_MAPPING = {
    # Standard colors
    "red": "preset0",
    "orange": "preset1",
    "brown": "preset2",
    "yellow": "preset3",
    "green": "preset4",
    "teal": "preset5",
    "olive": "preset6",
    "blue": "preset7",
    "purple": "preset8",
    "cranberry": "preset9",
    "steel": "preset10",
    "darksteel": "preset11",
    "gray": "preset12",
    "darkgray": "preset13",
    "black": "preset14",
    # Dark variants
    "darkred": "preset15",
    "darkorange": "preset16",
    "darkbrown": "preset17",
    "darkyellow": "preset18",
    "darkgreen": "preset19",
    "darkteal": "preset20",
    "darkolive": "preset21",
    "darkblue": "preset22",
    "darkpurple": "preset23",
    "darkcranberry": "preset24",
    # Aliases with spaces
    "dark red": "preset15",
    "dark orange": "preset16",
    "dark brown": "preset17",
    "dark yellow": "preset18",
    "dark green": "preset19",
    "dark teal": "preset20",
    "dark olive": "preset21",
    "dark blue": "preset22",
    "dark purple": "preset23",
    "dark cranberry": "preset24",
}

# Valid entity types based on Microsoft Graph API documentation
VALID_ENTITY_TYPES = {
    "event",
    "message",
    "driveItem",
    "externalItem",
    "site",
    "list",
    "listItem",
    "drive",
    "chatMessage",
    "person",
    "acronym",
    "bookmark",
}

mcp = FastMCP("microsoft-mcp")


def _normalize_request_params(request_data):
    """Normalize request parameters to ensure compatibility with MCP types.
    
    This function handles common parameter normalization issues that cause
    validation failures in built-in MCP methods.
    """
    if request_data is None:
        return None
    
    # Handle requests with missing params field
    if not hasattr(request_data, 'params') or request_data.params is None:
        # For ping and other methods that accept optional params,
        # ensure we return the correct None value
        return None
    
    # Handle requests with empty params dict
    if hasattr(request_data, 'params') and request_data.params == {}:
        # Convert empty dict to None for methods that expect optional params
        return None
    
    # Return the params as-is for methods that require specific parameter structures
    return request_data.params


# Enhanced built-in method handlers with robust parameter handling
def _create_enhanced_handlers():
    """Create enhanced handlers for built-in MCP methods with proper parameter validation."""
    
    # Import required types
    try:
        import mcp.types as types
        from mcp.shared.exceptions import McpError
        from mcp.server.lowlevel.server import _ping_handler
        
        async def enhanced_ping_handler(request=None):
            """Enhanced ping handler that properly handles all parameter variations.
            
            This handler accepts requests with no params, null params, or empty params
            and responds with a proper empty result.
            """
            try:
                # The ping request should accept None, empty dict, or valid RequestParams
                # No additional processing needed for ping - just return success
                return types.ServerResult(types.EmptyResult())
            except Exception as e:
                # Provide a more specific error response if something goes wrong
                return types.ServerResult(
                    types.ErrorData(
                        code=0,
                        message=f"Ping handler error: {str(e)}",
                        data=None
                    )
                )
        
        async def enhanced_list_tools_handler(request=None):
            """Enhanced tools/list handler with robust parameter handling."""
            try:
                # Delegate to FastMCP's existing implementation but with better error handling
                tools_list = await mcp._mcp_list_tools()
                return types.ServerResult(types.ListToolsResult(tools=tools_list))
            except Exception as e:
                # Provide more specific error information
                return types.ServerResult(
                    types.ErrorData(
                        code=0,
                        message=f"Failed to list tools: {str(e)}",
                        data={"error_type": type(e).__name__}
                    )
                )
        
        # Register the enhanced handlers
        # We need to access the underlying MCP server's request handlers directly
        try:
            # Ensure ping handler is properly registered
            mcp._mcp_server.request_handlers[types.PingRequest] = enhanced_ping_handler
            
            # Ensure list_tools handler is properly registered  
            mcp._mcp_server.request_handlers[types.ListToolsRequest] = enhanced_list_tools_handler
            
            # Ensure the default ping handler is also available
            if _ping_handler not in mcp._mcp_server.request_handlers.values():
                mcp._mcp_server.request_handlers[types.PingRequest] = _ping_handler
                
        except Exception:
            # If we can't register the handlers, continue with what we have
            pass
            
    except Exception:
        # If we can't import the required types, continue anyway
        pass

# Call the handler creation function
_create_enhanced_handlers()


# Direct handler registration for built-in MCP methods as fallback
try:
    import mcp.types as types
    from mcp.server.lowlevel.server import _ping_handler
    
    # Ensure the ping handler is properly registered
    # The underlying MCP library already has a working ping handler
    if types.PingRequest not in mcp._mcp_server.request_handlers:
        mcp._mcp_server.request_handlers[types.PingRequest] = _ping_handler
    
    # Create a simple list_tools handler if one doesn't exist
    async def _simple_list_tools_handler(request):
        """Simple list_tools handler that delegates to FastMCP's implementation."""
        try:
            tools = await mcp._mcp_list_tools()
            from mcp.types import ServerResult, ListToolsResult
            return ServerResult(ListToolsResult(tools=tools))
        except Exception as e:
            from mcp.types import ServerResult, ErrorData
            return ServerResult(ErrorData(code=0, message=str(e), data=None))
    
    # Register the list_tools handler
    if types.ListToolsRequest not in mcp._mcp_server.request_handlers:
        mcp._mcp_server.request_handlers[types.ListToolsRequest] = _simple_list_tools_handler
    
except Exception:
    # Silently fail to avoid breaking existing functionality
    pass


def _decode_email_id(email_id: str) -> str:
    """Safely decode email ID from URL encoding.
    
    Microsoft Graph search results return email IDs that are already URL-encoded,
    but they need to be decoded before being used in URL paths.
    """
    return urllib.parse.unquote(email_id)


FOLDERS = {
    k.casefold(): v
    for k, v in {
        "inbox": "inbox",
        "sent": "sentitems",
        "drafts": "drafts",
        "deleted": "deleteditems",
        "junk": "junkemail",
        "archive": "archive",
    }.items()
}


@mcp.tool
def debug_token_info(account_id: str) -> dict[str, Any]:
    """Debug function to inspect token scopes and claims
    
    This helps diagnose authentication and permission issues by showing
    what scopes the current token actually has.
    """
    try:
        # Get current user info to verify token works
        me_info = graph.request("GET", "/me", account_id)
        
        result = {
            "user_info": me_info,
            "token_status": "valid" if me_info else "invalid"
        }
        
        # Try to get more detailed token info
        try:
            token_info = graph.request("GET", "/me/oauth2PermissionGrants", account_id)
            result["permission_grants"] = token_info
        except Exception as e:
            result["permission_grants_error"] = str(e)
            
        return result
        
    except Exception as e:
        return {
            "error": str(e),
            "token_status": "error"
        }


@mcp.tool 
def test_mail_permissions(account_id: str) -> dict[str, Any]:
    """Test specific mail permissions to diagnose 403 errors
    
    This function tests read and write permissions systematically to identify
    exactly where the permission issue occurs.
    """
    results = {}
    
    # Test 1: Basic read access to user info
    try:
        me_info = graph.request("GET", "/me", account_id)
        results["me_endpoint"] = {"status": "success", "data": me_info}
    except Exception as e:
        results["me_endpoint"] = {"status": "error", "error": str(e)}
    
    # Test 2: List mail folders (read permission)
    try:
        folders = graph.request("GET", "/me/mailFolders", account_id)
        results["list_folders"] = {"status": "success", "folder_count": len(folders.get("value", []))}
    except Exception as e:
        results["list_folders"] = {"status": "error", "error": str(e)}
    
    # Test 3: List emails (read permission)  
    try:
        emails = graph.request("GET", "/me/messages?$top=1", account_id)
        results["list_emails"] = {"status": "success", "email_count": len(emails.get("value", []))}
    except Exception as e:
        results["list_emails"] = {"status": "error", "error": str(e)}
    
    # Test 4: Try to get first email for testing write operations
    first_email_id = None
    try:
        emails = graph.request("GET", "/me/messages?$top=1&$select=id", account_id)
        if emails and "value" in emails and emails["value"]:
            first_email_id = emails["value"][0]["id"]
            results["get_test_email"] = {"status": "success", "email_id": first_email_id}
        else:
            results["get_test_email"] = {"status": "no_emails", "message": "No emails found to test with"}
    except Exception as e:
        results["get_test_email"] = {"status": "error", "error": str(e)}
    
    # Test 5: Try write operation (this is where we expect the 403)
    if first_email_id:
        try:
            # Try to read the isRead property first
            current_email = graph.request("GET", f"/me/messages/{first_email_id}?$select=id,isRead", account_id)
            current_is_read = current_email.get("isRead", False) if current_email else False
            
            # Try to update it (this should trigger the 403 with detailed debug info)
            update_result = graph.request("PATCH", f"/me/messages/{first_email_id}", account_id, 
                                        json={"isRead": not current_is_read})
            results["test_update_email"] = {"status": "success", "data": update_result}
        except Exception as e:
            results["test_update_email"] = {"status": "error", "error": str(e)}
    else:
        results["test_update_email"] = {"status": "skipped", "reason": "No email ID available for testing"}
        
    return results


@mcp.tool
def list_accounts() -> list[dict[str, str]]:
    """List all signed-in Microsoft accounts"""
    return [
        {"username": acc.username, "account_id": acc.account_id}
        for acc in auth.list_accounts()
    ]


@mcp.tool
def authenticate_account() -> dict[str, str]:
    """Authenticate a new Microsoft account using device flow authentication

    Returns authentication instructions and device code for the user to complete authentication.
    The user must visit the URL and enter the code to authenticate their Microsoft account.
    """
    app = auth.get_app()
    flow = app.initiate_device_flow(scopes=auth.SCOPES)

    if "user_code" not in flow:
        error_msg = flow.get("error_description", "Unknown error")
        raise Exception(f"Failed to get device code: {error_msg}")

    verification_url = flow.get(
        "verification_uri",
        flow.get("verification_url", "https://microsoft.com/devicelogin"),
    )

    return {
        "status": "authentication_required",
        "instructions": "To authenticate a new Microsoft account:",
        "step1": f"Visit: {verification_url}",
        "step2": f"Enter code: {flow['user_code']}",
        "step3": "Sign in with the Microsoft account you want to add",
        "step4": "After authenticating, use the 'complete_authentication' tool to finish the process",
        "device_code": flow["user_code"],
        "verification_url": verification_url,
        "expires_in": str(flow.get("expires_in", 900)),
        "_flow_cache": str(flow),
    }


@mcp.tool
def complete_authentication(flow_cache: str) -> dict[str, str]:
    """Complete the authentication process after the user has entered the device code

    Args:
        flow_cache: The flow data returned from authenticate_account (the _flow_cache field)

    Returns:
        Account information if authentication was successful
    """
    import ast

    try:
        flow = ast.literal_eval(flow_cache)
    except (ValueError, SyntaxError):
        raise ValueError("Invalid flow cache data")

    app = auth.get_app()
    result = app.acquire_token_by_device_flow(flow)

    if "error" in result:
        error_msg = result.get("error_description", result["error"])
        if "authorization_pending" in error_msg:
            return {
                "status": "pending",
                "message": "Authentication is still pending. The user needs to complete the authentication process.",
                "instructions": "Please ensure you've visited the URL and entered the code, then try again.",
            }
        raise Exception(f"Authentication failed: {error_msg}")

    # Save the token cache
    cache = app.token_cache
    if isinstance(cache, auth.msal.SerializableTokenCache) and cache.has_state_changed:
        auth._write_cache(cache.serialize())

    # Get the newly added account
    accounts = app.get_accounts()
    if accounts:
        # Find the account that matches the token we just got
        for account in accounts:
            if (
                account.get("username", "").lower()
                == result.get("id_token_claims", {})
                .get("preferred_username", "")
                .lower()
            ):
                return {
                    "status": "success",
                    "username": account["username"],
                    "account_id": account["home_account_id"],
                    "message": f"Successfully authenticated {account['username']}",
                }
        # If exact match not found, return the last account
        account = accounts[-1]
        return {
            "status": "success",
            "username": account["username"],
            "account_id": account["home_account_id"],
            "message": f"Successfully authenticated {account['username']}",
        }

    return {
        "status": "error",
        "message": "Authentication succeeded but no account was found",
    }


@mcp.tool
def list_emails(
    account_id: str,
    folder: str = "inbox",
    limit: int = 10,
    include_body: bool = True,
) -> list[dict[str, Any]]:
    """List emails from specified folder"""
    folder_path = FOLDERS.get(folder.casefold(), folder)

    if include_body:
        select_fields = "id,subject,from,toRecipients,ccRecipients,receivedDateTime,hasAttachments,body,conversationId,isRead"
    else:
        select_fields = "id,subject,from,toRecipients,receivedDateTime,hasAttachments,conversationId,isRead"

    params = {
        "$top": min(limit, 100),
        "$select": select_fields,
        "$orderby": "receivedDateTime desc",
    }

    emails = list(
        graph.request_paginated(
            f"/me/mailFolders/{folder_path}/messages",
            account_id,
            params=params,
            limit=limit,
        )
    )

    return emails


@mcp.tool
def get_email(
    email_id: str,
    account_id: str,
    include_body: bool = True,
    body_max_length: int = 50000,
    include_attachments: bool = True,
) -> dict[str, Any]:
    """Get email details with size limits

    Args:
        email_id: The email ID
        account_id: The account ID
        include_body: Whether to include the email body (default: True)
        body_max_length: Maximum characters for body content (default: 50000)
        include_attachments: Whether to include attachment metadata (default: True)
    """
    params = {}
    if include_attachments:
        params["$expand"] = "attachments($select=id,name,size,contentType)"

    result = graph.request("GET", f"/me/messages/{_decode_email_id(email_id)}", account_id, params=params)
    if not result:
        raise ValueError(f"Email with ID {email_id} not found")

    # Truncate body if needed
    if include_body and "body" in result and "content" in result["body"]:
        content = result["body"]["content"]
        if len(content) > body_max_length:
            result["body"]["content"] = (
                content[:body_max_length]
                + f"\n\n[Content truncated - {len(content)} total characters]"
            )
            result["body"]["truncated"] = True
            result["body"]["total_length"] = len(content)
    elif not include_body and "body" in result:
        del result["body"]

    # Remove attachment content bytes to reduce size
    if "attachments" in result and result["attachments"]:
        for attachment in result["attachments"]:
            if "contentBytes" in attachment:
                del attachment["contentBytes"]

    return result


@mcp.tool
def create_email_draft(
    account_id: str,
    to: str | list[str],
    subject: str,
    body: str,
    cc: str | list[str] | None = None,
    attachments: str | list[str] | None = None,
) -> dict[str, Any]:
    """Create an email draft with file path(s) as attachments"""
    to_list = [to] if isinstance(to, str) else to

    message = {
        "subject": subject,
        "body": {"contentType": "Text", "content": body},
        "toRecipients": [{"emailAddress": {"address": addr}} for addr in to_list],
    }

    if cc:
        cc_list = [cc] if isinstance(cc, str) else cc
        message["ccRecipients"] = [
            {"emailAddress": {"address": addr}} for addr in cc_list
        ]

    small_attachments = []
    large_attachments = []

    if attachments:
        # Convert single path to list
        attachment_paths = (
            [attachments] if isinstance(attachments, str) else attachments
        )
        for file_path in attachment_paths:
            path = pl.Path(file_path).expanduser().resolve()
            content_bytes = path.read_bytes()
            att_size = len(content_bytes)
            att_name = path.name

            if att_size < 3 * 1024 * 1024:
                small_attachments.append(
                    {
                        "@odata.type": "#microsoft.graph.fileAttachment",
                        "name": att_name,
                        "contentBytes": base64.b64encode(content_bytes).decode("utf-8"),
                    }
                )
            else:
                large_attachments.append(
                    {
                        "name": att_name,
                        "content_bytes": content_bytes,
                        "content_type": "application/octet-stream",
                    }
                )

    if small_attachments:
        message["attachments"] = small_attachments

    result = graph.request("POST", "/me/messages", account_id, json=message)
    if not result:
        raise ValueError("Failed to create email draft")

    message_id = result["id"]

    for att in large_attachments:
        graph.upload_large_mail_attachment(
            message_id,
            att["name"],
            att["content_bytes"],
            account_id,
            att.get("content_type", "application/octet-stream"),
        )

    return result


@mcp.tool
def send_email(
    account_id: str,
    to: str | list[str],
    subject: str,
    body: str,
    cc: str | list[str] | None = None,
    attachments: str | list[str] | None = None,
) -> dict[str, str]:
    """Send an email immediately with file path(s) as attachments"""
    to_list = [to] if isinstance(to, str) else to

    message = {
        "subject": subject,
        "body": {"contentType": "Text", "content": body},
        "toRecipients": [{"emailAddress": {"address": addr}} for addr in to_list],
    }

    if cc:
        cc_list = [cc] if isinstance(cc, str) else cc
        message["ccRecipients"] = [
            {"emailAddress": {"address": addr}} for addr in cc_list
        ]

    # Check if we have large attachments
    has_large_attachments = False
    processed_attachments = []

    if attachments:
        # Convert single path to list
        attachment_paths = (
            [attachments] if isinstance(attachments, str) else attachments
        )
        for file_path in attachment_paths:
            path = pl.Path(file_path).expanduser().resolve()
            content_bytes = path.read_bytes()
            att_size = len(content_bytes)
            att_name = path.name

            processed_attachments.append(
                {
                    "name": att_name,
                    "content_bytes": content_bytes,
                    "content_type": "application/octet-stream",
                    "size": att_size,
                }
            )

            if att_size >= 3 * 1024 * 1024:
                has_large_attachments = True

    if not has_large_attachments and processed_attachments:
        message["attachments"] = [
            {
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": att["name"],
                "contentBytes": base64.b64encode(att["content_bytes"]).decode("utf-8"),
            }
            for att in processed_attachments
        ]
        graph.request("POST", "/me/sendMail", account_id, json={"message": message})
        return {"status": "sent"}
    elif has_large_attachments:
        # Create draft first, then add large attachments, then send
        # We need to handle large attachments manually here
        to_list = [to] if isinstance(to, str) else to
        message = {
            "subject": subject,
            "body": {"contentType": "Text", "content": body},
            "toRecipients": [{"emailAddress": {"address": addr}} for addr in to_list],
        }
        if cc:
            cc_list = [cc] if isinstance(cc, str) else cc
            message["ccRecipients"] = [
                {"emailAddress": {"address": addr}} for addr in cc_list
            ]

        result = graph.request("POST", "/me/messages", account_id, json=message)
        if not result:
            raise ValueError("Failed to create email draft")

        message_id = result["id"]

        for att in processed_attachments:
            if att["size"] >= 3 * 1024 * 1024:
                graph.upload_large_mail_attachment(
                    message_id,
                    att["name"],
                    att["content_bytes"],
                    account_id,
                    att.get("content_type", "application/octet-stream"),
                )
            else:
                small_att = {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": att["name"],
                    "contentBytes": base64.b64encode(att["content_bytes"]).decode(
                        "utf-8"
                    ),
                }
                graph.request(
                    "POST",
                    f"/me/messages/{message_id}/attachments",
                    account_id,
                    json=small_att,
                )

        graph.request("POST", f"/me/messages/{message_id}/send", account_id)
        return {"status": "sent"}
    else:
        graph.request("POST", "/me/sendMail", account_id, json={"message": message})
        return {"status": "sent"}


@mcp.tool
def update_email(
    email_id: str, updates: dict[str, Any], account_id: str
) -> dict[str, Any]:
    """Update email properties (isRead, categories, flag, etc.)"""
    result = graph.request(
        "PATCH", f"/me/messages/{_decode_email_id(email_id)}", account_id, json=updates
    )
    if not result:
        raise ValueError(f"Failed to update email {email_id} - no response")
    return result


@mcp.tool
def delete_email(email_id: str, account_id: str) -> dict[str, str]:
    """Delete an email"""
    graph.request("DELETE", f"/me/messages/{_decode_email_id(email_id)}", account_id)
    return {"status": "deleted"}


@mcp.tool
def move_email(
    email_id: str, destination_folder: str, account_id: str
) -> dict[str, Any]:
    """Move email to another folder"""
    folder_path = FOLDERS.get(destination_folder.casefold(), destination_folder)

    folders = graph.request("GET", "/me/mailFolders", account_id)
    folder_id = None

    if not folders:
        raise ValueError("Failed to retrieve mail folders")
    if "value" not in folders:
        raise ValueError(f"Unexpected folder response structure: {folders}")

    for folder in folders["value"]:
        if folder["displayName"].lower() == folder_path.lower():
            folder_id = folder["id"]
            break

    if not folder_id:
        raise ValueError(f"Folder '{destination_folder}' not found")

    payload = {"destinationId": folder_id}
    result = graph.request(
        "POST", f"/me/messages/{_decode_email_id(email_id)}/move", account_id, json=payload
    )
    if not result:
        raise ValueError("Failed to move email - no response from server")
    if "id" not in result:
        raise ValueError(f"Failed to move email - unexpected response: {result}")
    return {"status": "moved", "new_id": result["id"]}


@mcp.tool
def reply_to_email(account_id: str, email_id: str, body: str) -> dict[str, str]:
    """Reply to an email (sender only)"""
    endpoint = f"/me/messages/{_decode_email_id(email_id)}/reply"
    payload = {"message": {"body": {"contentType": "Text", "content": body}}}
    graph.request("POST", endpoint, account_id, json=payload)
    return {"status": "sent"}


@mcp.tool
def reply_all_email(account_id: str, email_id: str, body: str) -> dict[str, str]:
    """Reply to all recipients of an email"""
    endpoint = f"/me/messages/{_decode_email_id(email_id)}/replyAll"
    payload = {"message": {"body": {"contentType": "Text", "content": body}}}
    graph.request("POST", endpoint, account_id, json=payload)
    return {"status": "sent"}


@mcp.tool
def list_events(
    account_id: str,
    days_ahead: int = 7,
    days_back: int = 0,
    include_details: bool = True,
) -> list[dict[str, Any]]:
    """List calendar events within specified date range, including recurring event instances.

    Returns event details including categories (array of strings matching user's outlook categories).
    Use list_outlook_categories to see available categories, or create_outlook_category
    to create new ones with colors like 'red', 'blue', 'green', etc."""
    now = dt.datetime.now(dt.timezone.utc)
    start = (now - dt.timedelta(days=days_back)).isoformat()
    end = (now + dt.timedelta(days=days_ahead)).isoformat()

    params = {
        "startDateTime": start,
        "endDateTime": end,
        "$orderby": "start/dateTime",
        "$top": 100,
    }

    if include_details:
        params["$select"] = (
            "id,subject,start,end,location,body,attendees,organizer,isAllDay,recurrence,onlineMeeting,seriesMasterId,categories"
        )
    else:
        params["$select"] = (
            "id,subject,start,end,location,organizer,seriesMasterId,categories"
        )

    # Use calendarView to get recurring event instances
    events = list(
        graph.request_paginated("/me/calendarView", account_id, params=params)
    )

    return events


@mcp.tool
def list_events_by_date_range(
    account_id: str,
    start_date: str,
    end_date: str,
    category_filter: Optional[str] = None,
    exclude_categories: bool = False,
    include_details: bool = True,
) -> list[dict[str, Any]]:
    """List calendar events within a specific date range with optional category filtering.

    Args:
        account_id: The account ID to query
        start_date: Start date in ISO format (YYYY-MM-DD or YYYY-MM-DDTHH:MM:SS)
        end_date: End date in ISO format (YYYY-MM-DD or YYYY-MM-DDTHH:MM:SS)
        category_filter: Only return events with this specific category
        exclude_categories: If True, only return events with no categories
        include_details: Include full event details vs summary only

    Returns event details including categories. For events with specific categories,
    use category_filter. For events without any categories, use exclude_categories=True.
    """

    # Parse and validate dates
    def parse_date_string(date_str: str) -> str:
        """Parse date string and return ISO format with timezone"""
        try:
            # Try parsing as full datetime first
            if "T" in date_str:
                parsed = dt.datetime.fromisoformat(date_str.replace("Z", "+00:00"))
            else:
                # Parse as date only and add time
                parsed = dt.datetime.fromisoformat(date_str + "T00:00:00")

            # Ensure timezone aware
            if parsed.tzinfo is None:
                parsed = parsed.replace(tzinfo=dt.timezone.utc)

            return parsed.isoformat()
        except ValueError as e:
            raise ValueError(
                f"Invalid date format '{date_str}'. Use YYYY-MM-DD or YYYY-MM-DDTHH:MM:SS"
            ) from e

    start_datetime = parse_date_string(start_date)
    end_datetime = parse_date_string(end_date)

    # For date-only inputs, adjust end date to include the full end day
    if "T" not in end_date:
        end_parsed = dt.datetime.fromisoformat(end_datetime.replace("Z", "+00:00"))
        end_parsed = end_parsed.replace(hour=23, minute=59, second=59)
        end_datetime = end_parsed.isoformat()

    # Validate date range
    start_parsed = dt.datetime.fromisoformat(start_datetime.replace("Z", "+00:00"))
    end_parsed = dt.datetime.fromisoformat(end_datetime.replace("Z", "+00:00"))
    if start_parsed >= end_parsed:
        raise ValueError("start_date must be before end_date")

    params = {
        "startDateTime": start_datetime,
        "endDateTime": end_datetime,
        "$orderby": "start/dateTime",
        "$top": 100,
    }

    # Add category filter to OData query if specified
    if category_filter and not exclude_categories:
        params["$filter"] = f"categories/any(x:x eq '{category_filter}')"

    if include_details:
        params["$select"] = (
            "id,subject,start,end,location,body,attendees,organizer,isAllDay,recurrence,onlineMeeting,seriesMasterId,categories"
        )
    else:
        params["$select"] = (
            "id,subject,start,end,location,organizer,seriesMasterId,categories"
        )

    # Use calendarView to get recurring event instances
    events = list(
        graph.request_paginated("/me/calendarView", account_id, params=params)
    )

    # Client-side filtering for events without categories (API limitation)
    if exclude_categories:
        events = [event for event in events if not event.get("categories")]

    return events


@mcp.tool
def get_event(event_id: str, account_id: str) -> dict[str, Any]:
    """Get full event details"""
    result = graph.request("GET", f"/me/events/{event_id}", account_id)
    if not result:
        raise ValueError(f"Event with ID {event_id} not found")
    return result


@mcp.tool
def create_event(
    account_id: str,
    subject: str,
    start: str,
    end: str,
    location: str | None = None,
    body: str | None = None,
    attendees: str | list[str] | None = None,
    timezone: str = "UTC",
    categories: str | list[str] | None = None,
) -> dict[str, Any]:
    """Create a calendar event with optional categories.

    Categories must be strings that match existing outlook categories for the user.
    Use list_outlook_categories to see available categories, or create_outlook_category
    to create new ones with natural language colors like 'red', 'blue', 'green', etc.

    Multiple categories are fully supported using an optimized create-then-update approach
    that ensures reliable event creation. Single categories use direct creation when possible.
    Events will always be created successfully, with detailed status in the response."""

    event = {
        "subject": subject,
        "start": {"dateTime": start, "timeZone": timezone},
        "end": {"dateTime": end, "timeZone": timezone},
    }

    if location:
        event["location"] = {"displayName": location}

    if body:
        event["body"] = {"contentType": "Text", "content": body}

    if attendees:
        attendees_list = [attendees] if isinstance(attendees, str) else attendees
        event["attendees"] = [
            {"emailAddress": {"address": a}, "type": "required"} for a in attendees_list
        ]

    if categories:
        categories_list = [categories] if isinstance(categories, str) else categories

        # Proactive fallback strategy for multiple categories to avoid MCP layer issues
        if len(categories_list) > 1:
            # Use the proven fallback approach for multiple categories
            try:
                # Step 1: Create event with first category only
                event["categories"] = [categories_list[0]]
                result = graph.request("POST", "/me/events", account_id, json=event)

                if result and "id" in result:
                    event_id = result["id"]

                    # Step 2: Update with all categories
                    update_data = {"categories": categories_list}
                    try:
                        graph.request(
                            "PATCH",
                            f"/me/events/{event_id}",
                            account_id,
                            json=update_data,
                        )
                        result["categories"] = categories_list
                        result["_multiple_categories_method"] = "create_then_update"
                        return result
                    except Exception as update_error:
                        # Event created but update failed - still return success with warning
                        result["categories"] = [categories_list[0]]
                        result["_category_warning"] = (
                            f"Event created with first category only. Update failed: {str(update_error)}"
                        )
                        return result
                else:
                    # First step failed, try without categories
                    del event["categories"]
                    result = graph.request("POST", "/me/events", account_id, json=event)
                    if result:
                        result["categories"] = []
                        result["_category_warning"] = (
                            f"Event created without categories. Original categories: {categories_list}"
                        )
                        return result
                    else:
                        raise ValueError(
                            "Failed to create event even without categories"
                        )

            except Exception as e:
                # Comprehensive error handling - catch any exception
                raise ValueError(
                    f"Failed to create event with multiple categories using fallback method. Error: {str(e)}"
                )

        else:
            # Single category - use direct approach with fallback
            event["categories"] = categories_list

            try:
                result = graph.request("POST", "/me/events", account_id, json=event)
                if result:
                    return result
                else:
                    # API returned None - try fallback
                    del event["categories"]
                    result = graph.request("POST", "/me/events", account_id, json=event)
                    if result:
                        # Try to add category via update
                        try:
                            event_id = result["id"]
                            update_data = {"categories": categories_list}
                            graph.request(
                                "PATCH",
                                f"/me/events/{event_id}",
                                account_id,
                                json=update_data,
                            )
                            result["categories"] = categories_list
                        except Exception:
                            result["categories"] = []
                            result["_category_warning"] = (
                                f"Event created but couldn't add category: {categories_list[0]}"
                            )
                        return result
                    else:
                        raise ValueError("Failed to create event")

            except Exception as e:
                # Single category failed - try without categories
                try:
                    del event["categories"]
                    result = graph.request("POST", "/me/events", account_id, json=event)
                    if result:
                        result["categories"] = []
                        result["_category_warning"] = (
                            f"Event created without categories due to error: {str(e)}"
                        )
                        return result
                    else:
                        raise ValueError(
                            f"Failed to create event. Original error: {str(e)}"
                        )
                except Exception as final_error:
                    raise ValueError(
                        f"Failed to create event. Original error: {str(e)}. Final error: {str(final_error)}"
                    )

    # Original path for events without categories
    result = graph.request("POST", "/me/events", account_id, json=event)
    if not result:
        raise ValueError("Failed to create event")
    return result


@mcp.tool
def update_event(
    event_id: str, updates: dict[str, Any], account_id: str
) -> dict[str, Any]:
    """Update event properties including categories.

    Categories should be provided as a list of strings in the updates dict.
    Categories must match existing outlook categories for the user.
    Use list_outlook_categories to see available categories, or create_outlook_category
    to create new ones with colors like 'red', 'blue', 'green', etc."""
    formatted_updates = {}

    if "subject" in updates:
        formatted_updates["subject"] = updates["subject"]
    if "start" in updates:
        formatted_updates["start"] = {
            "dateTime": updates["start"],
            "timeZone": updates.get("timezone", "UTC"),
        }
    if "end" in updates:
        formatted_updates["end"] = {
            "dateTime": updates["end"],
            "timeZone": updates.get("timezone", "UTC"),
        }
    if "location" in updates:
        formatted_updates["location"] = {"displayName": updates["location"]}
    if "body" in updates:
        formatted_updates["body"] = {"contentType": "Text", "content": updates["body"]}
    if "categories" in updates:
        formatted_updates["categories"] = updates["categories"]

    result = graph.request(
        "PATCH", f"/me/events/{event_id}", account_id, json=formatted_updates
    )
    return result or {"status": "updated"}


@mcp.tool
def delete_event(
    account_id: str, event_id: str, send_cancellation: bool = True
) -> dict[str, str]:
    """Delete or cancel a calendar event"""
    if send_cancellation:
        graph.request("POST", f"/me/events/{event_id}/cancel", account_id, json={})
    else:
        graph.request("DELETE", f"/me/events/{event_id}", account_id)
    return {"status": "deleted"}


@mcp.tool
def respond_event(
    account_id: str,
    event_id: str,
    response: str = "accept",
    message: str | None = None,
) -> dict[str, str]:
    """Respond to event invitation (accept, decline, tentativelyAccept)"""
    payload: dict[str, Any] = {"sendResponse": True}
    if message:
        payload["comment"] = message

    graph.request("POST", f"/me/events/{event_id}/{response}", account_id, json=payload)
    return {"status": response}


@mcp.tool
def check_availability(
    account_id: str,
    start: str,
    end: str,
    attendees: str | list[str] | None = None,
) -> dict[str, Any]:
    """Check calendar availability for scheduling"""
    me_info = graph.request("GET", "/me", account_id)
    if not me_info or "mail" not in me_info:
        raise ValueError("Failed to get user email address")
    schedules = [me_info["mail"]]
    if attendees:
        attendees_list = [attendees] if isinstance(attendees, str) else attendees
        schedules.extend(attendees_list)

    payload = {
        "schedules": schedules,
        "startTime": {"dateTime": start, "timeZone": "UTC"},
        "endTime": {"dateTime": end, "timeZone": "UTC"},
        "availabilityViewInterval": 30,
    }

    result = graph.request("POST", "/me/calendar/getSchedule", account_id, json=payload)
    if not result:
        raise ValueError("Failed to check availability")
    return result


@mcp.tool
def list_contacts(account_id: str, limit: int = 50) -> list[dict[str, Any]]:
    """List contacts"""
    params = {"$top": min(limit, 100)}

    contacts = list(
        graph.request_paginated("/me/contacts", account_id, params=params, limit=limit)
    )

    return contacts


@mcp.tool
def get_contact(contact_id: str, account_id: str) -> dict[str, Any]:
    """Get contact details"""
    result = graph.request("GET", f"/me/contacts/{contact_id}", account_id)
    if not result:
        raise ValueError(f"Contact with ID {contact_id} not found")
    return result


@mcp.tool
def create_contact(
    account_id: str,
    given_name: str,
    surname: str | None = None,
    email_addresses: str | list[str] | None = None,
    phone_numbers: dict[str, str] | None = None,
) -> dict[str, Any]:
    """Create a new contact"""
    contact: dict[str, Any] = {"givenName": given_name}

    if surname:
        contact["surname"] = surname

    if email_addresses:
        email_list = (
            [email_addresses] if isinstance(email_addresses, str) else email_addresses
        )
        contact["emailAddresses"] = [
            {"address": email, "name": f"{given_name} {surname or ''}".strip()}
            for email in email_list
        ]

    if phone_numbers:
        if "business" in phone_numbers:
            contact["businessPhones"] = [phone_numbers["business"]]
        if "home" in phone_numbers:
            contact["homePhones"] = [phone_numbers["home"]]
        if "mobile" in phone_numbers:
            contact["mobilePhone"] = phone_numbers["mobile"]

    result = graph.request("POST", "/me/contacts", account_id, json=contact)
    if not result:
        raise ValueError("Failed to create contact")
    return result


@mcp.tool
def update_contact(
    contact_id: str, updates: dict[str, Any], account_id: str
) -> dict[str, Any]:
    """Update contact information"""
    result = graph.request(
        "PATCH", f"/me/contacts/{contact_id}", account_id, json=updates
    )
    return result or {"status": "updated"}


@mcp.tool
def delete_contact(contact_id: str, account_id: str) -> dict[str, str]:
    """Delete a contact"""
    graph.request("DELETE", f"/me/contacts/{contact_id}", account_id)
    return {"status": "deleted"}


@mcp.tool
def list_files(
    account_id: str, path: str = "/", limit: int = 50
) -> list[dict[str, Any]]:
    """List files and folders in OneDrive"""
    endpoint = (
        "/me/drive/root/children"
        if path == "/"
        else f"/me/drive/root:/{path}:/children"
    )
    params = {
        "$top": min(limit, 100),
        "$select": "id,name,size,lastModifiedDateTime,folder,file,@microsoft.graph.downloadUrl",
    }

    items = list(
        graph.request_paginated(endpoint, account_id, params=params, limit=limit)
    )

    return [
        {
            "id": item["id"],
            "name": item["name"],
            "type": "folder" if "folder" in item else "file",
            "size": item.get("size", 0),
            "modified": item.get("lastModifiedDateTime"),
            "download_url": item.get("@microsoft.graph.downloadUrl"),
        }
        for item in items
    ]


@mcp.tool
def get_file(file_id: str, account_id: str, download_path: str) -> dict[str, Any]:
    """Download a file from OneDrive to local path"""
    import subprocess

    metadata = graph.request("GET", f"/me/drive/items/{file_id}", account_id)
    if not metadata:
        raise ValueError(f"File with ID {file_id} not found")

    download_url = metadata.get("@microsoft.graph.downloadUrl")
    if not download_url:
        raise ValueError("No download URL available for this file")

    try:
        subprocess.run(
            ["curl", "-L", "-o", download_path, download_url],
            check=True,
            capture_output=True,
        )

        return {
            "path": download_path,
            "name": metadata.get("name", "unknown"),
            "size_mb": round(metadata.get("size", 0) / (1024 * 1024), 2),
            "mime_type": metadata.get("file", {}).get("mimeType") if metadata else None,
        }
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"Failed to download file: {e.stderr.decode()}")


@mcp.tool
def create_file(
    onedrive_path: str, local_file_path: str, account_id: str
) -> dict[str, Any]:
    """Upload a local file to OneDrive"""
    path = pl.Path(local_file_path).expanduser().resolve()
    data = path.read_bytes()
    result = graph.upload_large_file(
        f"/me/drive/root:/{onedrive_path}:", data, account_id
    )
    if not result:
        raise ValueError(f"Failed to create file at path: {onedrive_path}")
    return result


@mcp.tool
def update_file(file_id: str, local_file_path: str, account_id: str) -> dict[str, Any]:
    """Update OneDrive file content from a local file"""
    path = pl.Path(local_file_path).expanduser().resolve()
    data = path.read_bytes()
    result = graph.upload_large_file(f"/me/drive/items/{file_id}", data, account_id)
    if not result:
        raise ValueError(f"Failed to update file with ID: {file_id}")
    return result


@mcp.tool
def delete_file(file_id: str, account_id: str) -> dict[str, str]:
    """Delete a file or folder"""
    graph.request("DELETE", f"/me/drive/items/{file_id}", account_id)
    return {"status": "deleted"}


@mcp.tool
def get_attachment(
    email_id: str, attachment_id: str, save_path: str, account_id: str
) -> dict[str, Any]:
    """Download email attachment to a specified file path"""
    result = graph.request(
        "GET", f"/me/messages/{_decode_email_id(email_id)}/attachments/{attachment_id}", account_id
    )

    if not result:
        raise ValueError("Attachment not found")

    if "contentBytes" not in result:
        raise ValueError("Attachment content not available")

    # Save attachment to file
    path = pl.Path(save_path).expanduser().resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    content_bytes = base64.b64decode(result["contentBytes"])
    path.write_bytes(content_bytes)

    return {
        "name": result.get("name", "unknown"),
        "content_type": result.get("contentType", "application/octet-stream"),
        "size": result.get("size", 0),
        "saved_to": str(path),
    }


@mcp.tool
def search_files(
    query: str,
    account_id: str,
    limit: int = 50,
) -> list[dict[str, Any]]:
    """Search for files in OneDrive using the modern search API."""
    items = list(graph.search_query(query, ["driveItem"], account_id, limit))

    return [
        {
            "id": item["id"],
            "name": item["name"],
            "type": "folder" if "folder" in item else "file",
            "size": item.get("size", 0),
            "modified": item.get("lastModifiedDateTime"),
            "download_url": item.get("@microsoft.graph.downloadUrl"),
        }
        for item in items
    ]


@mcp.tool
def semantic_search_emails(
    query: str,
    account_id: str,
    limit: int = 50,
    include_content: bool = True,
    relevance_threshold: float = 0.0,
) -> list[dict[str, Any]]:
    """Search emails using semantic understanding and natural language queries.

    PREFERRED SEARCH TOOL: Use this for most email searches, especially:
    - Natural language queries ("find emails about budget meetings from last month")
    - Conceptual searches ("emails discussing project delays")
    - Intent-based searches ("emails asking for help with reports")
    - Relevance-ranked results

    Args:
        query: Natural language or keyword query
        account_id: Microsoft account ID
        limit: Maximum number of results (default: 50)
        include_content: Include email body content for better semantic matching
        relevance_threshold: Minimum relevance score (0.0-1.0) to include results

    Returns ranked results by semantic relevance. Falls back to keyword search if needed.
    For structured filtering (dates, senders), use search_emails_advanced instead.
    """
    # Use enhanced search query with semantic parameters
    results = list(
        graph.search_query(
            query,
            ["message"],
            account_id,
            limit,
            fields=[
                "id",
                "subject",
                "from",
                "toRecipients",
                "receivedDateTime",
                "hasAttachments",
                "body",
                "conversationId",
                "isRead",
            ]
            if include_content
            else None,
            semantic_search=True,
            relevance_threshold=relevance_threshold,
        )
    )

    return results


@mcp.tool
def search_emails(
    query: str,
    account_id: str,
    limit: int = 50,
    folder: str | None = None,
) -> list[dict[str, Any]]:
    """Search emails using basic keyword matching.

    FALLBACK SEARCH TOOL: Use semantic_search_emails for better results.
    Only use this for simple keyword searches in specific folders.

    Search Scope:
    - Default (no folder): Searches across most folders using Microsoft Search API
    - Archive folders: May require explicit specification with folder="archive"
    - If expected emails don't appear, try folder="archive", folder="sent", or folder="inbox"

    For advanced filtering (sender, subject, dates), use search_emails_advanced instead.
    For natural language queries, use semantic_search_emails instead.
    """
    if folder:
        # For folder-specific search, use the traditional endpoint
        folder_path = FOLDERS.get(folder.casefold(), folder)
        endpoint = f"/me/mailFolders/{folder_path}/messages"

        params = {
            "$search": f'"{query}"',
            "$top": min(limit, 100),
            "$select": "id,subject,from,toRecipients,receivedDateTime,hasAttachments,body,conversationId,isRead",
        }

        return list(
            graph.request_paginated(endpoint, account_id, params=params, limit=limit)
        )

    return list(graph.search_query(query, ["message"], account_id, limit))


@mcp.tool
def search_emails_advanced(
    account_id: str,
    query: str = "",
    sender: str = "",
    subject_contains: str = "",
    date_after: str = "",
    date_before: str = "",
    folder: str = "",
    has_attachments: bool = False,
    is_read: bool = True,
    importance: str = "",
    content_level: str = "summary",
    limit: int = 50,
) -> list[dict[str, Any]]:
    """Advanced email search with structured parameters and content optimization

    STRUCTURED SEARCH TOOL: Use this for precise filtering with specific criteria:
    - Date range filtering (date_after, date_before)
    - Sender filtering (sender email/domain)
    - Subject filtering (subject_contains)
    - Folder-specific searches
    - Attachment filtering
    - Read/unread status filtering
    - Importance level filtering

    For natural language queries, use semantic_search_emails instead.
    For simple keyword searches, use search_emails instead.

    Args:
        account_id: Microsoft account ID
        query: Free-text search query
        sender: Email address or domain to search from
        subject_contains: Text that must appear in the subject line
        date_after: ISO date string (e.g., "2025-01-01") for emails received after this date
        date_before: ISO date string (e.g., "2025-12-31") for emails received before this date
        folder: Folder to search in - empty string (default) searches most folders, or specify "inbox", "sent", "drafts", "archive", etc.
        has_attachments: Filter for emails with attachments (True) or without (False)
        is_read: Filter for read emails (True) or unread emails (False)
        importance: Filter by importance level ("high", "normal", "low")
        content_level: Content detail level - "summary", "preview", or "full"
        limit: Maximum number of results to return

    Search Scope:
    - Default (folder=""): Searches across most folders (inbox, sent, drafts, deleted items)
    - Archive folders: May require explicit specification with folder="archive"
    - If expected emails don't appear, try folder="archive", folder="sent", or folder="inbox"

    Specify folder parameter to limit search to specific folder or to access archive.
    """
    # STAGE 1: Build API-level filters for proven working fields
    api_filter_parts = []

    # Use $filter for fields that work reliably
    if not is_read:
        api_filter_parts.append("isRead eq false")

    if date_after:
        iso_date = f"{date_after}T00:00:00Z" if "T" not in date_after else date_after
        api_filter_parts.append(f"receivedDateTime ge {iso_date}")

    if date_before:
        iso_date = f"{date_before}T23:59:59Z" if "T" not in date_before else date_before
        api_filter_parts.append(f"receivedDateTime le {iso_date}")

    if importance:
        api_filter_parts.append(f"importance eq '{importance.lower()}'")

    api_filter_query = " and ".join(api_filter_parts) if api_filter_parts else ""

    # Determine if we need client-side filtering
    needs_client_filtering = bool(
        sender or subject_contains or has_attachments or query
    )

    # Smart limit adjustment: fetch more if we need to filter client-side
    fetch_limit = limit * 3 if needs_client_filtering else limit

    # Define field selection based on content_level
    if content_level == "summary":
        select_fields = "id,subject,from,toRecipients,receivedDateTime,hasAttachments,conversationId,isRead"
    elif content_level == "preview":
        select_fields = "id,subject,from,toRecipients,receivedDateTime,hasAttachments,bodyPreview,conversationId,isRead"
    else:  # full
        select_fields = "id,subject,from,toRecipients,ccRecipients,receivedDateTime,hasAttachments,body,conversationId,isRead"

    # Determine endpoint: cross-folder vs folder-specific search
    if folder:
        # Folder-specific search
        folder_path = FOLDERS.get(folder.casefold(), folder)
        endpoint = f"/me/mailFolders/{folder_path}/messages"
    else:
        # Cross-folder search (default behavior to match search_emails)
        endpoint = "/me/messages"

    params = {
        "$top": min(fetch_limit, 100),
        "$select": select_fields,
        "$orderby": "receivedDateTime desc",
    }

    # Use $filter for proven working API-level filters
    if api_filter_query:
        params["$filter"] = api_filter_query

    # Use $search only for free-text query if no other filters
    elif query and not needs_client_filtering:
        params["$search"] = f'"{query}"'

    # STAGE 1: Get initial results from API
    results = list(
        graph.request_paginated(endpoint, account_id, params=params, limit=fetch_limit)
    )

    # STAGE 2: Apply client-side filtering for unsupported fields
    if needs_client_filtering:
        # Filter by sender (domain or email)
        if sender:
            sender_lower = sender.lower()
            results = [
                r
                for r in results
                if sender_lower
                in r.get("from", {}).get("emailAddress", {}).get("address", "").lower()
            ]

        # Filter by subject contains
        if subject_contains:
            subject_lower = subject_contains.lower()
            results = [
                r for r in results if subject_lower in r.get("subject", "").lower()
            ]

        # Filter by has attachments
        if has_attachments is not None:
            results = [
                r for r in results if r.get("hasAttachments", False) == has_attachments
            ]

        # Filter by free-text query in subject/body
        if query:
            query_lower = query.lower()
            results = [
                r
                for r in results
                if query_lower in r.get("subject", "").lower()
                or query_lower in r.get("body", {}).get("content", "").lower()
                or query_lower in r.get("bodyPreview", "").lower()
            ]

    # Apply final limit after all filtering
    return results[:limit]


@mcp.tool
def semantic_search_calendar(
    query: str,
    account_id: str,
    days_ahead: int = 365,
    days_back: int = 365,
    limit: int = 50,
    include_content: bool = True,
    relevance_threshold: float = 0.0,
) -> list[dict[str, Any]]:
    """Search calendar events using semantic understanding and natural language queries.

    PREFERRED SEARCH TOOL: Use this for most calendar searches, especially:
    - Natural language queries ("find meetings about quarterly planning")
    - Conceptual searches ("events discussing budget reviews")
    - Intent-based searches ("meetings with external clients")
    - Relevance-ranked results

    Args:
        query: Natural language or keyword query
        account_id: Microsoft account ID
        days_ahead: Number of days ahead to search (default: 365)
        days_back: Number of days back to search (default: 365)
        limit: Maximum number of results (default: 50)
        include_content: Include event body content for better semantic matching
        relevance_threshold: Minimum relevance score (0.0-1.0) to include results

    Returns ranked results by semantic relevance. Falls back to keyword search if needed.
    """
    # Use enhanced search query with semantic parameters
    events = list(
        graph.search_query(
            query,
            ["event"],
            account_id,
            limit,
            fields=[
                "id",
                "subject",
                "start",
                "end",
                "location",
                "body",
                "attendees",
                "organizer",
                "categories",
            ]
            if include_content
            else None,
            semantic_search=True,
            relevance_threshold=relevance_threshold,
        )
    )

    # Filter by date range if needed
    if days_ahead != 365 or days_back != 365:
        now = dt.datetime.now(dt.timezone.utc)
        start = now - dt.timedelta(days=days_back)
        end = now + dt.timedelta(days=days_ahead)

        filtered_events = []
        for event in events:
            event_start = dt.datetime.fromisoformat(
                event.get("start", {}).get("dateTime", "").replace("Z", "+00:00")
            )
            event_end = dt.datetime.fromisoformat(
                event.get("end", {}).get("dateTime", "").replace("Z", "+00:00")
            )

            if event_start <= end and event_end >= start:
                filtered_events.append(event)

        return filtered_events

    return events


@mcp.tool
def search_events(
    query: str,
    account_id: str,
    days_ahead: int = 365,
    days_back: int = 365,
    limit: int = 50,
) -> list[dict[str, Any]]:
    """Search calendar events using basic keyword matching.

    FALLBACK SEARCH TOOL: Use semantic_search_calendar for better results.
    Only use this for simple keyword searches or when semantic search fails.

    For natural language queries, use semantic_search_calendar instead.
    """
    events = list(graph.search_query(query, ["event"], account_id, limit))

    # Filter by date range if needed
    if days_ahead != 365 or days_back != 365:
        now = dt.datetime.now(dt.timezone.utc)
        start = now - dt.timedelta(days=days_back)
        end = now + dt.timedelta(days=days_ahead)

        filtered_events = []
        for event in events:
            event_start = dt.datetime.fromisoformat(
                event.get("start", {}).get("dateTime", "").replace("Z", "+00:00")
            )
            event_end = dt.datetime.fromisoformat(
                event.get("end", {}).get("dateTime", "").replace("Z", "+00:00")
            )

            if event_start <= end and event_end >= start:
                filtered_events.append(event)

        return filtered_events

    return events


@mcp.tool
def search_contacts(
    query: str,
    account_id: str,
    limit: int = 50,
) -> list[dict[str, Any]]:
    """Search contacts. Uses traditional search since unified_search doesn't support contacts."""
    params = {
        "$search": f'"{query}"',
        "$top": min(limit, 100),
    }

    contacts = list(
        graph.request_paginated("/me/contacts", account_id, params=params, limit=limit)
    )

    return contacts


@mcp.tool
def semantic_unified_search(
    query: str,
    account_id: str,
    limit: int = 50,
    include_content: bool = True,
    relevance_threshold: float = 0.0,
) -> dict[str, list[dict[str, Any]]]:
    """Search across multiple Microsoft 365 resources using semantic understanding.

    PREFERRED UNIFIED SEARCH TOOL: Use this for cross-resource searches, especially:
    - Natural language queries ("find everything about project alpha")
    - Conceptual searches ("documents and emails about budget planning")
    - Intent-based searches ("meetings and files related to client feedback")
    - Relevance-ranked results across all resources

    Args:
        query: Natural language or keyword query
        account_id: Microsoft account ID
        limit: Maximum number of results per type (default: 50)
        include_content: Include content for better semantic matching
        relevance_threshold: Minimum relevance score (0.0-1.0) to include results

    Returns ranked results by semantic relevance across emails, events, and files.
    For targeted searches, use semantic_search_emails or semantic_search_calendar instead.
    """
    # Search across all supported entity types
    final_entity_types = ["message", "event", "driveItem"]

    results = {}  # Start with empty dict, only add entity types that have results

    # Microsoft Graph API doesn't support all entity type combinations
    # Search each entity type separately and combine results
    for entity_type in final_entity_types:
        try:
            # Use appropriate fields for each entity type
            if entity_type == "message":
                entity_fields = (
                    [
                        "id",
                        "subject",
                        "from",
                        "toRecipients",
                        "receivedDateTime",
                        "hasAttachments",
                        "body",
                        "conversationId",
                        "isRead",
                    ]
                    if include_content
                    else None
                )
            elif entity_type == "event":
                entity_fields = (
                    [
                        "id",
                        "subject",
                        "start",
                        "end",
                        "location",
                        "body",
                        "attendees",
                        "organizer",
                        "categories",
                    ]
                    if include_content
                    else None
                )
            elif entity_type == "driveItem":
                entity_fields = (
                    ["id", "name", "size", "lastModifiedDateTime", "webUrl"]
                    if include_content
                    else None
                )
            else:
                entity_fields = None

            # Search this entity type with semantic parameters
            items = list(
                graph.search_query(
                    query,
                    [entity_type],  # Single entity type per request
                    account_id,
                    limit,
                    fields=entity_fields,
                    semantic_search=True,
                    relevance_threshold=relevance_threshold,
                )
            )

            # Only add to results if there are actual items
            if items:
                results[entity_type] = items

        except Exception as e:
            # If individual entity type fails, continue with others
            print(f"Warning: Search failed for entity type {entity_type}: {e}")
            continue

    return results


@mcp.tool
def unified_search(
    query: str,
    account_id: str,
    limit: int = 50,
) -> dict[str, list[dict[str, Any]]]:
    """Search across multiple Microsoft 365 resources using basic keyword matching.

    FALLBACK UNIFIED SEARCH TOOL: Use semantic_unified_search for better results.
    Only use this for simple keyword searches or when semantic search fails.

    Searches across emails, calendar events, and files using basic keyword matching.
    For natural language queries, use semantic_unified_search instead.
    For targeted searches, use semantic_search_emails or semantic_search_calendar.
    """
    # Search across all supported entity types
    final_entity_types = ["message", "event", "driveItem"]

    results = {}  # Start with empty dict, only add entity types that have results

    # Microsoft Graph API doesn't support all entity type combinations
    # Search each entity type separately and combine results
    for entity_type in final_entity_types:
        try:
            items = list(graph.search_query(query, [entity_type], account_id, limit))
            # Only add to results if there are actual items
            if items:
                results[entity_type] = items
        except Exception as e:
            # If individual entity type fails, continue with others
            print(f"Warning: Search failed for entity type {entity_type}: {e}")
            continue

    return results


@mcp.tool
def list_outlook_categories(account_id: str) -> list[dict[str, Any]]:
    """List all outlook categories defined for the user.

    Returns categories with their displayName (used for event categories), color properties, and id.
    These are the categories that can be assigned to events, emails, and other items.

    The returned id field is needed for update_outlook_category_color() and delete_outlook_category() operations.
    Use get_outlook_category_by_name() to find a category by display name when you only know the name."""
    categories = list(
        graph.request_paginated("/me/outlook/masterCategories", account_id)
    )
    return categories


@mcp.tool
def create_outlook_category(
    account_id: str, display_name: str, color: str = "blue"
) -> dict[str, Any]:
    """Create a new outlook category that can be used for events, emails, and other items.

    IMPORTANT: Once created, category names cannot be changed. Only the color can be updated.
    To change a category name, you must delete the old category and create a new one.

    Args:
        display_name: The name of the category (must be unique for the user, cannot be changed later)
        color: Color name like 'red', 'blue', 'green', 'yellow', 'orange', 'purple',
               'cranberry', 'teal', 'olive', 'brown', 'steel', 'gray', 'black'.
               Add 'dark' prefix for darker variants (e.g. 'dark blue', 'dark green').

    Use this before assigning categories to events if the category doesn't exist yet.
    Use update_outlook_category_color() to change the color of existing categories."""

    # Convert natural language color to preset code
    color_lower = color.lower().strip()

    # Debug: Check if COLOR_MAPPING is available
    if not COLOR_MAPPING:
        raise ValueError("Color mapping not initialized. Please report this bug.")

    if color_lower in COLOR_MAPPING:
        preset_color = COLOR_MAPPING[color_lower]
    else:
        # If color not found, suggest available colors
        available_colors = sorted(set(COLOR_MAPPING.keys()))
        close_matches = [
            c for c in available_colors if color_lower in c or c in color_lower
        ]
        if close_matches:
            raise ValueError(
                f"Unknown color '{color}'. Did you mean: {', '.join(close_matches)}? All available colors: {', '.join(available_colors)}"
            )
        else:
            raise ValueError(
                f"Unknown color '{color}'. Available colors: {', '.join(available_colors)}"
            )

    category_data = {"displayName": display_name, "color": preset_color}

    try:
        result = graph.request(
            "POST", "/me/outlook/masterCategories", account_id, json=category_data
        )
        if not result:
            raise ValueError(
                f"Failed to create category '{display_name}' - no response from server"
            )
        return result
    except Exception as e:
        error_msg = str(e)
        # Enhanced error handling for API issues
        if "400" in error_msg:
            # Check if it's a duplicate category error
            if (
                "already exists" in error_msg.lower()
                or "duplicate" in error_msg.lower()
            ):
                raise ValueError(
                    f"Category '{display_name}' already exists. Please choose a different name or use the existing category."
                )
            # Check if it's a color-related error
            elif "color" in error_msg.lower() or "preset" in error_msg.lower():
                available_colors = sorted(set(COLOR_MAPPING.keys()))
                raise ValueError(
                    f"Invalid color preset '{preset_color}' for color '{color}'. Available colors: {', '.join(available_colors)}. Error: {error_msg}"
                )
            # General 400 error with detailed context
            else:
                available_colors = sorted(set(COLOR_MAPPING.keys()))
                raise ValueError(
                    f"Failed to create category '{display_name}' with color '{color}' (preset: {preset_color}). Available colors: {', '.join(available_colors)}. Server error: {error_msg}"
                )
        elif "401" in error_msg or "403" in error_msg:
            raise ValueError(
                f"Authentication error: Please ensure you're properly authenticated. Error: {error_msg}"
            )
        else:
            raise ValueError(f"Failed to create category '{display_name}': {error_msg}")


@mcp.tool
def list_available_colors() -> dict[str, str]:
    """List all available colors for outlook categories.

    Returns a mapping of color names to their descriptions, useful for creating
    categories with natural language color names.

    [VERSION: 2025-01-09-FIXED] - Includes parameter validation fixes"""

    # Create a clean mapping with descriptions
    color_descriptions = {
        "red": "Bright red",
        "orange": "Bright orange",
        "brown": "Brown",
        "yellow": "Bright yellow",
        "green": "Bright green",
        "teal": "Teal/cyan",
        "olive": "Olive green",
        "blue": "Bright blue",
        "purple": "Bright purple",
        "cranberry": "Cranberry red",
        "steel": "Steel blue",
        "gray": "Medium gray",
        "black": "Black",
        "dark red": "Dark red",
        "dark orange": "Dark orange",
        "dark brown": "Dark brown",
        "dark yellow": "Dark yellow",
        "dark green": "Dark green",
        "dark teal": "Dark teal",
        "dark olive": "Dark olive",
        "dark blue": "Dark blue",
        "dark purple": "Dark purple",
        "dark cranberry": "Dark cranberry",
    }

    return color_descriptions


@mcp.tool
def update_outlook_category_color(
    account_id: str, category_id: str, color: str
) -> dict[str, Any]:
    """Update the color of an existing outlook category.

    IMPORTANT: Only the color property can be updated for existing categories.
    To change a category name, you must delete the old category and create a new one.

    Args:
        account_id: The account ID to use for authentication
        category_id: The ID of the category to update (get from list_outlook_categories)
        color: Color name like 'red', 'blue', 'green', 'yellow', 'orange', 'purple',
               'teal', 'olive', 'brown', 'cranberry', 'steel', 'gray', 'black',
               or dark variants like 'dark red', 'dark blue', etc.

    Returns:
        Updated category object with id, displayName, and color

    Use this when you need to change the color of an existing category.
    Use get_outlook_category_by_name() to find category ID when you only know the display name."""

    # Convert natural language color to preset code
    color_lower = color.lower().strip()

    # Validate color exists in our mapping
    if color_lower not in COLOR_MAPPING:
        available_colors = sorted(set(COLOR_MAPPING.keys()))
        close_matches = [
            c for c in available_colors if color_lower in c or c in color_lower
        ]
        if close_matches:
            raise ValueError(
                f"Unknown color '{color}'. Did you mean: {', '.join(close_matches)}? All available colors: {', '.join(available_colors)}"
            )
        else:
            raise ValueError(
                f"Unknown color '{color}'. Available colors: {', '.join(available_colors)}"
            )

    preset_color = COLOR_MAPPING[color_lower]

    # Prepare update data
    update_data = {"color": preset_color}

    try:
        result = graph.request(
            "PATCH",
            f"/me/outlook/masterCategories/{category_id}",
            account_id,
            json=update_data,
        )
        if not result:
            raise ValueError(
                "Failed to update category color - no response from server"
            )

        return result

    except Exception as e:
        error_msg = str(e).lower()

        # Check for specific error types and provide helpful messages
        if "404" in error_msg or "not found" in error_msg:
            raise ValueError(
                f"Category with ID '{category_id}' not found. Use list_outlook_categories() to see available categories."
            )
        elif "403" in error_msg or "forbidden" in error_msg:
            raise ValueError(
                "Permission denied. Make sure you have MailboxSettings.ReadWrite permission to update categories."
            )
        elif "400" in error_msg or "bad request" in error_msg:
            available_colors = sorted(set(COLOR_MAPPING.keys()))
            raise ValueError(
                f"Invalid color preset '{preset_color}' for color '{color}'. Available colors: {', '.join(available_colors)}. Error: {error_msg}"
            )
        else:
            raise ValueError(f"Failed to update category color: {str(e)}")


@mcp.tool
def delete_outlook_category(account_id: str, category_id: str) -> dict[str, str]:
    """Delete an outlook category.

    IMPORTANT: Deleting a category will not remove it from existing messages -
    it will appear grayed out in those messages. You'll need to manually remove
    categories from individual messages if desired.

    Args:
        account_id: The account ID to use for authentication
        category_id: The ID of the category to delete (get from list_outlook_categories)

    Returns:
        Success confirmation message

    Use get_outlook_category_by_name() to find category ID when you only know the display name."""

    try:
        # DELETE request returns 204 No Content on success (no response body)
        graph.request(
            "DELETE", f"/me/outlook/masterCategories/{category_id}", account_id
        )

        return {
            "message": f"Category '{category_id}' deleted successfully",
            "warning": "This category will still appear grayed out in existing messages. Remove manually from individual messages if needed.",
        }

    except Exception as e:
        error_msg = str(e).lower()

        # Check for specific error types and provide helpful messages
        if "404" in error_msg or "not found" in error_msg:
            raise ValueError(
                f"Category with ID '{category_id}' not found. Use list_outlook_categories() to see available categories."
            )
        elif "403" in error_msg or "forbidden" in error_msg:
            raise ValueError(
                "Permission denied. Make sure you have MailboxSettings.ReadWrite permission to delete categories."
            )
        else:
            raise ValueError(f"Failed to delete category: {str(e)}")


@mcp.tool
def get_outlook_category_by_name(
    account_id: str, display_name: str
) -> dict[str, Any] | None:
    """Find an outlook category by its display name.

    This is a helper function to get the category ID when you only know the display name.
    Use the returned ID with update_outlook_category_color() or delete_outlook_category().

    Args:
        account_id: The account ID to use for authentication
        display_name: The display name of the category to find

    Returns:
        Category object with id, displayName, and color if found, None otherwise

    Example:
        category = get_outlook_category_by_name(account_id, "Personal")
        if category:
            update_outlook_category_color(account_id, category["id"], "red")
    """

    try:
        categories = list_outlook_categories(account_id)

        # Search for category by display name (case-insensitive)
        display_name_lower = display_name.lower().strip()
        for category in categories:
            if category["displayName"].lower() == display_name_lower:
                return category

        return None

    except Exception as e:
        raise ValueError(f"Failed to search for category '{display_name}': {str(e)}")
