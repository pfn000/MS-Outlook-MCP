"""
Microsoft MCP Server - Integration with Outlook, Calendar, OneDrive, SharePoint, and Teams
Provides comprehensive tools for email management, calendar operations, and cloud file access.
"""

import json
import httpx
import os
from datetime import datetime, timedelta
from typing import Optional, List, Dict, Any
from enum import Enum
from contextlib import asynccontextmanager
from urllib.parse import urlencode

from mcp.server.fastmcp import FastMCP
from pydantic import BaseModel, Field, field_validator, ConfigDict


# ============================================================================
# Configuration & Constants
# ============================================================================

GRAPH_API_ENDPOINT = "https://graph.microsoft.com/v1.0"
AZURE_SCOPE = [
    "Calendars.Read",
    "Calendars.ReadWrite",
    "Mail.Read",
    "Mail.ReadWrite",
    "Mail.Send",
    "Files.Read.All",
    "Files.ReadWrite.All",
    "Sites.Read.All",
    "Team.ReadBasic.All",
]

# Store for tokens (in production, use secure storage)
TOKEN_CACHE = {}


# ============================================================================
# Authentication Models & Helpers
# ============================================================================

class MicrosoftAuthConfig(BaseModel):
    """Configuration for Microsoft OAuth 2.0 authentication."""
    model_config = ConfigDict(
        str_strip_whitespace=True,
        validate_assignment=True,
        extra='forbid'
    )
    
    client_id: str = Field(..., description="Azure Application Client ID", min_length=1)
    client_secret: str = Field(..., description="Azure Application Client Secret", min_length=1)
    tenant_id: str = Field(..., description="Azure Tenant ID", min_length=1)
    redirect_uri: str = Field(
        default="http://localhost:8000/callback",
        description="OAuth redirect URI"
    )


class GraphAPIClient:
    """Client for Microsoft Graph API with OAuth token management."""
    
    def __init__(self, config: MicrosoftAuthConfig):
        self.config = config
        self.token = None
        self.token_expiry = None
    
    async def get_access_token(self) -> str:
        """Get a valid access token, refreshing if necessary."""
        if self.token and self.token_expiry and datetime.now() < self.token_expiry:
            return self.token
        
        async with httpx.AsyncClient() as client:
            token_url = f"https://login.microsoftonline.com/{self.config.tenant_id}/oauth2/v2.0/token"
            
            data = {
                "client_id": self.config.client_id,
                "client_secret": self.config.client_secret,
                "scope": " ".join(AZURE_SCOPE),
                "grant_type": "client_credentials",
            }
            
            response = await client.post(token_url, data=data)
            response.raise_for_status()
            
            token_data = response.json()
            self.token = token_data["access_token"]
            expires_in = token_data.get("expires_in", 3600)
            self.token_expiry = datetime.now() + timedelta(seconds=expires_in - 60)
            
            return self.token
    
    async def _make_request(
        self,
        method: str,
        endpoint: str,
        **kwargs
    ) -> Dict[str, Any]:
        """Make an authenticated request to Microsoft Graph API."""
        token = await self.get_access_token()
        
        headers = kwargs.pop("headers", {})
        headers["Authorization"] = f"Bearer {token}"
        headers["Content-Type"] = "application/json"
        
        url = f"{GRAPH_API_ENDPOINT}{endpoint}"
        
        async with httpx.AsyncClient() as client:
            response = await client.request(method, url, headers=headers, **kwargs)
            
            if response.status_code == 404:
                raise ValueError("Resource not found")
            elif response.status_code == 403:
                raise PermissionError("Access denied to this resource")
            elif response.status_code == 429:
                raise RuntimeError("Rate limit exceeded, please try again later")
            
            response.raise_for_status()
            return response.json()
    
    async def get(self, endpoint: str, **kwargs) -> Dict[str, Any]:
        """GET request to Graph API."""
        return await self._make_request("GET", endpoint, **kwargs)
    
    async def post(self, endpoint: str, json_data: Dict[str, Any], **kwargs) -> Dict[str, Any]:
        """POST request to Graph API."""
        return await self._make_request("POST", endpoint, json=json_data, **kwargs)
    
    async def patch(self, endpoint: str, json_data: Dict[str, Any], **kwargs) -> Dict[str, Any]:
        """PATCH request to Graph API."""
        return await self._make_request("PATCH", endpoint, json=json_data, **kwargs)
    
    async def delete(self, endpoint: str, **kwargs) -> None:
        """DELETE request to Graph API."""
        await self._make_request("DELETE", endpoint, **kwargs)


# ============================================================================
# Email Models & Validators
# ============================================================================

class EmailPriority(str, Enum):
    """Email priority levels."""
    LOW = "low"
    NORMAL = "normal"
    HIGH = "high"


class EmailImportance(str, Enum):
    """Email importance levels."""
    LOW = "low"
    NORMAL = "normal"
    HIGH = "high"


class SearchEmailsInput(BaseModel):
    """Input for searching emails."""
    model_config = ConfigDict(
        str_strip_whitespace=True,
        validate_assignment=True,
        extra='forbid'
    )
    
    query: str = Field(..., description="Search query (sender, subject, body content)", min_length=1, max_length=500)
    folder: str = Field(default="inbox", description="Email folder to search (inbox, sent, drafts, archive)")
    limit: int = Field(default=20, description="Maximum emails to return", ge=1, le=100)
    unread_only: bool = Field(default=False, description="Return only unread emails")


class SendEmailInput(BaseModel):
    """Input for sending an email."""
    model_config = ConfigDict(
        str_strip_whitespace=True,
        validate_assignment=True,
        extra='forbid'
    )
    
    to: List[str] = Field(..., description="Recipient email addresses", min_items=1)
    cc: Optional[List[str]] = Field(default=None, description="CC recipients")
    bcc: Optional[List[str]] = Field(default=None, description="BCC recipients")
    subject: str = Field(..., description="Email subject line", min_length=1, max_length=500)
    body: str = Field(..., description="Email body content (HTML or plain text)", min_length=1)
    importance: EmailImportance = Field(default=EmailImportance.NORMAL, description="Email importance")
    is_html: bool = Field(default=True, description="Whether body is HTML content")


class CreateDraftInput(BaseModel):
    """Input for creating an email draft."""
    model_config = ConfigDict(
        str_strip_whitespace=True,
        validate_assignment=True,
        extra='forbid'
    )
    
    to: Optional[List[str]] = Field(default=None, description="Recipient email addresses")
    cc: Optional[List[str]] = Field(default=None, description="CC recipients")
    subject: Optional[str] = Field(default="", description="Email subject line")
    body: str = Field(..., description="Email body content", min_length=1)
    is_html: bool = Field(default=True, description="Whether body is HTML content")


class GetEmailInput(BaseModel):
    """Input for retrieving a specific email."""
    model_config = ConfigDict(
        str_strip_whitespace=True,
        validate_assignment=True,
        extra='forbid'
    )
    
    message_id: str = Field(..., description="Email message ID", min_length=1)
    include_attachments: bool = Field(default=True, description="Include attachment metadata")


class ManageFolderInput(BaseModel):
    """Input for managing email folders."""
    model_config = ConfigDict(
        str_strip_whitespace=True,
        validate_assignment=True,
        extra='forbid'
    )
    
    action: str = Field(..., description="Action: list, create, delete, rename", min_length=1)
    folder_name: Optional[str] = Field(default=None, description="Folder name for create/rename")
    new_name: Optional[str] = Field(default=None, description="New folder name for rename action")


# ============================================================================
# Calendar Models & Validators
# ============================================================================

class CreateEventInput(BaseModel):
    """Input for creating a calendar event."""
    model_config = ConfigDict(
        str_strip_whitespace=True,
        validate_assignment=True,
        extra='forbid'
    )
    
    subject: str = Field(..., description="Event title", min_length=1, max_length=500)
    start_time: str = Field(..., description="Start time (ISO 8601 format: 2024-01-15T10:00:00Z)")
    end_time: str = Field(..., description="End time (ISO 8601 format: 2024-01-15T11:00:00Z)")
    attendees: Optional[List[str]] = Field(default=None, description="Attendee email addresses")
    description: Optional[str] = Field(default=None, description="Event description")
    location: Optional[str] = Field(default=None, description="Event location")
    is_reminder: bool = Field(default=True, description="Send reminder notification")


class CheckAvailabilityInput(BaseModel):
    """Input for checking availability."""
    model_config = ConfigDict(
        str_strip_whitespace=True,
        validate_assignment=True,
        extra='forbid'
    )
    
    start_time: str = Field(..., description="Start time (ISO 8601 format)")
    end_time: str = Field(..., description="End time (ISO 8601 format)")
    attendees: Optional[List[str]] = Field(default=None, description="Attendee email addresses to check")


class GetEventsInput(BaseModel):
    """Input for retrieving calendar events."""
    model_config = ConfigDict(
        str_strip_whitespace=True,
        validate_assignment=True,
        extra='forbid'
    )
    
    start_date: Optional[str] = Field(default=None, description="Start date (YYYY-MM-DD)")
    end_date: Optional[str] = Field(default=None, description="End date (YYYY-MM-DD)")
    limit: int = Field(default=20, description="Maximum events to return", ge=1, le=100)
    include_details: bool = Field(default=True, description="Include full event details")


# ============================================================================
# Cloud Storage Models & Validators
# ============================================================================

class ListFilesInput(BaseModel):
    """Input for listing files."""
    model_config = ConfigDict(
        str_strip_whitespace=True,
        validate_assignment=True,
        extra='forbid'
    )
    
    location: str = Field(default="onedrive", description="onedrive, sharepoint, or teams")
    folder_path: str = Field(default="/", description="Folder path to list")
    limit: int = Field(default=20, description="Maximum items to return", ge=1, le=100)
    recursive: bool = Field(default=False, description="Include subfolders recursively")


class UploadFileInput(BaseModel):
    """Input for uploading a file."""
    model_config = ConfigDict(
        str_strip_whitespace=True,
        validate_assignment=True,
        extra='forbid'
    )
    
    file_path: str = Field(..., description="Local file path to upload", min_length=1)
    destination: str = Field(..., description="Destination path in cloud", min_length=1)
    location: str = Field(default="onedrive", description="onedrive, sharepoint, or teams")


class CreateFolderInput(BaseModel):
    """Input for creating a folder."""
    model_config = ConfigDict(
        str_strip_whitespace=True,
        validate_assignment=True,
        extra='forbid'
    )
    
    folder_name: str = Field(..., description="Name of the folder to create", min_length=1)
    parent_path: str = Field(default="/", description="Parent folder path")
    location: str = Field(default="onedrive", description="onedrive, sharepoint, or teams")


# ============================================================================
# FastMCP Server Setup
# ============================================================================

@asynccontextmanager
async def app_lifespan():
    """Initialize and manage MCP server lifecycle."""
    # Load configuration from environment variables
    config = MicrosoftAuthConfig(
        client_id=os.getenv("MICROSOFT_CLIENT_ID", ""),
        client_secret=os.getenv("MICROSOFT_CLIENT_SECRET", ""),
        tenant_id=os.getenv("MICROSOFT_TENANT_ID", ""),
    )
    
    # Initialize Graph API client
    client = GraphAPIClient(config)
    
    yield {
        "config": config,
        "client": client,
    }


mcp = FastMCP("microsoft_mcp", lifespan=app_lifespan)


# ============================================================================
# Email Tools
# ============================================================================

@mcp.tool(
    name="microsoft_search_emails",
    annotations={
        "title": "Search Emails",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False
    }
)
async def search_emails(params: SearchEmailsInput, ctx) -> str:
    """Search emails with advanced filtering and full-text search capabilities.
    
    Args:
        params (SearchEmailsInput): Search criteria including:
            - query: Search terms (sender, subject, body)
            - folder: Email folder (inbox, sent, drafts, archive)
            - limit: Maximum results
            - unread_only: Filter unread emails
    
    Returns:
        str: JSON formatted search results with email metadata
    """
    client = ctx.request_context.lifespan_state["client"]
    
    try:
        # Build filter query
        filters = [f'search("{params.query}")']
        if params.unread_only:
            filters.append("isRead eq false")
        
        filter_query = " and ".join(filters)
        
        # Map folder names to Graph API folder names
        folder_map = {
            "inbox": "mailFolders/inbox",
            "sent": "mailFolders/sentItems",
            "drafts": "mailFolders/drafts",
            "archive": "mailFolders/archive"
        }
        
        folder_path = folder_map.get(params.folder.lower(), "mailFolders/inbox")
        
        # Search emails
        endpoint = f"/me/{folder_path}/messages"
        response = await client.get(
            endpoint,
            params={
                "$search": f'"{params.query}"',
                "$filter": filter_query if params.unread_only else None,
                "$top": params.limit,
                "$select": "id,subject,from,receivedDateTime,isRead,importance,preview"
            }
        )
        
        emails = response.get("value", [])
        
        result = {
            "search_query": params.query,
            "folder": params.folder,
            "results_count": len(emails),
            "emails": [
                {
                    "id": email.get("id"),
                    "subject": email.get("subject"),
                    "from": email.get("from", {}).get("emailAddress", {}).get("address"),
                    "received_date": email.get("receivedDateTime"),
                    "unread": not email.get("isRead"),
                    "importance": email.get("importance"),
                    "preview": email.get("preview", "")[:200]
                }
                for email in emails
            ]
        }
        
        return json.dumps(result, indent=2)
    
    except Exception as e:
        return json.dumps({"error": str(e)})


@mcp.tool(
    name="microsoft_send_email",
    annotations={
        "title": "Send Email",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True
    }
)
async def send_email(params: SendEmailInput, ctx) -> str:
    """Send an email message with optional CC/BCC and HTML formatting.
    
    Args:
        params (SendEmailInput): Email details including:
            - to: Recipient email addresses
            - cc: Optional CC recipients
            - bcc: Optional BCC recipients
            - subject: Email subject
            - body: Email body content
            - importance: Email priority
            - is_html: Whether body is HTML
    
    Returns:
        str: Confirmation with message ID or error
    """
    client = ctx.request_context.lifespan_state["client"]
    
    try:
        message = {
            "subject": params.subject,
            "body": {
                "contentType": "HTML" if params.is_html else "text",
                "content": params.body
            },
            "toRecipients": [{"emailAddress": {"address": addr}} for addr in params.to],
            "importance": params.importance.value.capitalize()
        }
        
        if params.cc:
            message["ccRecipients"] = [{"emailAddress": {"address": addr}} for addr in params.cc]
        
        if params.bcc:
            message["bccRecipients"] = [{"emailAddress": {"address": addr}} for addr in params.bcc]
        
        # Send email
        response = await client.post(
            "/me/sendMail",
            {"message": message, "saveToSentItems": True}
        )
        
        return json.dumps({
            "status": "success",
            "message": "Email sent successfully",
            "recipients": params.to,
            "subject": params.subject
        }, indent=2)
    
    except Exception as e:
        return json.dumps({"status": "error", "message": str(e)})


@mcp.tool(
    name="microsoft_create_email_draft",
    annotations={
        "title": "Create Email Draft",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": False
    }
)
async def create_email_draft(params: CreateDraftInput, ctx) -> str:
    """Create an email draft for later editing and sending.
    
    Args:
        params (CreateDraftInput): Draft content including:
            - to: Optional recipient addresses
            - cc: Optional CC recipients
            - subject: Email subject
            - body: Email body content
            - is_html: Whether body is HTML
    
    Returns:
        str: Draft creation confirmation with draft ID
    """
    client = ctx.request_context.lifespan_state["client"]
    
    try:
        message = {
            "subject": params.subject,
            "body": {
                "contentType": "HTML" if params.is_html else "text",
                "content": params.body
            }
        }
        
        if params.to:
            message["toRecipients"] = [{"emailAddress": {"address": addr}} for addr in params.to]
        
        if params.cc:
            message["ccRecipients"] = [{"emailAddress": {"address": addr}} for addr in params.cc]
        
        response = await client.post(
            "/me/mailFolders/drafts/messages",
            message
        )
        
        return json.dumps({
            "status": "success",
            "message": "Draft created successfully",
            "draft_id": response.get("id"),
            "subject": params.subject,
            "edit_link": f"Draft ID: {response.get('id')}"
        }, indent=2)
    
    except Exception as e:
        return json.dumps({"status": "error", "message": str(e)})


@mcp.tool(
    name="microsoft_get_email_details",
    annotations={
        "title": "Get Email Details",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False
    }
)
async def get_email_details(params: GetEmailInput, ctx) -> str:
    """Retrieve full details of a specific email including body and attachments.
    
    Args:
        params (GetEmailInput): Email selection with:
            - message_id: ID of the email
            - include_attachments: Whether to include attachment metadata
    
    Returns:
        str: Full email content and metadata
    """
    client = ctx.request_context.lifespan_state["client"]
    
    try:
        select_fields = "id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,importance,bodyPreview,body,categories"
        if params.include_attachments:
            select_fields += ",attachments"
        
        response = await client.get(
            f"/me/messages/{params.message_id}",
            params={"$select": select_fields}
        )
        
        email_detail = {
            "id": response.get("id"),
            "subject": response.get("subject"),
            "from": response.get("from", {}).get("emailAddress", {}).get("address"),
            "to": [addr.get("emailAddress", {}).get("address") for addr in response.get("toRecipients", [])],
            "cc": [addr.get("emailAddress", {}).get("address") for addr in response.get("ccRecipients", [])],
            "received_date": response.get("receivedDateTime"),
            "sent_date": response.get("sentDateTime"),
            "unread": not response.get("isRead"),
            "importance": response.get("importance"),
            "body": response.get("body", {}).get("content", ""),
            "categories": response.get("categories", []),
            "attachments": [
                {
                    "id": att.get("id"),
                    "name": att.get("name"),
                    "size": att.get("size")
                }
                for att in response.get("attachments", [])
            ] if params.include_attachments else []
        }
        
        return json.dumps(email_detail, indent=2)
    
    except Exception as e:
        return json.dumps({"status": "error", "message": str(e)})


@mcp.tool(
    name="microsoft_manage_email_folders",
    annotations={
        "title": "Manage Email Folders",
        "readOnlyHint": False,
        "destructiveHint": True,
        "idempotentHint": False,
        "openWorldHint": False
    }
)
async def manage_email_folders(params: ManageFolderInput, ctx) -> str:
    """Manage email folders - list, create, delete, or rename folders.
    
    Args:
        params (ManageFolderInput): Folder management with:
            - action: list, create, delete, rename
            - folder_name: Folder name (for create/delete)
            - new_name: New name (for rename)
    
    Returns:
        str: Operation result with folder information
    """
    client = ctx.request_context.lifespan_state["client"]
    
    try:
        if params.action.lower() == "list":
            response = await client.get(
                "/me/mailFolders",
                params={"$select": "id,displayName,childFolderCount,unreadItemCount"}
            )
            
            folders = response.get("value", [])
            return json.dumps({
                "action": "list",
                "folder_count": len(folders),
                "folders": [
                    {
                        "id": folder.get("id"),
                        "name": folder.get("displayName"),
                        "unread_count": folder.get("unreadItemCount"),
                        "subfolder_count": folder.get("childFolderCount")
                    }
                    for folder in folders
                ]
            }, indent=2)
        
        elif params.action.lower() == "create":
            if not params.folder_name:
                return json.dumps({"error": "folder_name required for create action"})
            
            response = await client.post(
                "/me/mailFolders",
                {"displayName": params.folder_name}
            )
            
            return json.dumps({
                "action": "create",
                "status": "success",
                "folder_id": response.get("id"),
                "folder_name": response.get("displayName")
            }, indent=2)
        
        elif params.action.lower() == "delete":
            if not params.folder_name:
                return json.dumps({"error": "folder_name required for delete action"})
            
            # First find folder by name
            list_response = await client.get(
                "/me/mailFolders",
                params={"$filter": f"displayName eq '{params.folder_name}'"}
            )
            
            folders = list_response.get("value", [])
            if not folders:
                return json.dumps({"error": f"Folder '{params.folder_name}' not found"})
            
            folder_id = folders[0].get("id")
            await client.delete(f"/me/mailFolders/{folder_id}")
            
            return json.dumps({
                "action": "delete",
                "status": "success",
                "folder_name": params.folder_name
            }, indent=2)
        
        elif params.action.lower() == "rename":
            if not params.folder_name or not params.new_name:
                return json.dumps({"error": "folder_name and new_name required for rename action"})
            
            # Find folder
            list_response = await client.get(
                "/me/mailFolders",
                params={"$filter": f"displayName eq '{params.folder_name}'"}
            )
            
            folders = list_response.get("value", [])
            if not folders:
                return json.dumps({"error": f"Folder '{params.folder_name}' not found"})
            
            folder_id = folders[0].get("id")
            await client.patch(f"/me/mailFolders/{folder_id}", {"displayName": params.new_name})
            
            return json.dumps({
                "action": "rename",
                "status": "success",
                "old_name": params.folder_name,
                "new_name": params.new_name
            }, indent=2)
        
        else:
            return json.dumps({"error": f"Unknown action: {params.action}"})
    
    except Exception as e:
        return json.dumps({"status": "error", "message": str(e)})


# ============================================================================
# Calendar Tools
# ============================================================================

@mcp.tool(
    name="microsoft_create_calendar_event",
    annotations={
        "title": "Create Calendar Event",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": False
    }
)
async def create_calendar_event(params: CreateEventInput, ctx) -> str:
    """Create a new calendar event with optional attendees and reminders.
    
    Args:
        params (CreateEventInput): Event details including:
            - subject: Event title
            - start_time: Start time (ISO 8601)
            - end_time: End time (ISO 8601)
            - attendees: Optional attendee emails
            - description: Optional description
            - location: Optional location
            - is_reminder: Whether to set reminder
    
    Returns:
        str: Event creation confirmation with event ID
    """
    client = ctx.request_context.lifespan_state["client"]
    
    try:
        event = {
            "subject": params.subject,
            "start": {
                "dateTime": params.start_time,
                "timeZone": "UTC"
            },
            "end": {
                "dateTime": params.end_time,
                "timeZone": "UTC"
            },
            "isReminderOn": params.is_reminder,
            "reminderMinutesBeforeStart": 15 if params.is_reminder else 0
        }
        
        if params.description:
            event["bodyPreview"] = params.description
            event["body"] = {
                "contentType": "HTML",
                "content": params.description
            }
        
        if params.location:
            event["location"] = {"displayName": params.location}
        
        if params.attendees:
            event["attendees"] = [
                {
                    "emailAddress": {"address": addr},
                    "type": "required"
                }
                for addr in params.attendees
            ]
        
        response = await client.post("/me/events", event)
        
        return json.dumps({
            "status": "success",
            "message": "Event created successfully",
            "event_id": response.get("id"),
            "subject": params.subject,
            "start_time": params.start_time,
            "end_time": params.end_time,
            "attendees_count": len(params.attendees) if params.attendees else 0
        }, indent=2)
    
    except Exception as e:
        return json.dumps({"status": "error", "message": str(e)})


@mcp.tool(
    name="microsoft_check_availability",
    annotations={
        "title": "Check Calendar Availability",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False
    }
)
async def check_availability(params: CheckAvailabilityInput, ctx) -> str:
    """Check calendar availability for a time slot and optionally for attendees.
    
    Args:
        params (CheckAvailabilityInput): Time slot and optional attendees:
            - start_time: Start time (ISO 8601)
            - end_time: End time (ISO 8601)
            - attendees: Optional attendee emails to check
    
    Returns:
        str: Availability status with conflict information
    """
    client = ctx.request_context.lifespan_state["client"]
    
    try:
        # Get user's calendar view for the time period
        response = await client.get(
            "/me/calendarview",
            params={
                "startDateTime": params.start_time,
                "endDateTime": params.end_time,
                "$select": "subject,start,end,isOrganizer"
            }
        )
        
        conflicts = response.get("value", [])
        
        availability = {
            "requested_start": params.start_time,
            "requested_end": params.end_time,
            "is_available": len(conflicts) == 0,
            "conflicts_count": len(conflicts),
            "conflicts": [
                {
                    "subject": conflict.get("subject"),
                    "start": conflict.get("start", {}).get("dateTime"),
                    "end": conflict.get("end", {}).get("dateTime"),
                    "is_organizer": conflict.get("isOrganizer")
                }
                for conflict in conflicts
            ]
        }
        
        if params.attendees:
            availability["attendees_checked"] = params.attendees
            availability["note"] = "Individual attendee availability requires separate checks through each attendee's calendar"
        
        return json.dumps(availability, indent=2)
    
    except Exception as e:
        return json.dumps({"status": "error", "message": str(e)})


@mcp.tool(
    name="microsoft_get_calendar_events",
    annotations={
        "title": "Get Calendar Events",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False
    }
)
async def get_calendar_events(params: GetEventsInput, ctx) -> str:
    """Retrieve calendar events for a date range with full details.
    
    Args:
        params (GetEventsInput): Date range and filter options:
            - start_date: Start date (YYYY-MM-DD)
            - end_date: End date (YYYY-MM-DD)
            - limit: Maximum events to return
            - include_details: Whether to include full event details
    
    Returns:
        str: List of calendar events with metadata
    """
    client = ctx.request_context.lifespan_state["client"]
    
    try:
        # Default to next 7 days if not specified
        start_date = params.start_date or datetime.now().date().isoformat()
        end_date = params.end_date or (datetime.now() + timedelta(days=7)).date().isoformat()
        
        start_datetime = f"{start_date}T00:00:00Z"
        end_datetime = f"{end_date}T23:59:59Z"
        
        select_fields = "id,subject,start,end,isOrganizer,organizer" if params.include_details else "subject,start,end"
        
        response = await client.get(
            "/me/calendarview",
            params={
                "startDateTime": start_datetime,
                "endDateTime": end_datetime,
                "$orderby": "start/dateTime",
                "$top": params.limit,
                "$select": select_fields
            }
        )
        
        events = response.get("value", [])
        
        result = {
            "date_range": f"{start_date} to {end_date}",
            "event_count": len(events),
            "events": [
                {
                    "id": event.get("id"),
                    "subject": event.get("subject"),
                    "start": event.get("start", {}).get("dateTime"),
                    "end": event.get("end", {}).get("dateTime"),
                    "is_organizer": event.get("isOrganizer"),
                    "organizer": event.get("organizer", {}).get("emailAddress", {}).get("address") if params.include_details else None
                }
                for event in events
            ]
        }
        
        return json.dumps(result, indent=2)
    
    except Exception as e:
        return json.dumps({"status": "error", "message": str(e)})


# ============================================================================
# Cloud Storage Tools
# ============================================================================

@mcp.tool(
    name="microsoft_list_cloud_files",
    annotations={
        "title": "List Cloud Files",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False
    }
)
async def list_cloud_files(params: ListFilesInput, ctx) -> str:
    """List files and folders in OneDrive, SharePoint, or Teams.
    
    Args:
        params (ListFilesInput): Location and path parameters:
            - location: onedrive, sharepoint, or teams
            - folder_path: Path to list
            - limit: Maximum items
            - recursive: Include subfolders
    
    Returns:
        str: List of files and folders with metadata
    """
    client = ctx.request_context.lifespan_state["client"]
    
    try:
        if params.location.lower() == "onedrive":
            endpoint = "/me/drive/root/children"
            if params.folder_path != "/":
                endpoint = f"/me/drive/root:/{params.folder_path}:/children"
        
        elif params.location.lower() == "sharepoint":
            endpoint = "/me/drive/root/children"
        
        elif params.location.lower() == "teams":
            endpoint = "/me/drive/root/children"
        
        else:
            return json.dumps({"error": f"Unknown location: {params.location}"})
        
        response = await client.get(
            endpoint,
            params={"$top": params.limit}
        )
        
        items = response.get("value", [])
        
        result = {
            "location": params.location,
            "path": params.folder_path,
            "item_count": len(items),
            "items": [
                {
                    "id": item.get("id"),
                    "name": item.get("name"),
                    "type": "folder" if "folder" in item else "file",
                    "size": item.get("size"),
                    "modified": item.get("lastModifiedDateTime"),
                    "web_url": item.get("webUrl")
                }
                for item in items
            ]
        }
        
        return json.dumps(result, indent=2)
    
    except Exception as e:
        return json.dumps({"status": "error", "message": str(e)})


@mcp.tool(
    name="microsoft_create_folder",
    annotations={
        "title": "Create Cloud Folder",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": False
    }
)
async def create_folder(params: CreateFolderInput, ctx) -> str:
    """Create a new folder in OneDrive, SharePoint, or Teams.
    
    Args:
        params (CreateFolderInput): Folder creation parameters:
            - folder_name: Name of new folder
            - parent_path: Parent folder path
            - location: onedrive, sharepoint, or teams
    
    Returns:
        str: Folder creation confirmation
    """
    client = ctx.request_context.lifespan_state["client"]
    
    try:
        if params.location.lower() == "onedrive":
            endpoint = "/me/drive/root/children" if params.parent_path == "/" else f"/me/drive/root:/{params.parent_path}:/children"
        else:
            endpoint = "/me/drive/root/children"
        
        folder_data = {
            "name": params.folder_name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "rename"
        }
        
        response = await client.post(endpoint, folder_data)
        
        return json.dumps({
            "status": "success",
            "message": "Folder created successfully",
            "folder_id": response.get("id"),
            "folder_name": response.get("name"),
            "web_url": response.get("webUrl")
        }, indent=2)
    
    except Exception as e:
        return json.dumps({"status": "error", "message": str(e)})


# ============================================================================
# Server Entry Point
# ============================================================================

if __name__ == "__main__":
    import sys
    
    # Check for required environment variables
    required_vars = ["MICROSOFT_CLIENT_ID", "MICROSOFT_CLIENT_SECRET", "MICROSOFT_TENANT_ID"]
    missing_vars = [var for var in required_vars if not os.getenv(var)]
    
    if missing_vars:
        print(f"Error: Missing required environment variables: {', '.join(missing_vars)}", file=sys.stderr)
        print("\nSetup Instructions:", file=sys.stderr)
        print("1. Register an Azure Application at https://portal.azure.com", file=sys.stderr)
        print("2. Add API permissions for Microsoft Graph:", file=sys.stderr)
        for scope in AZURE_SCOPE:
            print(f"   - {scope}", file=sys.stderr)
        print("3. Set environment variables:", file=sys.stderr)
        for var in required_vars:
            print(f"   export {var}=<your-value>", file=sys.stderr)
        sys.exit(1)
    
    mcp.run()
