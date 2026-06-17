"""Secure local bridge to the Windows OneNote COM API.

The OneNote COM type library is not reliably registered on all machines. The
desktop application is still scriptable through PowerShell, so this module uses
a fixed PowerShell program as a narrow COM bridge. User data is passed only via
JSON temp files and never interpolated into the script text.
"""

from __future__ import annotations

import json
import os
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any


POWERSHELL_BRIDGE = r'''
$ErrorActionPreference = "Stop"

function New-Ok($data) {
    return @{ ok = $true; data = $data; error = $null }
}

function New-Err($err) {
    $ex = $err.Exception
    return @{
        ok = $false
        data = $null
        error = @{
            message = $ex.Message
            hresult = $ex.HResult
            category = [string]$err.CategoryInfo
        }
    }
}

try {
    $requestPath = $env:LOCAL_ONENOTE_MCP_REQUEST
    $responsePath = $env:LOCAL_ONENOTE_MCP_RESPONSE
    if ([string]::IsNullOrWhiteSpace($requestPath) -or [string]::IsNullOrWhiteSpace($responsePath)) {
        throw "Bridge request/response paths are not set."
    }

    $request = Get-Content -LiteralPath $requestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $op = [string]$request.operation
    $p = $request.params
    $onenote = New-Object -ComObject OneNote.Application
    $data = $null

    switch ($op) {
        "get_hierarchy" {
            $xml = ""
            $onenote.GetHierarchy([string]$p.start_id, [int]$p.scope, [ref]$xml, [int]$p.schema)
            $data = @{ xml = $xml }
        }
        "open_hierarchy" {
            $objectId = ""
            $onenote.OpenHierarchy([string]$p.path, [string]$p.relative_to_id, [ref]$objectId, [int]$p.create_file_type)
            $data = @{ object_id = $objectId }
        }
        "update_hierarchy" {
            $onenote.UpdateHierarchy([string]$p.xml, [int]$p.schema)
            $data = @{ updated = $true }
        }
        "delete_hierarchy" {
            $onenote.DeleteHierarchy([string]$p.object_id, 0, [bool]$p.permanently)
            $data = @{ deleted = $true }
        }
        "close_notebook" {
            $onenote.CloseNotebook([string]$p.notebook_id, [bool]$p.force)
            $data = @{ closed = $true }
        }
        "get_hierarchy_parent" {
            $parentId = ""
            $onenote.GetHierarchyParent([string]$p.object_id, [ref]$parentId)
            $data = @{ parent_id = $parentId }
        }
        "get_special_location" {
            $location = ""
            $onenote.GetSpecialLocation([int]$p.location, [ref]$location)
            $data = @{ path = $location }
        }
        "create_new_page" {
            $pageId = ""
            $onenote.CreateNewPage([string]$p.section_id, [ref]$pageId, [int]$p.new_page_style)
            $data = @{ page_id = $pageId }
        }
        "get_page_content" {
            $xml = ""
            $onenote.GetPageContent([string]$p.page_id, [ref]$xml, [int]$p.page_info, [int]$p.schema)
            $data = @{ xml = $xml }
        }
        "update_page_content" {
            $onenote.UpdatePageContent([string]$p.xml, 0, [int]$p.schema, [bool]$p.force)
            $data = @{ updated = $true }
        }
        "delete_page_content" {
            $onenote.DeletePageContent([string]$p.page_id, [string]$p.object_id, 0, [bool]$p.force)
            $data = @{ deleted = $true }
        }
        "get_binary_page_content" {
            $content = ""
            $onenote.GetBinaryPageContent([string]$p.page_id, [string]$p.callback_id, [ref]$content)
            $data = @{ base64 = $content }
        }
        "publish" {
            $onenote.Publish([string]$p.object_id, [string]$p.target_path, [int]$p.format, "")
            $data = @{ path = [string]$p.target_path }
        }
        "find_pages" {
            $xml = ""
            $onenote.FindPages([string]$p.start_id, [string]$p.query, [ref]$xml, [bool]$p.include_unindexed, [bool]$p.display, [int]$p.schema)
            $data = @{ xml = $xml }
        }
        "find_meta" {
            $xml = ""
            $onenote.FindMeta([string]$p.start_id, [string]$p.name, [ref]$xml, [bool]$p.include_unindexed, [int]$p.schema)
            $data = @{ xml = $xml }
        }
        "get_hyperlink" {
            $link = ""
            $onenote.GetHyperlinkToObject([string]$p.object_id, [string]$p.page_content_object_id, [ref]$link)
            $data = @{ hyperlink = $link }
        }
        "get_web_hyperlink" {
            $link = ""
            $onenote.GetWebHyperlinkToObject([string]$p.object_id, [string]$p.page_content_object_id, [ref]$link)
            $data = @{ hyperlink = $link }
        }
        "navigate_to" {
            $onenote.NavigateTo([string]$p.object_id, [string]$p.page_content_object_id, [bool]$p.new_window)
            $data = @{ navigated = $true }
        }
        "navigate_to_url" {
            $onenote.NavigateToUrl([string]$p.url, [bool]$p.new_window)
            $data = @{ navigated = $true }
        }
        "sync_hierarchy" {
            $onenote.SyncHierarchy([string]$p.object_id)
            $data = @{ synced = $true }
        }
        "merge_sections" {
            $onenote.MergeSections([string]$p.source_section_id, [string]$p.destination_section_id)
            $data = @{ merged = $true }
        }
        "set_filing_location" {
            $onenote.SetFilingLocation([int]$p.filing_location, [int]$p.filing_location_type, [string]$p.section_or_page_id)
            $data = @{ updated = $true }
        }
        default {
            throw "Unsupported OneNote bridge operation: $op"
        }
    }

    $response = New-Ok $data
} catch {
    $response = New-Err $_
}

$response | ConvertTo-Json -Depth 100 -Compress | Set-Content -LiteralPath $env:LOCAL_ONENOTE_MCP_RESPONSE -Encoding UTF8
'''


class OneNoteBridgeError(RuntimeError):
    """Raised when the local OneNote COM bridge fails."""

    def __init__(self, message: str, *, hresult: int | None = None) -> None:
        super().__init__(message)
        self.hresult = hresult


@dataclass(frozen=True)
class OneNoteBridge:
    """Execute fixed OneNote COM operations through a local PowerShell bridge."""

    timeout_seconds: int = 90

    def call(self, operation: str, **params: Any) -> dict[str, Any]:
        request = {"operation": operation, "params": params}
        req_path = self._write_temp_json(request)
        resp_path = self._reserve_temp_path()
        env = os.environ.copy()
        env["LOCAL_ONENOTE_MCP_REQUEST"] = str(req_path)
        env["LOCAL_ONENOTE_MCP_RESPONSE"] = str(resp_path)
        try:
            completed = subprocess.run(
                ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", "-"],
                input=POWERSHELL_BRIDGE,
                text=True,
                capture_output=True,
                timeout=self.timeout_seconds,
                env=env,
            )
            if completed.returncode != 0:
                stderr = completed.stderr.strip()
                stdout = completed.stdout.strip()
                raise OneNoteBridgeError(stderr or stdout or "PowerShell bridge failed.")
            if not resp_path.exists():
                raise OneNoteBridgeError("PowerShell bridge did not write a response.")
            response = json.loads(resp_path.read_text(encoding="utf-8-sig"))
            if not response.get("ok"):
                err = response.get("error") or {}
                raise OneNoteBridgeError(
                    err.get("message") or "OneNote COM operation failed.",
                    hresult=err.get("hresult"),
                )
            data = response.get("data")
            return data if isinstance(data, dict) else {"value": data}
        except subprocess.TimeoutExpired as exc:
            raise OneNoteBridgeError(
                f"OneNote COM operation timed out after {self.timeout_seconds} seconds."
            ) from exc
        finally:
            self._remove_quietly(req_path)
            self._remove_quietly(resp_path)

    @staticmethod
    def _write_temp_json(payload: dict[str, Any]) -> Path:
        handle = tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            prefix="local-onenote-mcp-",
            suffix=".json",
            delete=False,
        )
        with handle:
            json.dump(payload, handle, ensure_ascii=False)
        return Path(handle.name)

    @staticmethod
    def _reserve_temp_path() -> Path:
        handle = tempfile.NamedTemporaryFile(
            prefix="local-onenote-mcp-",
            suffix=".response.json",
            delete=False,
        )
        path = Path(handle.name)
        handle.close()
        path.unlink(missing_ok=True)
        return path

    @staticmethod
    def _remove_quietly(path: Path) -> None:
        try:
            path.unlink(missing_ok=True)
        except OSError:
            pass
