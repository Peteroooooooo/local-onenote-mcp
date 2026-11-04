"""Secure local bridge to the Windows OneNote COM API.

The OneNote COM interface is not reliably usable from Python automation
libraries (pywin32 cannot build its makepy wrapper for OneNote.Application), but
the desktop application is scriptable through PowerShell. This module keeps one
long-lived PowerShell process running a fixed script and talks to it with one
JSON object per line over stdin/stdout. User data travels only inside those
JSON messages and is never interpolated into script text.
"""

from __future__ import annotations

import atexit
import base64
import collections
import json
import queue
import subprocess
import threading
import time
from typing import Any


POWERSHELL_BRIDGE = r'''
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
$utf8 = New-Object System.Text.UTF8Encoding($false)
$stdin = New-Object System.IO.StreamReader([Console]::OpenStandardInput(), $utf8)
$stdout = New-Object System.IO.StreamWriter([Console]::OpenStandardOutput(), $utf8)
$stdout.AutoFlush = $true
$script:onenote = $null

# RPC_E_DISCONNECTED, RPC_S_SERVER_UNAVAILABLE, CO_E_OBJNOTCONNECTED: the call never reached OneNote.
$reconnectHResults = @(-2147417848, -2147023174, -2147220995)

function Get-OneNote {
    if ($null -eq $script:onenote) {
        $script:onenote = New-Object -ComObject OneNote.Application
    }
    return $script:onenote
}

function Get-HResult($exception) {
    $current = $exception
    while ($null -ne $current) {
        if ($current -is [System.Runtime.InteropServices.COMException]) {
            return $current.HResult
        }
        $current = $current.InnerException
    }
    return $exception.HResult
}

function Invoke-Operation([string]$op, $p) {
    $onenote = Get-OneNote
    switch ($op) {
        "get_hierarchy" {
            $xml = ""
            $null = $onenote.GetHierarchy([string]$p.start_id, [int]$p.scope, [ref]$xml, [int]$p.schema)
            return @{ xml = $xml }
        }
        "open_hierarchy" {
            $objectId = ""
            $null = $onenote.OpenHierarchy([string]$p.path, [string]$p.relative_to_id, [ref]$objectId, [int]$p.create_file_type)
            return @{ object_id = $objectId }
        }
        "update_hierarchy" {
            $null = $onenote.UpdateHierarchy([string]$p.xml, [int]$p.schema)
            return @{ updated = $true }
        }
        "delete_hierarchy" {
            $null = $onenote.DeleteHierarchy([string]$p.object_id, 0, [bool]$p.permanently)
            return @{ deleted = $true }
        }
        "close_notebook" {
            $null = $onenote.CloseNotebook([string]$p.notebook_id, [bool]$p.force)
            return @{ closed = $true }
        }
        "get_hierarchy_parent" {
            $parentId = ""
            $null = $onenote.GetHierarchyParent([string]$p.object_id, [ref]$parentId)
            return @{ parent_id = $parentId }
        }
        "get_special_location" {
            $location = ""
            $null = $onenote.GetSpecialLocation([int]$p.location, [ref]$location)
            return @{ path = $location }
        }
        "create_new_page" {
            $pageId = ""
            $null = $onenote.CreateNewPage([string]$p.section_id, [ref]$pageId, [int]$p.new_page_style)
            return @{ page_id = $pageId }
        }
        "get_page_content" {
            $xml = ""
            $null = $onenote.GetPageContent([string]$p.page_id, [ref]$xml, [int]$p.page_info, [int]$p.schema)
            return @{ xml = $xml }
        }
        "update_page_content" {
            $null = $onenote.UpdatePageContent([string]$p.xml, 0, [int]$p.schema, [bool]$p.force)
            return @{ updated = $true }
        }
        "delete_page_content" {
            $null = $onenote.DeletePageContent([string]$p.page_id, [string]$p.object_id, 0, [bool]$p.force)
            return @{ deleted = $true }
        }
        "get_binary_page_content" {
            $content = ""
            $null = $onenote.GetBinaryPageContent([string]$p.page_id, [string]$p.callback_id, [ref]$content)
            return @{ base64 = $content }
        }
        "publish" {
            $null = $onenote.Publish([string]$p.object_id, [string]$p.target_path, [int]$p.format, "")
            return @{ path = [string]$p.target_path }
        }
        "find_pages" {
            $xml = ""
            $null = $onenote.FindPages([string]$p.start_id, [string]$p.query, [ref]$xml, [bool]$p.include_unindexed, [bool]$p.display, [int]$p.schema)
            return @{ xml = $xml }
        }
        "find_meta" {
            $xml = ""
            $null = $onenote.FindMeta([string]$p.start_id, [string]$p.name, [ref]$xml, [bool]$p.include_unindexed, [int]$p.schema)
            return @{ xml = $xml }
        }
        "get_hyperlink" {
            $link = ""
            $null = $onenote.GetHyperlinkToObject([string]$p.object_id, [string]$p.page_content_object_id, [ref]$link)
            return @{ hyperlink = $link }
        }
        "get_web_hyperlink" {
            $link = ""
            $null = $onenote.GetWebHyperlinkToObject([string]$p.object_id, [string]$p.page_content_object_id, [ref]$link)
            return @{ hyperlink = $link }
        }
        "navigate_to" {
            $null = $onenote.NavigateTo([string]$p.object_id, [string]$p.page_content_object_id, [bool]$p.new_window)
            return @{ navigated = $true }
        }
        "navigate_to_url" {
            $null = $onenote.NavigateToUrl([string]$p.url, [bool]$p.new_window)
            return @{ navigated = $true }
        }
        "sync_hierarchy" {
            $null = $onenote.SyncHierarchy([string]$p.object_id)
            return @{ synced = $true }
        }
        "merge_sections" {
            $null = $onenote.MergeSections([string]$p.source_section_id, [string]$p.destination_section_id)
            return @{ merged = $true }
        }
        "render_pdf" {
            return Render-Pdf ([string]$p.path) ([int]$p.width) ([int]$p.max_pages)
        }
        "ensure_window" {
            # UpdateHierarchy crashes OneNote (onmain.dll access violation) on any page reorder while OneNote runs
            # only as a COM server ("ONENOTE.EXE -Embedding"). With the app itself running it works, so start the
            # app first: minimized, and through WMI so it is not a child of this process and outlives the server.
            $uiPids = {
                @(Get-CimInstance Win32_Process -Filter "Name = 'ONENOTE.EXE'" | Where-Object { [string]$_.CommandLine -notmatch '(?i)[-/]embedding' } | ForEach-Object { [int]$_.ProcessId })
            }
            $hasUi = {
                @(& $uiPids | Where-Object { $proc = Get-Process -Id $_ -ErrorAction SilentlyContinue; $proc -and $proc.MainWindowHandle -ne 0 }).Count -gt 0
            }
            if (& $hasUi) { return @{ window = $true; started = $false } }
            $exe = (Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OneNote.exe' -ErrorAction SilentlyContinue).'(default)'
            if (-not $exe -or -not (Test-Path -LiteralPath $exe)) { return @{ window = $false; started = $false } }
            $startup = New-CimInstance -ClassName Win32_ProcessStartup -ClientOnly -Property @{ ShowWindow = [uint16]7 }
            $null = Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{ CommandLine = "`"$exe`""; ProcessStartupInformation = $startup }
            $deadline = (Get-Date).AddSeconds([int]$p.timeout_seconds)
            $seen = 0
            while ((Get-Date) -lt $deadline) {
                # The window must stay up: a second ONENOTE.EXE can hand over to the running one and exit.
                if (& $hasUi) { $seen++ } else { $seen = 0 }
                if ($seen -ge 8) { break }
                Start-Sleep -Milliseconds 250
            }
            if ($seen -lt 8) { return @{ window = $false; started = $true } }
            # Reconnect, so later calls are served by the instance that owns the window.
            if ($null -ne $script:onenote) {
                $null = [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:onenote)
                $script:onenote = $null
                [GC]::Collect()
                [GC]::WaitForPendingFinalizers()
            }
            return @{ window = $true; started = $true }
        }
        "image_to_png" {
            # OneNote's XML takes only PNG, JPEG and EMF images; GIF, BMP and TIFF are converted first.
            Add-Type -AssemblyName System.Drawing
            $image = [System.Drawing.Image]::FromFile([string]$p.path)
            $stream = New-Object System.IO.MemoryStream
            try {
                $image.Save($stream, [System.Drawing.Imaging.ImageFormat]::Png)
                return @{ base64 = [Convert]::ToBase64String($stream.ToArray()); width = $image.Width; height = $image.Height }
            } finally {
                $stream.Dispose()
                $image.Dispose()
            }
        }
        default {
            throw "Unsupported OneNote bridge operation: $op"
        }
    }
}

# Windows.Data.Pdf (WinRT) turns a published PDF into JPEG images, one per printed page.
function Initialize-WinRT {
    if ($null -ne $script:asTask) { return }
    Add-Type -AssemblyName System.Runtime.WindowsRuntime
    $null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
    $null = [Windows.Graphics.Imaging.BitmapEncoder, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
    $null = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
    $methods = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq "AsTask" -and $_.GetParameters().Count -eq 1 }
    $script:asAction = $methods | Where-Object { $_.GetParameters()[0].ParameterType.Name -eq "IAsyncAction" } | Select-Object -First 1
    $script:asTask = $methods | Where-Object { $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' } | Select-Object -First 1
}

function Wait-WinRT($operation, [Type]$resultType) {
    if ($null -eq $resultType) {
        $script:asAction.Invoke($null, @($operation)).Wait()
        return
    }
    $task = $script:asTask.MakeGenericMethod($resultType).Invoke($null, @($operation))
    $task.Wait()
    return $task.Result
}

function Render-Pdf([string]$path, [int]$width, [int]$maxPages) {
    Initialize-WinRT
    $bytes = [IO.File]::ReadAllBytes($path)
    $source = [System.IO.WindowsRuntimeStreamExtensions]::AsRandomAccessStream((New-Object IO.MemoryStream(, $bytes)))
    $document = Wait-WinRT ([Windows.Data.Pdf.PdfDocument]::LoadFromStreamAsync($source)) ([Windows.Data.Pdf.PdfDocument])
    $images = @()
    for ($i = 0; $i -lt [Math]::Min([int]$document.PageCount, $maxPages); $i++) {
        $page = $document.GetPage($i)
        $options = New-Object Windows.Data.Pdf.PdfPageRenderOptions
        $options.DestinationWidth = $width
        $options.BitmapEncoderId = [Windows.Graphics.Imaging.BitmapEncoder]::JpegEncoderId
        $output = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
        Wait-WinRT ($page.RenderToStreamAsync($output, $options)) $null
        $memory = New-Object IO.MemoryStream
        [System.IO.WindowsRuntimeStreamExtensions]::AsStreamForRead($output.GetInputStreamAt(0)).CopyTo($memory)
        $images += [Convert]::ToBase64String($memory.ToArray())
        $page.Dispose()
        $output.Dispose()
    }
    return @{ pages = [int]$document.PageCount; images = $images }
}

while ($true) {
    $line = $stdin.ReadLine()
    if ($null -eq $line) { break }
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    $requestId = $null
    try {
        $request = $line | ConvertFrom-Json
        $requestId = $request.id
        $op = [string]$request.operation
        if ($op -eq "ping") {
            $data = @{ pong = $true; value = [string]$request.params.value }
        } else {
            try {
                $data = Invoke-Operation $op $request.params
            } catch {
                if ($reconnectHResults -notcontains (Get-HResult $_.Exception)) { throw }
                $script:onenote = $null
                $data = Invoke-Operation $op $request.params
            }
        }
        $response = @{ id = $requestId; ok = $true; data = $data; error = $null }
    } catch {
        $response = @{
            id = $requestId
            ok = $false
            data = $null
            error = @{
                message = $_.Exception.Message
                hresult = Get-HResult $_.Exception
                category = [string]$_.CategoryInfo
            }
        }
    }
    $stdout.WriteLine(($response | ConvertTo-Json -Depth 20 -Compress))
}
'''


def _encoded_command(script: str) -> str:
    return base64.b64encode(script.encode("utf-16-le")).decode("ascii")


class OneNoteBridgeError(RuntimeError):
    """Raised when the local OneNote COM bridge fails."""

    def __init__(self, message: str, *, hresult: int | None = None) -> None:
        super().__init__(message)
        self.hresult = hresult


class OneNoteBridge:
    """Execute fixed OneNote COM operations through one persistent PowerShell process.

    Calls are serialized. A call that times out kills the process; the next call
    starts a fresh one.
    """

    def __init__(self, timeout_seconds: int = 90) -> None:
        self.timeout_seconds = timeout_seconds
        self._lock = threading.Lock()
        self._process: subprocess.Popen[bytes] | None = None
        self._responses: queue.Queue[bytes | None] = queue.Queue()
        self._stderr: collections.deque[str] = collections.deque(maxlen=40)
        self._next_id = 0
        atexit.register(self.close)

    def call(self, operation: str, **params: Any) -> dict[str, Any]:
        with self._lock:
            process = self._ensure_process()
            self._next_id += 1
            request_id = self._next_id
            line = json.dumps({"id": request_id, "operation": operation, "params": params}, ensure_ascii=True)
            try:
                assert process.stdin is not None
                process.stdin.write(line.encode("ascii") + b"\n")
                process.stdin.flush()
            except OSError as exc:
                self._stop_process()
                raise OneNoteBridgeError(f"PowerShell bridge is not running: {exc}{self._stderr_tail()}") from exc
            response = self._read_response(request_id)

        if not response.get("ok"):
            err = response.get("error") or {}
            raise OneNoteBridgeError(
                err.get("message") or "OneNote COM operation failed.",
                hresult=err.get("hresult"),
            )
        data = response.get("data")
        return data if isinstance(data, dict) else {"value": data}

    def close(self) -> None:
        with self._lock:
            self._stop_process()

    def _read_response(self, request_id: int) -> dict[str, Any]:
        deadline = time.monotonic() + self.timeout_seconds
        while True:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                self._stop_process()
                raise OneNoteBridgeError(f"OneNote COM operation timed out after {self.timeout_seconds} seconds.")
            try:
                raw = self._responses.get(timeout=remaining)
            except queue.Empty:
                continue
            if raw is None:
                self._stop_process()
                raise OneNoteBridgeError(f"PowerShell bridge exited unexpectedly.{self._stderr_tail()}")
            try:
                response = json.loads(raw.decode("utf-8-sig"))
            except (UnicodeDecodeError, json.JSONDecodeError):
                self._stderr.append(raw.decode("utf-8", "replace")[:500])
                continue
            if isinstance(response, dict) and response.get("id") == request_id:
                return response

    def _ensure_process(self) -> subprocess.Popen[bytes]:
        if self._process is not None and self._process.poll() is None:
            return self._process
        self._stop_process()
        self._responses = queue.Queue()
        self._stderr.clear()
        process = subprocess.Popen(
            [
                "powershell.exe",
                "-NoLogo",
                "-NoProfile",
                "-NonInteractive",
                "-EncodedCommand",
                _encoded_command(POWERSHELL_BRIDGE),
            ],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        threading.Thread(target=self._pump_stdout, args=(process, self._responses), daemon=True).start()
        threading.Thread(target=self._pump_stderr, args=(process, self._stderr), daemon=True).start()
        self._process = process
        return process

    def _stop_process(self) -> None:
        process, self._process = self._process, None
        if process is None:
            return
        try:
            if process.stdin is not None:
                process.stdin.close()
        except OSError:
            pass
        try:
            process.wait(timeout=2)
        except subprocess.TimeoutExpired:
            process.kill()
            try:
                process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                pass

    @staticmethod
    def _pump_stdout(process: subprocess.Popen[bytes], responses: queue.Queue[bytes | None]) -> None:
        assert process.stdout is not None
        for raw in process.stdout:
            responses.put(raw)
        responses.put(None)

    @staticmethod
    def _pump_stderr(process: subprocess.Popen[bytes], lines: collections.deque[str]) -> None:
        assert process.stderr is not None
        for raw in process.stderr:
            lines.append(raw.decode("utf-8", "replace").rstrip())

    def _stderr_tail(self) -> str:
        tail = "\n".join(line for line in self._stderr if line).strip()
        return f"\n{tail[-2000:]}" if tail else ""
