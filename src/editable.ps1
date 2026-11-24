[CmdletBinding(DefaultParameterSetName = 'Convert')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Convert')][string]$InputPath,
    [Parameter(Mandatory, ParameterSetName = 'Convert')][string]$OutputPath,
    [Parameter(ParameterSetName = 'Convert')][ValidateRange(1, 86400)][int]$TimeoutSeconds = 300,
    [Parameter(Mandatory, ParameterSetName = 'Worker')][string]$JobDirectory
)

$ErrorActionPreference = 'Stop'
$utf8 = New-Object System.Text.UTF8Encoding($false)
[Console]::OutputEncoding = $utf8

function Invoke-WordConversion([string]$directory) {
    $config = [IO.File]::ReadAllText((Join-Path $directory 'request.json')) | ConvertFrom-Json
    $word = $null
    $ownsWord = $false
    $blank = $null
    $pdf = $null
    try {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class DocumentWindow {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr window, out int processId);
}
'@
        $existing = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
        try { $word = New-Object -ComObject Word.Application }
        catch { throw 'Microsoft Word desktop could not start. Install and activate Word 2013 or later.' }
        # Word exposes the HWND on a document window, not on Application.
        $blank = $word.Documents.Add()
        $wordProcessId = 0
        [void][DocumentWindow]::GetWindowThreadProcessId([IntPtr]$blank.ActiveWindow.Hwnd, [ref]$wordProcessId)
        $process = Get-Process -Id $wordProcessId
        if ($process.ProcessName -ne 'WINWORD' -or $existing -contains $wordProcessId) {
            throw 'Word did not start an isolated instance. Please retry.'
        }
        $ownsWord = $true
        $identity = @{ id = $process.Id; started = $process.StartTime.ToUniversalTime().Ticks.ToString() }
        [IO.File]::WriteAllText((Join-Path $directory 'word.json'), ($identity | ConvertTo-Json -Compress))
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $word.AutomationSecurity = 3
        $blank.Close([ref]0)
        $blank = $null
        # Open read-only, disable format confirmation, and omit from recent files.
        $pdf = $word.Documents.Open([string]$config.input, $false, $true, $false)
        $target = Join-Path $directory 'converted.docx'
        $pdf.SaveAs2([string]$target, 12)
    } finally {
        if ($null -ne $pdf) { try { $pdf.Close([ref]0) } catch {} }
        if ($null -ne $blank) { try { $blank.Close([ref]0) } catch {} }
        if ($ownsWord) { try { $word.Quit([ref]0) } catch {} }
        if ($null -ne $word) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) }
    }
}

function Stop-IsolatedWord([string]$directory) {
    $identityPath = Join-Path $directory 'word.json'
    if (![IO.File]::Exists($identityPath)) { return }
    try {
        $identity = [IO.File]::ReadAllText($identityPath) | ConvertFrom-Json
        $process = Get-Process -Id $identity.id -ErrorAction SilentlyContinue
        if ($null -ne $process -and $process.ProcessName -eq 'WINWORD' -and
            $process.StartTime.ToUniversalTime().Ticks.ToString() -eq $identity.started) {
            $process.Kill()
            [void]$process.WaitForExit(5000)
        }
    } catch {}
}

function Initialize-PromptReader {
    Add-Type -AssemblyName UIAutomationClient
    Add-Type -TypeDefinition @'
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
public static class PdfReflowPrompt {
    private delegate bool Visit(IntPtr h, IntPtr parameter);
    [DllImport("user32.dll")] private static extern bool EnumWindows(Visit visit, IntPtr parameter);
    [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr h, uint command);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr h, out int pid);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] private static extern int GetClassName(IntPtr h, StringBuilder name, int length);
    public static IntPtr Find(int ownerProcess) {
        IntPtr found = IntPtr.Zero;
        EnumWindows(delegate(IntPtr window, IntPtr unused) {
            int owner, pid;
            GetWindowThreadProcessId(GetWindow(window, 4), out owner);
            if (owner != ownerProcess) return true;
            var name = new StringBuilder(64);
            GetClassName(window, name, name.Capacity);
            if (name.ToString() != "NUIDialog") return true;
            GetWindowThreadProcessId(window, out pid);
            try {
                using (var process = Process.GetProcessById(pid)) {
                    if (!process.ProcessName.Equals("PDFREFLOW", StringComparison.OrdinalIgnoreCase)) return true;
                }
            } catch (ArgumentException) { return true; }
            found = window;
            return false;
        }, IntPtr.Zero);
        return found;
    }
}
'@
}

function Confirm-ReflowNotice([int]$wordProcessId) {
    $handle = [PdfReflowPrompt]::Find($wordProcessId)
    if ($handle -eq [IntPtr]::Zero) { return }
    try {
        $dialog = [System.Windows.Automation.AutomationElement]::FromHandle($handle)
        $children = $dialog.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)
        $hasLabel = $false
        $hasCheckbox = $false
        $ok = $null
        foreach ($child in $children) {
            $item = $child.Current
            if ($item.AutomationId -eq '4001' -and $item.ClassName -eq 'NetUILabel' -and $item.Name -match 'PDF' -and $item.Name -match 'Word') { $hasLabel = $true }
            if ($item.ClassName -eq 'NetUICheckbox') { $hasCheckbox = $true }
            if ($item.AutomationId -eq '1' -and $item.ClassName -eq 'NetUIButton') { $ok = $child }
        }
        if ($hasLabel -and $hasCheckbox -and $ok) {
            $ok.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
        }
    } catch [System.Windows.Automation.ElementNotAvailableException] {}
}

if ($PSCmdlet.ParameterSetName -eq 'Worker') {
    try { Invoke-WordConversion $JobDirectory; exit 0 }
    catch {
        [IO.File]::WriteAllText((Join-Path $JobDirectory 'error.txt'), $_.Exception.Message, $utf8)
        exit 1
    }
}

$jobPath = $null
$worker = $null
try {
    $source = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($InputPath)
    $target = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
    if (![IO.File]::Exists($source) -or [IO.Path]::GetExtension($source) -ne '.pdf') {
        throw 'Select an existing PDF file.'
    }
    if ([IO.Path]::GetExtension($target) -ne '.docx') { throw 'The output file must end in .docx.' }
    if ([IO.File]::Exists($target)) { throw 'The output file already exists. Choose a new filename.' }
    $parent = [IO.Path]::GetDirectoryName($target)
    if (![IO.Directory]::Exists($parent)) { throw 'The output folder does not exist.' }
    $jobPath = Join-Path $parent ('.pdf-to-word-' + [Guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($jobPath)
    [IO.File]::WriteAllText((Join-Path $jobPath 'request.json'), (@{ input = $source } | ConvertTo-Json -Compress), $utf8)
    $psi = New-Object Diagnostics.ProcessStartInfo
    $psi.FileName = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $psi.Arguments = '-NoLogo -NoProfile -NonInteractive -STA -ExecutionPolicy Bypass -File "' + $PSCommandPath + '" -JobDirectory "' + $jobPath + '"'
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    Initialize-PromptReader
    $worker = [Diagnostics.Process]::Start($psi)
    $elapsed = [Diagnostics.Stopwatch]::StartNew()
    $identityPath = Join-Path $jobPath 'word.json'
    $identity = $null
    while (!$worker.WaitForExit(500)) {
        if ($elapsed.Elapsed.TotalSeconds -ge $TimeoutSeconds) {
            $worker.Kill()
            [void]$worker.WaitForExit(5000)
            Stop-IsolatedWord $jobPath
            throw "Word conversion timed out after $TimeoutSeconds seconds."
        }
        if (!$identity -and [IO.File]::Exists($identityPath)) {
            try { $identity = [IO.File]::ReadAllText($identityPath) | ConvertFrom-Json } catch {}
        }
        if ($identity) { try { Confirm-ReflowNotice $identity.id } catch {} }
    }
    if ($worker.ExitCode -ne 0) {
        $errorFile = Join-Path $jobPath 'error.txt'
        if ([IO.File]::Exists($errorFile)) { throw [IO.File]::ReadAllText($errorFile) }
        throw 'The Word conversion process exited unexpectedly.'
    }
    $converted = Join-Path $jobPath 'converted.docx'
    [IO.File]::Move($converted, $target)
    exit 0
} catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
} finally {
    if ($null -ne $worker) { $worker.Dispose() }
    if ($jobPath -and [IO.Directory]::Exists($jobPath)) {
        # Only delete known child files in our newly-created private staging directory.
        foreach ($name in @('request.json', 'word.json', 'error.txt', 'converted.docx')) {
            try { [IO.File]::Delete((Join-Path $jobPath $name)) } catch {}
        }
        try { [IO.Directory]::Delete($jobPath) } catch {}
    }
}
