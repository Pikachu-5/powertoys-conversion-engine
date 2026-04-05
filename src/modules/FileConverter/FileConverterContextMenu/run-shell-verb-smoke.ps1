param(
    [string]$TestDirectory = "x64\Debug\WinUI3Apps\FileConverterSmokeTest",
    [string]$InputFileName = "sample.bmp",
    [string]$ExpectedOutputFileName = "sample_converted.png",
    [string]$VerbName = "Convert to PNG",
    [int]$InvokeTimeoutMs = 20000,
    [int]$OutputWaitTimeoutMs = 10000
)

$ErrorActionPreference = "Stop"

$resolvedTestDir = (Resolve-Path $TestDirectory).Path
$outputPath = Join-Path $resolvedTestDir $ExpectedOutputFileName
if (Test-Path $outputPath)
{
    Remove-Item $outputPath -Force
}

$code = @"
using System;
using System.Reflection;
using System.Threading;

public static class ShellVerbRunner
{
    public static string Invoke(string directoryPath, string fileName, string targetVerb, int timeoutMs)
    {
        string result = "Unknown";
        Exception error = null;
        bool completed = false;

        Thread thread = new Thread(() =>
        {
            try
            {
                Type shellType = Type.GetTypeFromProgID("Shell.Application");
                object shell = Activator.CreateInstance(shellType);
                object folder = shellType.InvokeMember("NameSpace", BindingFlags.InvokeMethod, null, shell, new object[] { directoryPath });
                if (folder == null)
                {
                    result = "Folder not found";
                    return;
                }

                Type folderType = folder.GetType();
                object item = folderType.InvokeMember("ParseName", BindingFlags.InvokeMethod, null, folder, new object[] { fileName });
                if (item == null)
                {
                    result = "Item not found";
                    return;
                }

                Type itemType = item.GetType();
                object verbs = itemType.InvokeMember("Verbs", BindingFlags.InvokeMethod, null, item, null);
                Type verbsType = verbs.GetType();
                int count = (int)verbsType.InvokeMember("Count", BindingFlags.GetProperty, null, verbs, null);

                for (int index = 0; index < count; index++)
                {
                    object verb = verbsType.InvokeMember("Item", BindingFlags.InvokeMethod, null, verbs, new object[] { index });
                    if (verb == null)
                    {
                        continue;
                    }

                    Type verbType = verb.GetType();
                    string name = (verbType.InvokeMember("Name", BindingFlags.GetProperty, null, verb, null) as string ?? string.Empty)
                        .Replace("&", string.Empty)
                        .Trim();

                    if (!string.Equals(name, targetVerb, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    verbType.InvokeMember("DoIt", BindingFlags.InvokeMethod, null, verb, null);
                    result = "Invoked";
                    return;
                }

                result = "Verb not found";
            }
            catch (Exception ex)
            {
                Exception current = ex;
                string details = string.Empty;
                while (current != null)
                {
                    details += current.GetType().FullName + ": " + current.Message + Environment.NewLine;
                    current = current.InnerException;
                }

                error = new Exception(details.Trim());
            }
            finally
            {
                completed = true;
            }
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join(timeoutMs);

        if (!completed)
        {
            return "Timeout";
        }

        if (error != null)
        {
            return "Error: " + error.Message;
        }

        return result;
    }
}
"@

Add-Type -TypeDefinition $code -Language CSharp
$invokeResult = [ShellVerbRunner]::Invoke($resolvedTestDir, $InputFileName, $VerbName, $InvokeTimeoutMs)
Write-Host "Invoke result: $invokeResult"

if ($invokeResult -ne "Invoked")
{
    throw "Verb invocation failed: $invokeResult"
}

$waited = 0
$step = 250
while ($waited -lt $OutputWaitTimeoutMs -and -not (Test-Path $outputPath))
{
    Start-Sleep -Milliseconds $step
    $waited += $step
}

if (-not (Test-Path $outputPath))
{
    throw "Output file was not created: $outputPath"
}

$item = Get-Item $outputPath
Write-Host "Created: $($item.FullName)"
Write-Host "Size: $($item.Length)"
