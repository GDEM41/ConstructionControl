using System.Diagnostics;

string GetArg(string name)
{
    for (var i = 0; i < args.Length - 1; i++)
    {
        if (string.Equals(args[i], name, StringComparison.OrdinalIgnoreCase))
            return args[i + 1];
    }
    return string.Empty;
}

var sourceDir = GetArg("--source");
var targetDir = GetArg("--target");
var exeName = GetArg("--exe");
var pidText = GetArg("--pid");

if (string.IsNullOrWhiteSpace(sourceDir) || string.IsNullOrWhiteSpace(targetDir))
    return;

if (!int.TryParse(pidText, out var pid))
    pid = -1;

if (pid > 0)
{
    try
    {
        var process = Process.GetProcessById(pid);
        process.WaitForExit(15000);
    }
    catch
    {
        // ignore
    }
}

try
{
    if (Directory.Exists(sourceDir))
    {
        foreach (var file in Directory.GetFiles(sourceDir, "*", SearchOption.AllDirectories))
        {
            var relative = Path.GetRelativePath(sourceDir, file);
            var destPath = Path.Combine(targetDir, relative);
            Directory.CreateDirectory(Path.GetDirectoryName(destPath)!);
            File.Copy(file, destPath, overwrite: true);
        }
    }
}
catch
{
    // ignore copy errors
}

if (!string.IsNullOrWhiteSpace(exeName))
{
    try
    {
        var exePath = Path.Combine(targetDir, exeName);
        if (File.Exists(exePath))
            Process.Start(new ProcessStartInfo { FileName = exePath, UseShellExecute = true });
    }
    catch
    {
        // ignore
    }
}
