using System.Diagnostics;

namespace CriadorXLSB.Helpers;
public class PythonHelper
{
    public static string? GetPythonPath()
    {
        try
        {
            ProcessStartInfo startInfo = new()
            {
                FileName = "cmd.exe",
                Arguments = "/c where python",
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using Process process = new() { StartInfo = startInfo };

            process.Start();
            string? output = process.StandardOutput.ReadLine();
            process.WaitForExit();
            return output?.Trim();
        }
        catch
        {
            throw new Exception("Python não encontrado! Verifique se está instalado.");
        }
    }

    public static string GetScriptPath(string arquivo)
    {
        string scriptPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, arquivo);

        if(File.Exists(scriptPath))
            return scriptPath;

        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            @".nuget\packages\criadorxlsb\1.0.6\contentFiles\any\any", arquivo);
    }
}
