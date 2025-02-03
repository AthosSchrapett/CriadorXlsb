using CriadorXLSB.Helpers;
using CriadorXLSB.Models;
using System.Diagnostics;

namespace CriadorXLSB;
public class GeradorXlsb
{
    public Caminhos Caminhos { get; private set; }

    public GeradorXlsb(Caminhos caminhos)
    {
        Caminhos = caminhos;
    }

    public void GerarXlsb()
    {
        ProcessStartInfo startInfo = new()
        {
            FileName = Caminhos.PythonPath,
            Arguments = $"\"{Caminhos.ScriptPath}\" \"{Caminhos.InputPath}\" \"{Caminhos.OutputPath}\"",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        using Process process = new() { StartInfo = startInfo };
        process.Start();

        string output = process.StandardOutput.ReadToEnd();
        string error = process.StandardError.ReadToEnd();

        process.WaitForExit();

        Console.WriteLine("Saída do script Python:");
        Console.WriteLine(output);

        if (!string.IsNullOrEmpty(error))
        {
            Console.WriteLine("Erros do script Python:");
            Console.WriteLine(error);
        }
    }
}
