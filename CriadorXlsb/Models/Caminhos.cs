using CriadorXLSB.Helpers;

namespace CriadorXLSB.Models;
public class Caminhos(string inputPath, string outputPath)
{
    public string? PythonPath { get; private set; } = PythonHelper.GetPythonPath();
    public string ScriptPath { get; private set; } = PythonHelper.GetScriptPath("converterXlsxEmXlsb.py");
    public string InputPath { get; private set; } = inputPath;
    public string OutputPath { get; private set; } = outputPath;
}
