using System.IO;
using System.Reflection;
using PowerShell.MamlGenerator;

namespace Runner
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var assembly =
                Assembly.LoadFile(Path.GetFullPath(@"..\..\..\Documentation\Microsoft.Samples.PowerShell.dll"));
            var outputPath = Path.GetFullPath(@"..\..\..\Documentation");
            var inputFiles = Directory.GetFiles(@"..\..\..\Documentation\", "*.ps1");
            CmdletHelpGenerator.GenerateHelp(assembly, outputPath, inputFiles);
        }
    }
}