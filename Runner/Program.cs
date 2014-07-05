using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using PowerShell.MamlGenerator;

namespace Runner
{
    class Program
    {
        static void Main(string[] args)
        {
            var assembly = Assembly.LoadFile(Path.GetFullPath(@"..\..\..\Documentation\Microsoft.Samples.PowerShell.dll"));
            var outputPath = Path.GetFullPath(@"..\..\..\Documentation");
            var inputFiles = Directory.GetFiles(@"..\..\..\Documentation\", "*.ps1");
            CmdletHelpGenerator.GenerateHelp(assembly,outputPath,inputFiles);
        }
    }
}
