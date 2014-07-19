using System;
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using System.Management.Automation.Language;
using System.Reflection;
using System.Text;
using System.Xml;

namespace PowerShell.MamlGenerator
{
    public class CmdletHelpGenerator
    {
        private static string _copyright;

        private static XmlTextWriter _writer;

        public static void GenerateHelp(string assemblyPath, string outputPath, string[] inputFiles)
        {
            GenerateHelp(Assembly.LoadFrom(assemblyPath), outputPath, inputFiles);
        }

        public static void GenerateHelp(Assembly asm, string outputPath, string[] inputFiles)
        {
            var attr = asm.GetCustomAttributes(typeof (AssemblyCopyrightAttribute), false);
            if (attr.Length == 1)
                _copyright = ((AssemblyCopyrightAttribute) attr[0]).Copyright;
            attr = asm.GetCustomAttributes(typeof (AssemblyDescriptionAttribute), false);
            if (attr.Length == 1)
                _copyright += " " + ((AssemblyDescriptionAttribute) attr[0]).Description;

            var comments = new Dictionary<string, CommentHelpInfo>();
            foreach (var input in inputFiles)
            {
                var fInfo = new FileInfo(input);
                if (!fInfo.Exists) continue;

                Token[] tokens;
                ParseError[] errors;
                var help = Parser.ParseFile(input, out tokens, out errors).GetHelpContent();
                if (help == null) continue;

                comments.Add(Path.GetFileNameWithoutExtension(fInfo.Name), help);
            }

            var sb = new StringBuilder();
            _writer = new XmlTextWriter(new StringWriter(sb)) {Formatting = Formatting.Indented};

            _writer.WriteStartDocument();
            _writer.WriteStartElement("helpItems");
            _writer.WriteAttributeString("xmlns", "http://msh");
            _writer.WriteAttributeString("schema", "maml");

            foreach (var type in asm.GetExportedTypes())
            {
                var ca = GetAttribute<CmdletAttribute>(type);
                if (ca == null) continue;

                var commandName = String.Format("{0}-{1}", ca.VerbName, ca.NounName);
                var commentHelpInfo = comments.ContainsKey(commandName) ? comments[commandName] : new CommentHelpInfo();

                _writer.WriteStartElement("command", "command",
                    "http://schemas.microsoft.com/maml/dev/command/2004/10");
                _writer.WriteAttributeString("xmlns", "maml", null, "http://schemas.microsoft.com/maml/2004/1");
                _writer.WriteAttributeString("xmlns", "dev", null, "http://schemas.microsoft.com/maml/dev/2004/10");
                _writer.WriteAttributeString("xmlns", "gl", null,
                    "http://schemas.sitecorepowershellextensions.com/maml/gl/2013/02");

                _writer.WriteStartElement("command", "details", null);

                _writer.WriteElementString("command", "name", null, commandName);

                //var group = GetAttribute<CmdletGroupAttribute>(type);
                //if (group != null && !string.IsNullOrEmpty(group.Group))
                //    _writer.WriteElementString("gl", "group", null, group.Group);
                //else
                _writer.WriteElementString("gl", "group", null, ca.NounName);

                WriteDescription(true, false, commentHelpInfo);

                WriteCopyright();

                _writer.WriteElementString("command", "verb", null, ca.VerbName);
                _writer.WriteElementString("command", "noun", null, ca.NounName);

                _writer.WriteElementString("dev", "version", null, asm.GetName().Version.ToString(3));

                _writer.WriteEndElement(); //command:details

                WriteDescription(false, true, commentHelpInfo);

                WriteSyntax(ca, type, commentHelpInfo);

                _writer.WriteStartElement("command", "parameters", null);

                foreach (var pi in type.GetProperties())
                {
                    var pas = GetAttribute<ParameterAttribute>(pi);
                    if (pas == null)
                        continue;

                    ParameterAttribute pa = null;
                    if (pas.Count == 1)
                        pa = pas[0];
                    else
                    {
                        // Determine the defualt property parameter set to use for details.
                        ParameterAttribute defaultPA = null;
                        foreach (var temp in pas)
                        {
                            var defaultSet = ca.DefaultParameterSetName;
                            if (string.IsNullOrEmpty(ca.DefaultParameterSetName))
                                defaultSet = string.Empty;

                            var set = temp.ParameterSetName;
                            if (string.IsNullOrEmpty(set) || set == DefaultParameterSetName)
                            {
                                set = string.Empty;
                                defaultPA = temp;
                            }
                            if (set.ToLower() != defaultSet.ToLower()) continue;

                            pa = temp;
                            defaultPA = temp;
                            break;
                        }
                        if (pa == null && defaultPA != null)
                            pa = defaultPA;
                        if (pa == null)
                            pa = pas[0];
                    }

                    _writer.WriteStartElement("command", "parameter", null);
                    _writer.WriteAttributeString("required", pa.Mandatory.ToString().ToLower());

                    var supportsWildcard = false; //GetAttribute<SupportsWildcardsAttribute>(pi) != null;
                    _writer.WriteAttributeString("globbing", supportsWildcard.ToString().ToLower());

                    if (!pa.ValueFromPipeline && !pa.ValueFromPipelineByPropertyName)
                        _writer.WriteAttributeString("pipelineInput", "false");
                    else if (pa.ValueFromPipeline && pa.ValueFromPipelineByPropertyName)
                        _writer.WriteAttributeString("pipelineInput", "true (ByValue, ByPropertyName)");
                    else if (!pa.ValueFromPipeline && pa.ValueFromPipelineByPropertyName)
                        _writer.WriteAttributeString("pipelineInput", "true (ByPropertyName)");
                    else if (pa.ValueFromPipeline && !pa.ValueFromPipelineByPropertyName)
                        _writer.WriteAttributeString("pipelineInput", "true (ByValue)");

                    if (pa.Position < 0)
                        _writer.WriteAttributeString("position", "named");
                    else
                        _writer.WriteAttributeString("position", (pa.Position + 1).ToString());

                    var variableLength = pi.PropertyType.IsArray;
                    _writer.WriteAttributeString("variableLength", variableLength.ToString().ToLower());

                    _writer.WriteElementString("maml", "name", null, pi.Name);

                    var helpMessage = pa.HelpMessage;
                    if (String.IsNullOrEmpty(helpMessage) && commentHelpInfo.Parameters != null)
                    {
                        if (commentHelpInfo.Parameters.ContainsKey(pi.Name.ToUpper()))
                        {
                            helpMessage = commentHelpInfo.Parameters[pi.Name.ToUpper()];
                        }
                    }
                    WriteDescription(helpMessage, false);

                    _writer.WriteStartElement("command", "parameterValue", null);
                    _writer.WriteAttributeString("required", pa.Mandatory.ToString().ToLower());
                    _writer.WriteAttributeString("variableLength", variableLength.ToString().ToLower());
                    _writer.WriteValue(pi.PropertyType.Name);
                    _writer.WriteEndElement(); //command:parameterValue

                    WriteDevType(pi.PropertyType.Name, null);

                    _writer.WriteEndElement(); //command:parameter
                }
                _writer.WriteEndElement(); //command:parameters

                WriteInputs(commentHelpInfo);

                WriteOutputs(commentHelpInfo);

                _writer.WriteElementString("command", "terminatingErrors", null, null);
                _writer.WriteElementString("command", "nonTerminatingErrors", null, null);

                WriteNotes(commentHelpInfo);

                WriteExamples(commentHelpInfo);
                WriteRelatedLinks(commentHelpInfo);

                _writer.WriteEndElement(); //command:command
            }

            _writer.WriteEndElement(); //helpItems
            _writer.WriteEndDocument();
            _writer.Flush();
            File.WriteAllText(Path.Combine(outputPath, string.Format("{0}.dll-help.xml", asm.GetName().Name)),
                sb.ToString());
        }

        private const string DefaultParameterSetName = "__AllParameterSets";

        private static void WriteSyntax(CmdletAttribute ca, Type type, CommentHelpInfo comment)
        {
            var parameterSets = new Dictionary<string, List<PropertyInfo>>();

            List<PropertyInfo> defaultSet = null;
            foreach (var pi in type.GetProperties())
            {
                var pas = GetAttribute<ParameterAttribute>(pi);
                if (pas == null)
                    continue;

                foreach (var temp in pas)
                {
                    var set = temp.ParameterSetName;
                    if (!parameterSets.ContainsKey(set))
                    {
                        var piList = new List<PropertyInfo>();
                        parameterSets.Add(set, piList);
                    }

                    parameterSets[set].Add(pi);
                }
            }

            if (parameterSets.Count > 1 && parameterSets.ContainsKey(DefaultParameterSetName))
            {
                defaultSet = parameterSets[DefaultParameterSetName];
                parameterSets.Remove(DefaultParameterSetName);
            }

            _writer.WriteStartElement("command", "syntax", null);
            foreach (var parameterSetName in parameterSets.Keys)
            {
                WriteSyntaxItem(ca, parameterSets, parameterSetName, defaultSet);
            }
            _writer.WriteEndElement(); //command:syntax
        }

        private static void WriteSyntaxItem(CmdletAttribute ca, Dictionary<string, List<PropertyInfo>> parameterSets,
            string parameterSetName, IEnumerable<PropertyInfo> defaultSet)
        {
            _writer.WriteStartElement("command", "syntaxItem", null);
            _writer.WriteElementString("maml", "name", null, string.Format("{0}-{1}", ca.VerbName, ca.NounName));
            foreach (var pi in parameterSets[parameterSetName])
            {
                var pa = GetParameterAttribute(pi, parameterSetName);
                if (pa == null)
                    continue;

                WriteParameter(pi, pa);
            }
            if (defaultSet != null)
            {
                foreach (var pi in defaultSet)
                {
                    var pas = GetAttribute<ParameterAttribute>(pi);
                    if (pas == null)
                        continue;
                    WriteParameter(pi, pas[0]);
                }
            }
            _writer.WriteEndElement(); //command:syntaxItem
        }

        private static ParameterAttribute GetParameterAttribute(PropertyInfo pi, string parameterSetName)
        {
            var pas = GetAttribute<ParameterAttribute>(pi);
            if (pas == null)
                return null;
            ParameterAttribute pa = null;
            foreach (var temp in pas)
            {
                if (temp.ParameterSetName.ToLower() == parameterSetName.ToLower())
                {
                    pa = temp;
                    break;
                }
            }
            return pa;
        }

        private static void WriteParameter(PropertyInfo pi, ParameterAttribute pa)
        {
            _writer.WriteStartElement("command", "parameter", null);
            _writer.WriteAttributeString("required", pa.Mandatory.ToString().ToLower());
            _writer.WriteAttributeString("parameterSetName", pa.ParameterSetName);
            if (pa.Position < 0)
                _writer.WriteAttributeString("position", "named");
            else
                _writer.WriteAttributeString("position", (pa.Position + 1).ToString());

            _writer.WriteElementString("maml", "name", null, pi.Name);
            _writer.WriteStartElement("command", "parameterValue", null);

            if (pi.DeclaringType == typeof (PSCmdlet))
                _writer.WriteAttributeString("required", "false");
            else
                _writer.WriteAttributeString("required", "true");

            if (pi.PropertyType.Name == "Nullable`1")
            {
                var coreType = pi.PropertyType.GetGenericArguments()[0];
                if (coreType.IsEnum)
                    _writer.WriteValue(string.Join(" | ", Enum.GetNames(coreType)));
                else
                    _writer.WriteValue(coreType.Name);
            }
            else
            {
                if (pi.PropertyType.IsEnum)
                    _writer.WriteValue(string.Join(" | ", Enum.GetNames(pi.PropertyType)));
                else
                    _writer.WriteValue(pi.PropertyType.Name);
            }

            _writer.WriteEndElement(); //command:parameterValue
            _writer.WriteEndElement(); //command:parameter
        }

        private static void WriteDevType(string name, string description)
        {
            _writer.WriteStartElement("dev", "type", null);
            _writer.WriteElementString("maml", "name", null, name);
            _writer.WriteElementString("maml", "uri", null, null);
            WriteDescription(description, false);
            _writer.WriteEndElement(); //dev:type
        }

        private static void WriteDescription(bool synopsis, bool addCopyright, CommentHelpInfo comment)
        {
            _writer.WriteStartElement("maml", "description", null);

            var desc = comment.Description;
            if (synopsis)
            {
                desc = comment.Synopsis;
            }

            WritePara(desc);
            if (addCopyright)
            {
                WritePara(null);
                WritePara(_copyright);
            }

            _writer.WriteEndElement(); //maml:description
        }

        private static void WriteDescription(string desc, bool addCopyright)
        {
            _writer.WriteStartElement("maml", "description", null);
            WritePara(desc);
            if (addCopyright)
            {
                WritePara(null);
                WritePara(_copyright);
            }
            _writer.WriteEndElement(); //maml:description
        }

        private static void WriteExamples(CommentHelpInfo comment)
        {
            if (comment.Examples == null || comment.Examples.Count == 0)
            {
                _writer.WriteElementString("command", "examples", null, null);
            }
            else
            {
                _writer.WriteStartElement("command", "examples", null);

                for (var i = 0; i < comment.Examples.Count; i++)
                {
                    var ex = comment.Examples[i];
                    _writer.WriteStartElement("command", "example", null);
                    if (comment.Examples.Count == 1)
                        _writer.WriteElementString("maml", "title", null, "------------------EXAMPLE------------------");
                    else
                        _writer.WriteElementString("maml", "title", null,
                            string.Format("------------------EXAMPLE {0}-----------------------", i + 1));

                    _writer.WriteElementString("dev", "code", null, ex);
                    _writer.WriteStartElement("dev", "remarks", null);
                    //WritePara(ex.Remarks);
                    _writer.WriteEndElement(); //dev:remarks
                    _writer.WriteEndElement(); //command:example
                }
                _writer.WriteEndElement(); //command:examples
            }
        }

        private static void WriteInputs(CommentHelpInfo comment)
        {
            if (comment.Inputs == null || comment.Inputs.Count == 0)
            {
                _writer.WriteElementString("command", "inputTypes", null);
            }
            else
            {
                _writer.WriteStartElement("command", "inputTypes", null);
                foreach (var input in comment.Inputs)
                {
                    _writer.WriteStartElement("command", "inputType", null);
                    //WriteDevType(null, null);
                    WriteDescription(input, false);
                    _writer.WriteEndElement(); //command:inputType
                }
                _writer.WriteEndElement(); //command:inputTypes
            }
        }

        private static void WriteOutputs(CommentHelpInfo comment)
        {
            if (comment.Outputs == null || comment.Outputs.Count == 0)
            {
                _writer.WriteElementString("command", "returnValues", null);
            }
            else
            {
                _writer.WriteStartElement("command", "returnValues", null);
                foreach (var input in comment.Inputs)
                {
                    _writer.WriteStartElement("command", "returnValue", null);
                    //WriteDevType(null, null);
                    WriteDescription(input, false);
                    _writer.WriteEndElement(); //command:returnValue
                }
                _writer.WriteEndElement(); //command:returnValues
            }
        }

        private static void WriteNotes(CommentHelpInfo comment)
        {
            _writer.WriteStartElement("maml", "alertSet", null);
            //_writer.WriteElementString("maml", "title", null, null);
            _writer.WriteStartElement("maml", "alert", null);
            WritePara(comment.Notes);
            /*
            WritePara(
                string.Format(
                    "For more information, type \"Get-Help {0}-{1} -detailed\". For technical information, type \"Get-Help {0}-{1} -full\".",
                    ca.VerbName, ca.NounName));
            */
            _writer.WriteEndElement(); //maml:alert
            _writer.WriteEndElement(); //maml:alertSet
        }

        private static void WriteRelatedLinks(CommentHelpInfo comment)
        {
            if (comment.Links == null || comment.Links.Count == 0)
            {
                _writer.WriteElementString("maml", "relatedLinks", null, null);
            }
            else
            {
                _writer.WriteStartElement("maml", "relatedLinks", null);

                foreach (var link in comment.Links)
                {
                    _writer.WriteStartElement("maml", "navigationLink", null);
                    _writer.WriteElementString("maml", "linkText", null, link);
                    _writer.WriteElementString("maml", "uri", null, null);
                    _writer.WriteEndElement(); //maml:navigationLink
                }
                _writer.WriteEndElement(); //maml:relatedLinks
            }
        }

        private static T GetAttribute<T>(Type type)
        {
            var attrs = type.GetCustomAttributes(typeof (T), true);
            if (attrs == null || attrs.Length == 0)
                return default(T);
            return (T) attrs[0];
        }

        private static List<T> GetAttribute<T>(PropertyInfo pi)
        {
            var attrs = pi.GetCustomAttributes(typeof (T), true);
            var attributes = new List<T>();
            if (attrs == null || attrs.Length == 0)
                return null;

            foreach (T t in attrs)
            {
                attributes.Add(t);
            }
            return attributes;
        }

        private static void WriteCopyright()
        {
            _writer.WriteStartElement("maml", "copyright", null);
            WritePara(_copyright);
            _writer.WriteEndElement(); //maml:copyright
        }

        private static void WritePara(string para)
        {
            if (string.IsNullOrEmpty(para))
            {
                _writer.WriteElementString("maml", "para", null, null);
                return;
            }
            var paragraphs = para.Split(new[] {"\r\n"}, StringSplitOptions.None);
            foreach (var p in paragraphs)
                _writer.WriteElementString("maml", "para", null, p);
        }
    }
}