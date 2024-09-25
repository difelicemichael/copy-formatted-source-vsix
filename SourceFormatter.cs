using EnvDTE;
using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CopyMarkdownSource
{
    /// <summary>
    /// Example source formats.
    /// </summary>
    internal enum SourceFileType
    {
        Unknown,
        CSharp,
        CPlusPlus,
        Python,
        XML,
        Javascript,
        Html,
        Css,
        BatchScript,
        ShellScript,
        GLSourceLang,
        PlainText
    }

    internal static class SourceFormatter
    {
        /// <summary>
        /// Const to avoid instantiation.
        /// </summary>
        private const string SourceFormatStr = "```";

        /// <summary>
        /// Wraps a source code selection in backticks, and includes a line denoting the
        /// language of the source code, retrieved from the file path of the provided
        /// document.
        /// </summary>
        internal static IEnumerable<string> WrapSource(this IEnumerable<string> sourceSelection, Document document, int line)
        {
            /* header first */
            yield return GetLanguageHeader(document);

            /* then path comment */
            yield return GetPathComment(document, line);

            /* then, every line in the selection */
            foreach (var s in sourceSelection)
                yield return s;

            /* finally, the footer */
            yield return SourceFormatStr;
        }

        private static SourceFileType IdentifyByExtension(string path)
        {
            switch (Path.GetExtension(path).ToLower())
            {
                case ".cs":
                case ".cshtml":
                    return SourceFileType.CSharp;

                case ".c":
                case ".h":
                case ".cpp":
                case ".hpp":
                case ".ih":
                    return SourceFileType.CPlusPlus;

                case ".glsl":
                    return SourceFileType.GLSourceLang;

                case ".html":
                    return SourceFileType.Html;

                case ".css":
                    return SourceFileType.Css;

                case ".sh":
                    return SourceFileType.ShellScript;

                case ".bat":
                case ".cmd":
                    return SourceFileType.BatchScript;

                case ".py":
                    return SourceFileType.Python;

                case ".res":
                case ".csv":
                    return SourceFileType.PlainText;

                case ".js":
                    return SourceFileType.Javascript;

                case ".xml":
                case ".resx":
                case ".config":
                case ".wxs":
                    return SourceFileType.XML;

                default:
                    return SourceFileType.Unknown;
            }
        }

        private static string GetPathComment(Document activeDocument, int line)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (string.IsNullOrEmpty(activeDocument?.FullName))
                return string.Empty;

            var pathAndLine = $"{activeDocument.FullName}({line})";
            switch (IdentifyByExtension(activeDocument?.FullName))
            {
                case SourceFileType.CSharp:
                case SourceFileType.CPlusPlus:
                case SourceFileType.Javascript:
                case SourceFileType.Css:
                case SourceFileType.GLSourceLang:
                    return $"/* {pathAndLine} */";

                case SourceFileType.Python:
                case SourceFileType.ShellScript:
                    return $"# {pathAndLine}";

                case SourceFileType.XML:
                case SourceFileType.Html:
                    return $"<!-- {pathAndLine} -->";

                case SourceFileType.BatchScript:
                    return $"REM {pathAndLine}";

                case SourceFileType.PlainText:
                case SourceFileType.Unknown:
                default:
                    return $"// {pathAndLine}";
            }
        }

        private static string GetLanguageHeader(Document activeDocument)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            switch (IdentifyByExtension(activeDocument?.FullName))
            {
                case SourceFileType.CSharp:
                    return $"{SourceFormatStr}cs";

                case SourceFileType.CPlusPlus:
                    return $"{SourceFormatStr}c++";

                case SourceFileType.Javascript:
                    return $"{SourceFormatStr}javascript";

                case SourceFileType.Css:
                    return $"{SourceFormatStr}css";

                case SourceFileType.GLSourceLang:
                    return $"{SourceFormatStr}glsl";

                case SourceFileType.Python:
                    return $"{SourceFormatStr}python";

                case SourceFileType.ShellScript:
                    return $"{SourceFormatStr}sh";

                case SourceFileType.XML:
                    return $"{SourceFormatStr}xml";

                case SourceFileType.Html:
                    return $"{SourceFormatStr}html";

                case SourceFileType.BatchScript:
                    return $"{SourceFormatStr}bat";

                case SourceFileType.PlainText:
                case SourceFileType.Unknown:
                default:
                    return $"{SourceFormatStr}Plain Text";
            }
        }
    }
}
