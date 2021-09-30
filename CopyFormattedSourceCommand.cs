using EnvDTE;
using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Task = System.Threading.Tasks.Task;

namespace CopyFormattedSource
{
   /// <summary>
   /// Command handler
   /// </summary>
   internal sealed class CopyFormattedSourceCommand
   {
      /// <summary>
      /// Default tab length to apply (for converting tabs to spaces).
      /// </summary>
      public const int DefaultTabLength = 4;

      /// <summary>
      /// Command ID.
      /// </summary>
      public const int CommandId = 0x0100;

      /// <summary>
      /// Command menu group (command set GUID).
      /// </summary>
      public static readonly Guid CommandSet = new Guid("ece55ea7-6c9c-4263-a78b-41d591745d30");

      /// <summary>
      /// VS Package that provides this command, not null.
      /// </summary>
      private readonly AsyncPackage package;

      /// <summary>
      /// Initializes a new instance of the <see cref="CopyFormattedSourceCommand"/> class.
      /// Adds our command handlers for menu (commands must exist in the command table file)
      /// </summary>
      /// <param name="package">Owner package, not null.</param>
      /// <param name="commandService">Command service to add command to, not null.</param>
      private CopyFormattedSourceCommand(AsyncPackage package, OleMenuCommandService commandService)
      {
         this.package = package ?? throw new ArgumentNullException(nameof(package));
         commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

         var menuCommandID = new CommandID(CommandSet, CommandId);
         var menuItem = new MenuCommand(this.Execute, menuCommandID);
         commandService.AddCommand(menuItem);
      }

      /// <summary>
      /// Gets the instance of the command.
      /// </summary>
      public static CopyFormattedSourceCommand Instance
      {
         get;
         private set;
      }

      /// <summary>
      /// Gets the service provider from the owner package.
      /// </summary>
      private IAsyncServiceProvider ServiceProvider
      {
         get => package;
      }

      /// <summary>
      /// Initializes the singleton instance of the command.
      /// </summary>
      /// <param name="package">Owner package, not null.</param>
      public static async Task InitializeAsync(AsyncPackage package)
      {
         await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

         OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
         Instance = new CopyFormattedSourceCommand(package, commandService);
      }

      private void Execute(object sender, EventArgs e)
         => _ = ExecuteAsync(sender, e);

      private async Task ExecuteAsync(object sender, EventArgs e)
      {
         await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

         /* get document object */
         var dte = await ServiceProvider.GetServiceAsync(typeof(DTE)) as DTE;
         var activeDocument = dte?.ActiveDocument;
         if (activeDocument == null)
            return;

         /* get active selection */
         var textDocument = activeDocument.Object() as TextDocument;
         var selection = activeDocument.Selection as TextSelection;
         if (selection.TopLine > selection.BottomLine)
            return;

         /* split lines */
         var text = textDocument.CreateEditPoint(textDocument.StartPoint).GetLines(selection.TopLine, selection.BottomLine + 1);
         var lines = Regex.Split(text, "\\r+\\n+").ToList();
         var lineOffset = 0;

         var firstNonEmptyLine = lines.FindIndex(s => !string.IsNullOrWhiteSpace(s));
         var lastNonEmptyLine = lines.FindLastIndex(s => !string.IsNullOrWhiteSpace(s));
         
         if (lines.Count <= 0
            || firstNonEmptyLine < 0
            || lastNonEmptyLine < 0)
            return;

         /* get the trimmed selection (includes the last line) */
         lines = lines.GetRange(firstNonEmptyLine, lastNonEmptyLine - firstNonEmptyLine + 1);

         /* convert tabs to spaces */
         lines = ApplyForEach(ref lines, (s) =>
         {
            var sb = new StringBuilder();
            var nonWhitespaceEncountered = false;
            foreach (var c in s.TrimEnd())
            {
               nonWhitespaceEncountered |= !char.IsWhiteSpace(c);
               if (!nonWhitespaceEncountered && c == '\t')
                  sb.Append(new string(' ', DefaultTabLength));
               else
                  sb.Append(c);
            }

            return sb.ToString();
         });

         /* determine the minimum whitespace to remove by */
         var minWhitespaceOnNonEmptyLine = int.MaxValue;
         foreach (var s in lines)
         {
            if (string.IsNullOrWhiteSpace(s))
               continue;

            minWhitespaceOnNonEmptyLine = Math.Min(minWhitespaceOnNonEmptyLine, CountWhitespaceChars(s));
         }

         /* remove N chars from the beginning of every line */
         lines = ApplyForEach(ref lines, (s) =>
         {
            if (string.IsNullOrWhiteSpace(s))
               return s;

            return (s.Length > minWhitespaceOnNonEmptyLine)
               ? s.Remove(0, minWhitespaceOnNonEmptyLine)
               : s;
         });

         /* wrap lines in language tags */
         var lineNo = selection.TopLine + lineOffset;
         Clipboard.SetText(string.Join(Environment.NewLine, lines.WrapSource(activeDocument, lineNo)));
      }

      private int CountWhitespaceChars(string input)
      {
         return string.IsNullOrEmpty(input)
            ? 0
            : input.Length - input.TrimStart().Length;
      }

      private List<string> ApplyForEach(ref List<string> input, Func<string, string> transform)
      {
         List<string> outputList = new List<string>();
         foreach (var s in input)
            outputList.Add(transform(s));
         return outputList;
      }
   }
}