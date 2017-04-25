//------------------------------------------------------------------------------
// <copyright file="JoinLines.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using EnvDTE;
using EnvDTE80;
using JoinLines.Helpers;
using System.Text.RegularExpressions;

namespace JoinLines
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class JoinLines
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("8ed65bca-91d1-41a5-8072-a267c17c86ee");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="JoinLines"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private JoinLines(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static JoinLines Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new JoinLines(package);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            var undoTransaction = new UndoTransactionHelper(ServiceProvider as JoinLinesPackage, "JoinLines");

            var activeTextDocument = (ServiceProvider as JoinLinesPackage).ActiveTextDocument;
            if (activeTextDocument != null)
            {
                var textSelection = activeTextDocument.Selection;

                if (textSelection != null)
                {
                    undoTransaction.Run(() => JoinLine(textSelection));
                }
            }
        }

        private void JoinLine(TextSelection textSelection)
        {
            // If the selection has no length, try to pick up the next line.
            if (textSelection.IsEmpty)
            {
                textSelection.LineDown(true);
                textSelection.EndOfLine(true);
            }

            const string fluentPattern = @"[ \t]*\r?\n[ \t]*\.";
            const string pattern = @"[ \t]*\r?\n[ \t]*";

            var selection = textSelection.Text;

            // do regex replace for fluent style
            selection = Regex.Replace(selection, fluentPattern, ".");

            // do regex replace for everything else
            selection = Regex.Replace(selection, pattern, " ");

            textSelection.Text = selection;
            
            // Move the cursor forward, clearing the selection.
            textSelection.CharRight();
        }
    }
}
