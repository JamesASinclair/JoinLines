//using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JoinLines.Helpers
{
    internal static class OutputWindowHelper
    {
        #region Fields

        private static IVsOutputWindowPane _outputWindowPane;

        #endregion Fields

        #region Properties

        private static IVsOutputWindowPane OutputWindowPane => _outputWindowPane ?? (_outputWindowPane = GetOutputWindowPane());

        #endregion Properties

        #region Methods

        /// <summary>
        /// Writes the specified diagnostic line to the output pane.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="ex">An optional exception that was handled.</param>
        internal static void DiagnosticWriteLine(string message, Exception ex = null)
        {
            if (ex != null)
            {
                message += $": {ex}";
            }

            WriteLine("Diagnostic", message);
        }

        /// <summary>
        /// Writes the specified exception line to the output pane.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="ex">The exception that was handled.</param>
        internal static void ExceptionWriteLine(string message, Exception ex)
        {
            var exceptionMessage = $"{message}: {ex}";

            WriteLine("Handled Exception", exceptionMessage);
        }

        /// <summary>
        /// Writes the specified warning line to the output pane.
        /// </summary>
        /// <param name="message">The message.</param>
        internal static void WarningWriteLine(string message)
        {
            WriteLine("Warning", message);
        }

        /// <summary>
        /// Attempts to create and retrieve the output window pane.
        /// </summary>
        /// <returns>The output window pane, otherwise null.</returns>
        private static IVsOutputWindowPane GetOutputWindowPane()
        {
            var outputWindow = Package.GetGlobalService(typeof(SVsOutputWindow)) as IVsOutputWindow;
            if (outputWindow == null) return null;
            
            // guid is for the output window
            Guid outputPaneGuid = new Guid("4e7ba904-9311-4dd0-abe5-a61c1739780f");
            IVsOutputWindowPane windowPane;

            outputWindow.CreatePane(ref outputPaneGuid, "JoinLines", 1, 1);
            outputWindow.GetPane(ref outputPaneGuid, out windowPane);

            return windowPane;
        }
        
        /// <summary>
        /// Writes the specified line to the output pane.
        /// </summary>
        /// <param name="category">The category.</param>
        /// <param name="message">The message.</param>
        private static void WriteLine(string category, string message)
        {
            var outputWindowPane = OutputWindowPane;
            if (outputWindowPane != null)
            {
                string outputMessage = $"[JoinLines {category} {DateTime.Now.ToString("hh:mm:ss tt")}] {message}{Environment.NewLine}";

                outputWindowPane.OutputString(outputMessage);
            }
        }

        #endregion Methods
    }
}