using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JoinLines.Helpers
{
    public class UndoTransactionHelper
    {
        #region Fields

        private readonly JoinLinesPackage _package;
        private readonly string _transactionName;

        #endregion Fields

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="UndoTransactionHelper" /> class.
        /// </summary>
        /// <param name="package">The hosting package.</param>
        /// <param name="transactionName">The name of the transaction.</param>
        public UndoTransactionHelper(JoinLinesPackage package, string transactionName)
        {
            _package = package;
            _transactionName = transactionName;
        }

        #endregion Constructors

        #region Methods

        /// <summary>
        /// Runs the specified try action within a try block, and conditionally the catch action
        /// within a catch block all conditionally within the context of an undo transaction.
        /// </summary>
        /// <param name="tryAction">The action to be performed within a try block.</param>
        /// <param name="catchAction">The action to be performed wihin a catch block.</param>
        public void Run(Action tryAction, Action<Exception> catchAction = null)
        {
            bool shouldCloseUndoContext = false;

            // Start an undo transaction (unless inside one already or within an auto save context).
            if (!_package.IDE.UndoContext.IsOpen)
            {
                _package.IDE.UndoContext.Open(_transactionName);
                shouldCloseUndoContext = true;
            }

            try
            {
                tryAction();
            }
            catch (Exception ex)
            {
                var message = $"{_transactionName} was stopped";
                OutputWindowHelper.ExceptionWriteLine(message, ex);
                _package.IDE.StatusBar.Text = $"{message}.  See output window for more details.";

                catchAction?.Invoke(ex);

                if (shouldCloseUndoContext)
                {
                    _package.IDE.UndoContext.SetAborted();
                    shouldCloseUndoContext = false;
                }
            }
            finally
            {
                // Always close the undo transaction to prevent ongoing interference with the IDE.
                if (shouldCloseUndoContext)
                {
                    _package.IDE.UndoContext.Close();
                }
            }
        }

        #endregion Methods
    }
}
