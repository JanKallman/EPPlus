using System;

namespace OfficeOpenXml.FormulaParsing.Logging
{
    /// <summary>
    /// Used for logging during FormulaParsing
    /// </summary>
    public interface IFormulaParserLogger : IDisposable
    {
        /// <summary>
        /// Called each time an exception occurs during formula parsing.
        /// </summary>
        /// <param name="context"></param>
        /// <param name="ex"></param>
        void Log(ParsingContext context, Exception ex);
        /// <summary>
        /// Called each time information should be logged during formula parsing.
        /// </summary>
        /// <param name="context"></param>
        /// <param name="message"></param>
        void Log(ParsingContext context, string message);
        /// <summary>
        /// Called to log a message outside the parsing context.
        /// </summary>
        /// <param name="message"></param>
        void Log(string message);
        /// <summary>
        /// Called each time a cell within the calc chain is accessed during formula parsing.
        /// </summary>
        void LogCellCounted();

        /// <summary>
        /// Called each time a function is called during formula parsing.
        /// </summary>
        /// <param name="func"></param>
        void LogFunction(string func);
        /// <summary>
        /// Some functions measure performance, if so this function will be called.
        /// </summary>
        /// <param name="func"></param>
        /// <param name="milliseconds"></param>
        void LogFunction(string func, long milliseconds);
    }
}
