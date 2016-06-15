using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Logging
{
    /// <summary>
    /// Create loggers that can be used for logging the formula parser.
    /// </summary>
    public static class LoggerFactory
    {
        /// <summary>
        /// Creates a logger that logs to a simple textfile.
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public static IFormulaParserLogger CreateTextFileLogger(FileInfo file)
        {
            return new TextFileLogger(file);
        }
    }
}
