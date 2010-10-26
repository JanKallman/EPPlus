using System;
using System.Collections.Generic;
using System.Text;
using System.Globalization;

namespace OfficeOpenXml
{
    public class ExcelTextFormat
    {
        public enum eDataTypes
        {
            Unknown,
            String,
            Number,
            DateTime,
        }
        public ExcelTextFormat()
        {
            Delimiter = ',';
            TextQualifier = '\0';
            EOL = "\r\n";
            Culture = CultureInfo.InvariantCulture;
            DataTypes=null;
            SkipLinesBeginning = 0;
            SkipLinesEnd = 0;
            Encoding=Encoding.ASCII;
        }
        public char Delimiter { get; set; }
        public char TextQualifier {get; set; }
        public string EOL { get; set; }
        public eDataTypes[] DataTypes { get; set; }
        public CultureInfo Culture {get; set; }
        public int SkipLinesBeginning { get; set; }
        public int SkipLinesEnd { get; set; }
        public Encoding Encoding { get; set; }
    }
}
