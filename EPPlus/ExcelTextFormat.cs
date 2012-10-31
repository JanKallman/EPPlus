/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Jan Källman		    Initial Release		        2011-01-01
 * Jan Källman		    License changed GPL-->LGPL 2011-12-27
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Globalization;

namespace OfficeOpenXml
{
    /// <summary>
    /// Discribes a column when reading a text using the ExcelRangeBase.LoadFromText method
    /// </summary>
    public enum eDataTypes
    {
        /// <summary>
        /// Let the the import decide.
        /// </summary>
        Unknown,
        /// <summary>
        /// Always a string.
        /// </summary>
        String,
        /// <summary>
        /// Try to convert it to a number. If it fails then add it as a string.
        /// </summary>
        Number,
        /// <summary>
        /// Try to convert it to a date. If it fails then add it as a string.
        /// </summary>
        DateTime,
        /// <summary>
        /// Try to convert it to a number and divide with 100. 
        /// Removes any tailing percent sign (%). If it fails then add it as a string.
        /// </summary>
        Percent
    }
    /// <summary>
    /// Describes how to split a CSV text. Used by the ExcelRange.LoadFromText method
    /// </summary>
    public class ExcelTextFormat
    {
        /// <summary>
        /// Describes how to split a CSV text
        /// 
        /// Default values
        /// <list>
        /// <listheader><term>Property</term><description>Value</description></listheader>
        /// <item><term>Delimiter</term><description>,</description></item>
        /// <item><term>TextQualifier</term><description>None (\0)</description></item>
        /// <item><term>EOL</term><description>CRLF</description></item>
        /// <item><term>Culture</term><description>CultureInfo.InvariantCulture</description></item>
        /// <item><term>DataTypes</term><description>End of line default CRLF</description></item>
        /// <item><term>SkipLinesBeginning</term><description>0</description></item>
        /// <item><term>SkipLinesEnd</term><description>0</description></item>
        /// <item><term>Encoding</term><description>Encoding.ASCII</description></item>
        /// </list>
        /// </summary>
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
        /// <summary>
        /// Delimiter character
        /// </summary>
        public char Delimiter { get; set; }
        /// <summary>
        /// Text qualifier character 
        /// </summary>
        public char TextQualifier {get; set; }
        /// <summary>
        /// End of line characters. Default CRLF
        /// </summary>
        public string EOL { get; set; }
        /// <summary>
        /// Datatypes list for each column (if column is not present Unknown is assumed)
        /// </summary>
        public eDataTypes[] DataTypes { get; set; }
        /// <summary>
        /// Culture used when parsing. Default CultureInfo.InvariantCulture
        /// </summary>
        public CultureInfo Culture {get; set; }
        /// <summary>
        /// Number of lines skiped in the begining of the file. Default 0.
        /// </summary>
        public int SkipLinesBeginning { get; set; }
        /// <summary>
        /// Number of lines skiped at the end of the file. Default 0.
        /// </summary>
        public int SkipLinesEnd { get; set; }
        /// <summary>
        /// Only used when reading files from disk using a FileInfo object. Default AscII
        /// </summary>
        public Encoding Encoding { get; set; }
    }
}
