using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.VBA
{
    public enum eAttributeDataType
    {
        String=0,
        NonString=1
    }
    public class ExcelVbaModuleAttribute
    {
        internal ExcelVbaModuleAttribute()
        {

        }
        /// <summary>
        /// The name of the attribute
        /// </summary>
        public string Name { get; internal set; }
        /// <summary>
        /// The datatype.
        /// </summary>
        public eAttributeDataType DataType { get; internal set; }
        /// <summary>
        /// The value of the attribute without any double quotes.
        /// </summary>
        public string Value { get; set; }
        public override string ToString()
        {
            return Name;
        }
    }
}
