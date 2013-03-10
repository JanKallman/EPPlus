using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
    public class RangeAddress
    {
        public RangeAddress()
        {
            Address = string.Empty;
        }

        internal string Address { get; set; }

        public string Worksheet { get; internal set; }

        public int FromCol { get; internal set; }

        public int ToCol { get; internal set; }

        public int FromRow { get; internal set; }

        public int ToRow { get; internal set; }

        public override string ToString()
        {
            return Address;
        }

        private static RangeAddress _empty = new RangeAddress();
        public static RangeAddress Empty
        {
            get { return _empty; }
        }

        /// <summary>
        /// Returns true if this range collides (full or partly) with the supplied range
        /// </summary>
        /// <param name="other">The range to check</param>
        /// <returns></returns>
        public bool CollidesWith(RangeAddress other)
        {
            if (other.Worksheet != Worksheet)
            {
                return false;
            }
            if (other.FromRow > ToRow || other.FromCol > ToCol
                ||
                FromRow > other.ToRow || FromCol > other.ToCol)
            {
                return false;
            }
            return true;
        }
    }
}
