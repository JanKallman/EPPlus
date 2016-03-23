using System;

namespace OfficeOpenXml
{
    /// <summary>
    /// Flag enum, specify all flags that you want to exclude from the copy.
    /// </summary>
    [Flags]    
    public enum ExcelRangeCopyOptionFlags : int
    {
        /// <summary>
        /// Exclude formulas from being copied
        /// </summary>
        ExcludeFormulas = 0x1,
    }
}
