using System;

namespace OfficeOpenXml
{
    /// <summary>
    /// Flag enum, specify all flags that you want to exclude from the copy.
    /// </summary>
    [Flags]    
    public enum ExcelRangeCopyExcludeFlags : int
    {
        /// <summary>
        /// Exclude forumlas from being copied
        /// </summary>
        Forumla = 0x1,
    }
}
