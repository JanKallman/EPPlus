using System.IO;
using OfficeOpenXml;

namespace EPPlusTest
{
    internal static class WorksheetExtensionMethods
    {
        public static ExcelPackage AsExcelPackage(this string fileName) 
            => new ExcelPackage(fileName.AsWorkSheetFileInfo());

        public static string AsWorkSheetPath(this string fileName) 
            => Path.Combine(Scaffolding.WorksheetPath, fileName);

        public static FileInfo AsWorkSheetFileInfo(this string fileName) 
            => new FileInfo(fileName.AsWorkSheetPath());

        public static bool WorkSheetExists(this string fileName)
            => fileName.AsWorkSheetFileInfo().Exists;
    }
}