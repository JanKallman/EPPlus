using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing
{
    internal class DependencyChain
    {
        internal List<FormulaCell> list = new List<FormulaCell>();
        internal Dictionary<ulong, int> index = new Dictionary<ulong, int>();
        internal List<int> CalcOrder = new List<int>();
        internal void Add(FormulaCell f)
        {
            list.Add(f);
            f.Index = list.Count - 1;
            index.Add(ExcelCellBase.GetCellID(f.SheetID, f.Row, f.Column), f.Index);
        }
    }
}