using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using OfficeOpenXml.Utils;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using System.Text.RegularExpressions;
using System.Collections;

namespace OfficeOpenXml.DataValidation.Formulas
{
    internal class ExcelDataValidationFormulaList : ExcelDataValidationFormula, IExcelDataValidationFormulaList
    {
        #region class DataValidationList
        private class DataValidationList : IList<string>, ICollection
        {
            private IList<string> _items = new List<string>();
            private EventHandler<EventArgs> _listChanged;

            public event EventHandler<EventArgs> ListChanged
            {
                add { _listChanged += value; }
                remove { _listChanged -= value; }
            }

            private void OnListChanged()
            {
                if (_listChanged != null)
                {
                    _listChanged(this, EventArgs.Empty);
                }
            }

            #region IList members
            int IList<string>.IndexOf(string item)
            {
                return _items.IndexOf(item);
            }

            void IList<string>.Insert(int index, string item)
            {
                _items.Insert(index, item);
                OnListChanged();
            }

            void IList<string>.RemoveAt(int index)
            {
                _items.RemoveAt(index);
                OnListChanged();
            }

            string IList<string>.this[int index]
            {
                get
                {
                    return _items[index];
                }
                set
                {
                    _items[index] = value;
                    OnListChanged();
                }
            }

            void ICollection<string>.Add(string item)
            {
                _items.Add(item);
                OnListChanged();
            }

            void ICollection<string>.Clear()
            {
                _items.Clear();
                OnListChanged();
            }

            bool ICollection<string>.Contains(string item)
            {
                return _items.Contains(item);
            }

            void ICollection<string>.CopyTo(string[] array, int arrayIndex)
            {
                _items.CopyTo(array, arrayIndex);
            }

            int ICollection<string>.Count
            {
                get { return _items.Count; }
            }

            bool ICollection<string>.IsReadOnly
            {
                get { return false; }
            }

            bool ICollection<string>.Remove(string item)
            {
                var retVal = _items.Remove(item);
                OnListChanged();
                return retVal;
            }

            IEnumerator<string> IEnumerable<string>.GetEnumerator()
            {
                return _items.GetEnumerator();
            }

            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
            {
                return _items.GetEnumerator();
            }
            #endregion

            public void CopyTo(Array array, int index)
            {
                _items.CopyTo((string[])array, index);
            }

            int ICollection.Count
            {
                get { return _items.Count; }
            }

            public bool IsSynchronized
            {
                get { return ((ICollection)_items).IsSynchronized; }
            }

            public object SyncRoot
            {
                get { return ((ICollection)_items).SyncRoot; }
            }
        }
        #endregion

        public ExcelDataValidationFormulaList(XmlNamespaceManager namespaceManager, XmlNode itemNode, string formulaPath)
            : base(namespaceManager, itemNode, formulaPath)
        {
            Require.Argument(formulaPath).IsNotNullOrEmpty("formulaPath");
            _formulaPath = formulaPath;
            var values = new DataValidationList();
            values.ListChanged += new EventHandler<EventArgs>(values_ListChanged);
            Values = values;
            SetInitialValues();
        }

        private string _formulaPath;

        private void SetInitialValues()
        {
            var @value = GetXmlNodeString(_formulaPath);
            if (!string.IsNullOrEmpty(@value))
            {
                if (Regex.IsMatch(@value, "^\"[^\\,]+,[^\\,]+\"$"))
                {
                    @value = @value.TrimStart('"').TrimEnd('"');
                    var items = @value.Split(',');
                    foreach (var item in items)
                    {
                        Values.Add(item);
                    }
                }
                else
                {
                    ExcelFormula = @value;
                }
            }
        }

        void values_ListChanged(object sender, EventArgs e)
        {
            if (Values.Count > 0)
            {
                State = FormulaState.Value;
            }
            SetXmlNodeString(_formulaPath, GetValueAsString());
        }

        public IList<string> Values
        {
            get;
            private set;
        }

        protected override string  GetValueAsString()
        {
            var sb = new StringBuilder();
            foreach (var val in Values)
            {
                if (sb.Length == 0)
                {
                    sb.Append("\"");
                    sb.Append(val);
                }
                else
                {
                    sb.AppendFormat(", {0}", val);
                }
            }
            sb.Append("\"");
            return sb.ToString();
        }

        internal override void ResetValue()
        {
            Values.Clear();
        }
    }
}
