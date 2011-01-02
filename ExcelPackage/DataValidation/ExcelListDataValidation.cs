using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.DataValidation
{
    public class ExcelListDataValidation : ExcelDataValidation
    {
        #region class DataValidationList
        private class DataValidationList : IList<string>
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
        }
        #endregion

        public ExcelListDataValidation(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType)
            : base(worksheet, address, validationType)
        {
            var values = new DataValidationList();
            values.ListChanged += new EventHandler<EventArgs>(values_ListChanged);
            Values = values;
        }

        void values_ListChanged(object sender, EventArgs e)
        {
            SetXmlNodeString(_formula1Path, ValuesToString());
        }

        public IList<string> Values
        {
            get;
            private set;
        }

        private string ValuesToString()
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
    }
}
