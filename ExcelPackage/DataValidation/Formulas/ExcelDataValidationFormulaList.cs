/* 
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * EPPlus provides server-side generation of Excel 2007 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 * 
 * All rights reserved.
 * 
 * EPPlus is an Open Source project provided under the 
 * GNU General Public License (GPL) as published by the 
 * Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
 * 
 * The GNU General Public License can be viewed at http://www.opensource.org/licenses/gpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 * 
 * The code for this project may be used and redistributed by any means PROVIDING it is 
 * not sold for profit without the author's written consent, and providing that this notice 
 * and the author's name and all copyright notices remain intact.
 * 
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 *  Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Mats Alm   		                Added       		        2011-01-08
 *******************************************************************************/
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
                if (@value.StartsWith("\"") && @value.EndsWith("\""))
                {
                    @value = @value.TrimStart('"').TrimEnd('"');
                    var items = @value.Split(new char[]{','}, StringSplitOptions.RemoveEmptyEntries);
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
