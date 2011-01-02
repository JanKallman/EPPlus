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
 * Mats Alm   		                Added       		        2011-01-01
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using OfficeOpenXml.Utils;
using System.Xml;

namespace OfficeOpenXml.DataValidation
{
    /// <summary>
    /// Collection of <see cref="ExcelDataValidation"/>
    /// </summary>
    public class ExcelDataValidationCollection : XmlHelper, IEnumerable<ExcelDataValidation>
    {
        private List<ExcelDataValidation> _validations = new List<ExcelDataValidation>();
        private ExcelWorksheet _worksheet = null;

        private const string DataValidationPath = "//d:dataValidations";
        private readonly string DataValidationItemsPath = string.Format("{0}/d:dataValidation", DataValidationPath);

        internal ExcelDataValidationCollection(ExcelWorksheet worksheet)
            : base(worksheet.NameSpaceManager, worksheet.WorksheetXml.DocumentElement)
        {
            Require.Argument(worksheet).IsNotNull("worksheet");
            _worksheet = worksheet;

            // check existing nodes and load them
            var dataValidationNodes = worksheet.WorksheetXml.SelectNodes(DataValidationItemsPath, worksheet.NameSpaceManager);
            if (dataValidationNodes != null && dataValidationNodes.Count > 0)
            {
                foreach (XmlNode node in dataValidationNodes)
                {
                    var addr = node.Attributes["sqref"].Value;
                    var dataValidationType = (eDataValidationType)Enum.Parse(typeof(eDataValidationType), node.Attributes["type"].Value);
                    var type = ExcelDataValidationType.GetByValidationType(dataValidationType);
                    _validations.Add(new ExcelDataValidation(worksheet, addr, type, node));
                }
            }

        }

        private void EnsureRootElementExists()
        {
            var node = _worksheet.WorksheetXml.SelectSingleNode(DataValidationPath, _worksheet.NameSpaceManager);
            if (node == null)
            {
                //node = _worksheet.WorksheetXml.CreateElement(DataValidationPath.TrimStart('/'));
                CreateNode(DataValidationPath.TrimStart('/'));
                //_worksheet.WorksheetXml.DocumentElement.AppendChild(node);
            }
        }

        private XmlNode GetRootNode()
        {
            EnsureRootElementExists();
            return _worksheet.WorksheetXml.SelectSingleNode(DataValidationPath, _worksheet.NameSpaceManager);
        }

        ///// <summary>
        ///// Creates an <see cref="ExcelDataValidation"/> instance
        ///// </summary>
        ///// <param name="address"></param>
        ///// <param name="validationType">Validation type</param>
        ///// <returns>An instance of <see cref="ExcelDataValidation"/></returns>
        ///// <exception cref="ArgumentNullException">If address is null</exception>
        //private ExcelDataValidation Create(string address, ExcelDataValidationType validationType)
        //{
        //    return new ExcelDataValidation(_worksheet.NameSpaceManager, GetRootNode(), address, validationType);
        //}

        /// <summary>
        /// Adds an <see cref="ExcelIntDataValidation"/> to the worksheet. Whole means that the only accepted values
        /// are integer values.
        /// </summary>
        /// <param name="address">the range/address to validate</param>
        public ExcelIntDataValidation AddWholeValidation(string address)
        {
            EnsureRootElementExists(); 
            var item = new ExcelIntDataValidation(_worksheet, address, ExcelDataValidationType.Whole);
            _validations.Add(item);
            return item;
        }

        /// <summary>
        /// Addes an <see cref="ExcelDataValidation"/> to the worksheet. The only accepted values are
        /// decimal values.
        /// </summary>
        /// <param name="address">The range/address to validate</param>
        /// <returns></returns>
        public ExcelDecimalDataValidation AddDecimalValidation(string address)
        {
            EnsureRootElementExists();
            var item = new ExcelDecimalDataValidation(_worksheet, address, ExcelDataValidationType.Decimal);
            _validations.Add(item);
            return item;
        }

        /// <summary>
        /// Adds an <see cref="ExcelListDataValidation"/> to the worksheet. The accepted values are defined
        /// in a list.
        /// </summary>
        /// <param name="address">The range/address to validate</param>
        /// <returns></returns>
        public ExcelListDataValidation AddListValidation(string address)
        {
            EnsureRootElementExists();
            var item = new ExcelListDataValidation(_worksheet, address, ExcelDataValidationType.List);
            _validations.Add(item);
            return item;
        }

        /// <summary>
        /// Adds an <see cref="ExcelIntDataValidation"/> regarding text length to the worksheet.
        /// </summary>
        /// <param name="address">The range/address to validate</param>
        /// <returns></returns>
        public ExcelIntDataValidation AddTextLengthValidation(string address)
        {
            EnsureRootElementExists();
            var item = new ExcelIntDataValidation(_worksheet, address, ExcelDataValidationType.TextLength);
            _validations.Add(item);
            return item;
        }

        /// <summary>
        /// Removes an <see cref="ExcelDataValidation"/> from the collection.
        /// </summary>
        /// <param name="item">The item to remove</param>
        /// <returns>True if remove succeeds, otherwise false</returns>
        /// <exception cref="ArgumentNullException">if <paramref name="item"/> is null</exception>
        public bool Remove(ExcelDataValidation item)
        {
            Require.Argument(item).IsNotNull("item");
            return _validations.Remove(item);
        }

        public int Count
        {
            get { return _validations.Count; }
        }

        public ExcelDataValidation this[int index]
        {
            get { return _validations[index]; }
            set { _validations[index] = value; }
        }

        public IEnumerable<ExcelDataValidation> FindAll(Predicate<ExcelDataValidation> match)
        {
            return _validations.FindAll(match);
        }

        public ExcelDataValidation Find(Predicate<ExcelDataValidation> match)
        {
            return _validations.Find(match);
        }

        public void Clear()
        {
            DeleteAllNode(DataValidationItemsPath);
            _validations.Clear();
        }

        /// <summary>
        /// Removes the validations that matches the predicate
        /// </summary>
        /// <param name="match"></param>
        public void RemoveAll(Predicate<ExcelDataValidation> match)
        {
            _validations.RemoveAll(match);
        }

        IEnumerator<ExcelDataValidation> IEnumerable<ExcelDataValidation>.GetEnumerator()
        {
            return _validations.GetEnumerator();
        }

        IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _validations.GetEnumerator();
        }
    }
}
