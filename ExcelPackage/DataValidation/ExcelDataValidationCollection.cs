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
    /// <para>
    /// Collection of <see cref="ExcelDataValidation"/>. This class is providing the API for EPPlus data validation.
    /// </para>
    /// <para>
    /// The public methods of this class (Add[...]Validation) will create a datavalidation entry in the worksheet. When this
    /// validation has been created changes to the properties will affect the workbook immediately.
    /// </para>
    /// <para>
    /// Each type of validation has either a formula or a typed value.
    /// </para>
    /// <code>
    /// // Add a date time validation
    /// var validation = worksheet.DataValidation.AddDateTimeValidation("A1");
    /// // set validation properties
    /// validation.ShowErrorMessage = true;
    /// validation.ErrorTitle = "An invalid date was entered";
    /// validation.Error = "The date must be between 2011-01-31 and 2011-12-31";
    /// validation.Prompt = "Enter date here";
    /// validation.Formula.Value = DateTime.Parse("2011-01-01");
    /// validation.Formula2.Value = DateTime.Parse("2011-12-31");
    /// validation.Operator = ExcelDataValidationOperator.between;
    /// </code>
    /// </summary>
    public class ExcelDataValidationCollection : XmlHelper, IEnumerable<ExcelDataValidation>
    {
        private List<ExcelDataValidation> _validations = new List<ExcelDataValidation>();
        private ExcelWorksheet _worksheet = null;

        private const string DataValidationPath = "//d:dataValidations";
        private readonly string DataValidationItemsPath = string.Format("{0}/d:dataValidation", DataValidationPath);

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet"></param>
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
                    _validations.Add(ExcelDataValidationFactory.Create(type, worksheet, addr, node));
                }
            }

        }

        private void EnsureRootElementExists()
        {
            var node = _worksheet.WorksheetXml.SelectSingleNode(DataValidationPath, _worksheet.NameSpaceManager);
            if (node == null)
            {
                CreateNode(DataValidationPath.TrimStart('/'));
            }
        }

        private XmlNode GetRootNode()
        {
            EnsureRootElementExists();
            return _worksheet.WorksheetXml.SelectSingleNode(DataValidationPath, _worksheet.NameSpaceManager);
        }

        private void ValidateAddress(string address)
        {
            Require.Argument(address).IsNotNullOrEmpty("address");
            
            // ensure that the new address does not collide with an existing validation.
            var newAddress = new ExcelAddress(address);
            if (_validations.Count > 0)
            {
                foreach (var validation in _validations)
                {
                    var result = validation.Address.Collide(newAddress);
                    if (result != ExcelAddressBase.eAddressCollition.No)
                    {
                        throw new InvalidOperationException(string.Format("The new address ({0}) collides with an existing validation ({1})", address, validation.Address.Address));
                    }
                }
            }
        }

        /// <summary>
        /// Adds an <see cref="ExcelDataValidationInt"/> to the worksheet. Whole means that the only accepted values
        /// are integer values.
        /// </summary>
        /// <param name="address">the range/address to validate</param>
        public ExcelDataValidationInt AddWholeValidation(string address)
        {
            ValidateAddress(address);
            EnsureRootElementExists(); 
            var item = new ExcelDataValidationInt(_worksheet, address, ExcelDataValidationType.Whole);
            _validations.Add(item);
            return item;
        }

        /// <summary>
        /// Addes an <see cref="ExcelDataValidation"/> to the worksheet. The only accepted values are
        /// decimal values.
        /// </summary>
        /// <param name="address">The range/address to validate</param>
        /// <returns></returns>
        public ExcelDataValidationDecimal AddDecimalValidation(string address)
        {
            ValidateAddress(address);
            EnsureRootElementExists();
            var item = new ExcelDataValidationDecimal(_worksheet, address, ExcelDataValidationType.Decimal);
            _validations.Add(item);
            return item;
        }

        /// <summary>
        /// Adds an <see cref="ExcelDataValidationList"/> to the worksheet. The accepted values are defined
        /// in a list.
        /// </summary>
        /// <param name="address">The range/address to validate</param>
        /// <returns></returns>
        public ExcelDataValidationList AddListValidation(string address)
        {
            ValidateAddress(address);
            EnsureRootElementExists();
            var item = new ExcelDataValidationList(_worksheet, address, ExcelDataValidationType.List);
            _validations.Add(item);
            return item;
        }

        /// <summary>
        /// Adds an <see cref="ExcelDataValidationInt"/> regarding text length to the worksheet.
        /// </summary>
        /// <param name="address">The range/address to validate</param>
        /// <returns></returns>
        public ExcelDataValidationInt AddTextLengthValidation(string address)
        {
            ValidateAddress(address);
            EnsureRootElementExists();
            var item = new ExcelDataValidationInt(_worksheet, address, ExcelDataValidationType.TextLength);
            _validations.Add(item);
            return item;
        }

        /// <summary>
        /// Adds an <see cref="ExcelDataValidationDateTime"/> to the worksheet.
        /// </summary>
        /// <param name="address">The range/address to validate</param>
        /// <returns></returns>
        public ExcelDataValidationDateTime AddDateTimeValidation(string address)
        {
            ValidateAddress(address);
            EnsureRootElementExists();
            var item = new ExcelDataValidationDateTime(_worksheet, address, ExcelDataValidationType.DateTime);
            _validations.Add(item);
            return item;
        }

        /// <summary>
        /// Adds a <see cref="ExcelDataValidationTime"/> to the worksheet.
        /// </summary>
        /// <param name="address">The range/address to validation</param>
        /// <returns></returns>
        public ExcelDataValidationTime AddTimeValidation(string address)
        {
            ValidateAddress(address);
            EnsureRootElementExists();
            var item = new ExcelDataValidationTime(_worksheet, address, ExcelDataValidationType.Time);
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
            TopNode.RemoveChild(item.TopNode);
            return _validations.Remove(item);
        }

        /// <summary>
        /// Number of validations
        /// </summary>
        public int Count
        {
            get { return _validations.Count; }
        }

        /// <summary>
        /// Index operator, returns by 0-based index
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public ExcelDataValidation this[int index]
        {
            get { return _validations[index]; }
            set { _validations[index] = value; }
        }

        /// <summary>
        /// Index operator, returns a data validation which address partly or exactly matches the searched address.
        /// </summary>
        /// <param name="address">A cell address or range</param>
        /// <returns>A <see cref="ExcelDataValidation"/> or null if no match</returns>
        public ExcelDataValidation this[string address]
        {
            get
            {
                var searchedAddress = new ExcelAddress(address);
                return _validations.Find(x => x.Address.Collide(searchedAddress) != ExcelAddressBase.eAddressCollition.No);
            }
        }

        /// <summary>
        /// Returns all validations that matches the supplied predicate <paramref name="match"/>.
        /// </summary>
        /// <param name="match">predicate to filter out matching validations</param>
        /// <returns></returns>
        public IEnumerable<ExcelDataValidation> FindAll(Predicate<ExcelDataValidation> match)
        {
            return _validations.FindAll(match);
        }

        /// <summary>
        /// Returns the first matching validation.
        /// </summary>
        /// <param name="match"></param>
        /// <returns></returns>
        public ExcelDataValidation Find(Predicate<ExcelDataValidation> match)
        {
            return _validations.Find(match);
        }

        /// <summary>
        /// Removes all validations from the collection.
        /// </summary>
        public void Clear()
        {
            DeleteAllNode(DataValidationItemsPath.TrimStart('/'));
            _validations.Clear();
        }

        /// <summary>
        /// Removes the validations that matches the predicate
        /// </summary>
        /// <param name="match"></param>
        public void RemoveAll(Predicate<ExcelDataValidation> match)
        {
            var matches = _validations.FindAll(match);
            foreach (var m in matches)
            {
                TopNode.SelectSingleNode(DataValidationPath.TrimStart('/'), NameSpaceManager).RemoveChild(m.TopNode);
            }
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
