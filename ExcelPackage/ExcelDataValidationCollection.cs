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

namespace OfficeOpenXml
{
    /// <summary>
    /// Collection of <see cref="ExcelDataValidation"/>
    /// </summary>
    public class ExcelDataValidationCollection : IEnumerable<ExcelDataValidation>
    {
        private List<ExcelDataValidation> _validations = new List<ExcelDataValidation>();

        /// <summary>
        /// Adds an <see cref="ExcelDataValidation"/> to the collection.
        /// </summary>
        /// <param name="item">The item to add</param>
        /// <exception cref="ArgumentNullException">if <paramref name="item"/> is null</exception>
        public void Add(ExcelDataValidation item)
        {
            Require.Argument(item).IsNotNull("item");
            _validations.Add(item);
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
