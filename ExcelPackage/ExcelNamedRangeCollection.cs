/* 
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * EPPlus provides server-side generation of Excel 2007 spreadsheets.
 * EPPlus is a fork of the ExcelPackage project
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
 * 
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Jan Källman		                Added this class		        2010-01-28
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;

namespace OfficeOpenXml
{
    public class ExcelNamedRangeCollection : IEnumerable<ExcelNamedRange>
    {
        internal ExcelWorksheet _ws;
        internal ExcelNamedRangeCollection()
        {
            _ws = null;
        }
        internal ExcelNamedRangeCollection(ExcelWorksheet ws)
        {
            _ws = ws;
        }
        Dictionary<string, ExcelNamedRange> _dic = new Dictionary<string, ExcelNamedRange>();
        public ExcelNamedRange Add(string Name, ExcelRangeBase Range)
        {
            var item = new ExcelNamedRange(Name, _ws , Range.Worksheet, Range.Address);
            _dic.Add(Name, item);
            return item;
        }
        public void Remove(string Name)
        {
            _dic.Remove(Name);
        }
        public bool ContainsKey(string key)
        {
            return _dic.ContainsKey(key);
        }
        public int Count
        {
            get
            {
                return _dic.Count;
            }
        }
        /// <summary>
        /// Names
        /// </summary>
        /// <param name="Name">The name</param>
        /// <returns></returns>
        public ExcelNamedRange this[string Name]
        {
            get
            {
                return _dic[Name];
            }
        }

        #region "IEnumerable"
        #region IEnumerable<ExcelNamedRange> Members
        public IEnumerator<ExcelNamedRange> GetEnumerator()
        {
            return _dic.Values.GetEnumerator();
        }
        #endregion
        #region IEnumerable Members
        IEnumerator IEnumerable.GetEnumerator()
        {
            return _dic.Values.GetEnumerator();
        }

        #endregion
        #endregion
    }
}
