/* 
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * EPPlus provides server-side generation of Excel 2007 spreadsheets.
 *
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
 * All code and executables are provided     "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * 
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Jan Källman		                Initial Release		        2009-10-01
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Globalization;

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// The numberformat of the cell
    /// </summary>
    public sealed class ExcelNumberFormat : StyleBase
    {
        internal ExcelNumberFormat(ExcelStyles styles, OfficeOpenXml.XmlHelper.ChangedEventHandler ChangedEvent, int PositionID, string Address, int index) :
            base(styles, ChangedEvent, PositionID, Address)
        {
            Index = index;
        }

        public int NumFmtID 
        {
            get
            {
                return Index;
            }
            //set
            //{
            //    _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Numberformat, "NumFmtID", value, _workSheetID, _address));
            //}
        }
        /// <summary>
        /// The numberformat 
        /// </summary>
        public string Format
        {
            get
            {
                for(int i=0;i<_styles.NumberFormats.Count;i++)
                {
                    if(Index==_styles.NumberFormats[i].NumFmtId)
                    {
                        return _styles.NumberFormats[i].Format;
                    }
                }
                return "general";
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Numberformat, eStyleProperty.Format, value, _positionID, _address));
            }
        }

        internal override string Id
        {
            get 
            {
                return Format;
            }
        }
        public bool BuildIn { get; private set; }
    }
}
