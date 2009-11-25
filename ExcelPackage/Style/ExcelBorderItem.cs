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
 * Jan Källman		                Initial Release		        2009-10-01
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml.Style.XmlAccess;

namespace OfficeOpenXml.Style
{
    public class ExcelBorderItem : StyleBase
    {
        eStyleClass _cls;
        StyleBase _parent;
        internal ExcelBorderItem (ExcelStyles styles, OfficeOpenXml.XmlHelper.ChangedEventHandler ChangedEvent, int worksheetID, string address, eStyleClass cls, StyleBase parent) : 
            base(styles, ChangedEvent, worksheetID, address)
	    {
            _cls=cls;
            _parent = parent;
	    }
        public ExcelBorderStyle Style
        {
            get
            {
                return GetSource().Style;
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(_cls, eStyleProperty.Style, value, _positionID, _address));
            }
        }
        ExcelColor _color=null;
        public ExcelColor Color
        {
            get
            {
                if (_color == null)
                {
                    _color = new ExcelColor(_styles, _ChangedEvent, _positionID, _address, _cls, _parent);
                }
                return _color;
            }
        }

        internal override string Id
        {
            get { return Style + Color.Id; }
        }
        internal override void SetIndex(int index)
        {
            _parent.Index = index;
        }
        private ExcelBorderItemXml GetSource()
        {
            int ix = _parent.Index < 0 ? 0 : _parent.Index;

            switch(_cls)
            {
                case eStyleClass.BorderTop:
                    return _styles.Borders[ix].Top;
                    break;
                case eStyleClass.BorderBottom:
                    return _styles.Borders[ix].Bottom;
                    break;
                case eStyleClass.BorderLeft:
                    return _styles.Borders[ix].Left;
                    break;
                case eStyleClass.BorderRight:
                    return _styles.Borders[ix].Right;
                    break;
                case eStyleClass.BorderDiagonal:
                    return _styles.Borders[ix].Diagonal;
                    break;
                default:
                    throw new Exception("Invalid class for Borderitem");
            }

        }
    }
}
