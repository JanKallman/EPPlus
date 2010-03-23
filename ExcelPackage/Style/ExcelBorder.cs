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
    public class Border : StyleBase
    {
        internal Border(ExcelStyles styles, OfficeOpenXml.XmlHelper.ChangedEventHandler ChangedEvent, int PositionID, string address, int index) :
            base(styles, ChangedEvent, PositionID, address)
	    {
            Index = index;
        }
        public ExcelBorderItem Left
        {
            get
            {
                return new ExcelBorderItem(_styles, _ChangedEvent, _positionID, _address, eStyleClass.BorderLeft, this);
            }
        }
        public ExcelBorderItem Right
        {
            get
            {
                return new ExcelBorderItem(_styles, _ChangedEvent, _positionID, _address, eStyleClass.BorderRight, this);
            }
        }
        public ExcelBorderItem Top
        {
            get
            {
                return new ExcelBorderItem(_styles, _ChangedEvent, _positionID, _address, eStyleClass.BorderTop, this);
            }
        }
        public ExcelBorderItem Bottom
        {
            get
            {
                return new ExcelBorderItem(_styles, _ChangedEvent, _positionID, _address, eStyleClass.BorderBottom, this);
            }
        }
        public ExcelBorderItem Diagonal
        {
            get
            {
                return new ExcelBorderItem(_styles, _ChangedEvent, _positionID, _address, eStyleClass.BorderDiagonal, this);
            }
        }
        public bool DiagonalUp 
        {
            get
            {
                if (Index >=0)
                {
                    return _styles.Borders[Index].DiagonalUp;
                }
                else
                {
                    return false;
                }
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Border, eStyleProperty.BorderDiagonalUp, value, _positionID, _address));
            }
        }
        public bool DiagonalDown 
        {
            get
            {
                if (Index >= 0)
                {
                    return _styles.Borders[Index].DiagonalDown;
                }
                else
                {
                    return false;
                }
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Border, eStyleProperty.BorderDiagonalDown, value, _positionID, _address));
            }
        }
        internal override string Id
        {
            get { return Top.Id + Bottom.Id +Left.Id + Right.Id + Diagonal.Id + DiagonalUp + DiagonalDown; }
        }
    }
}
