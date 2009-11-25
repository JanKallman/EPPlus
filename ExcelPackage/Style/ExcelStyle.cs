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
    public class ExcelStyle : StyleBase
    {
        internal ExcelStyle(ExcelStyles styles, OfficeOpenXml.XmlHelper.ChangedEventHandler ChangedEvent, int PositionID, string Address, int xfsId) :
            base(styles, ChangedEvent, PositionID, Address)
        {
            Index = xfsId;
            ExcelXfs xfs = _styles.CellXfs[xfsId];
            Numberformat = new ExcelNumberFormat(styles, ChangedEvent, PositionID, Address, xfs.NumberFormatId);
            Font = new ExcelFont(styles, ChangedEvent, PositionID, Address, xfs.FontId);
            Fill = new ExcelFill(styles, ChangedEvent, PositionID, Address, xfs.FillId);
            Border = new Border(styles, ChangedEvent, PositionID, Address, xfs.BorderId); 
        }
        public ExcelNumberFormat Numberformat { get; set; }
        public ExcelFont Font { get; set; }
        public ExcelFill Fill { get; set; }
        public Border Border { get; set; }
        ExcelHorizontalAlignment _horizontalAlignment = ExcelHorizontalAlignment.Left;
        public ExcelHorizontalAlignment HorizontalAlignment
        {
            get
            {
                return _styles.CellXfs[Index].HorizontalAlignment;
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.HorizontalAlign, value, _positionID, _address));
            }
        }
        ExcelVerticalAlignment _verticalAlignment = ExcelVerticalAlignment.Bottom;
        const string verticalAlignPath = "d:alignment/@vertical";
        public ExcelVerticalAlignment VerticalAlignment
        {
            get
            {
                return _styles.CellXfs[Index].VerticalAlignment;
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.VerticalAlign, value, _positionID, _address));
            }
        }
        bool _wrapText=false;
        const string wrapTextPath = "d:alignment/@wrapText";
        public bool WrapText
        {
            get
            {
                return _styles.CellXfs[Index].WrapText;
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.WrapText, value, _positionID, _address));
            }
        }
        bool _readingOrder = false;
        const string readingOrderPath = "d:alignment/@readingOrder";
        public bool ReadingOrder
        {
            get
            {
                return _styles.CellXfs[Index].ReadingOrder;
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.ReadingOrder, value, _positionID, _address));
            }
        }
        bool _shrinkToFit = false;
        const string shrinkToFitPath = "d:alignment/@shrinkToFit";
        public bool ShrinkToFit
        {
            get
            {
                return _styles.CellXfs[Index].ShrinkToFit;
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.ShrinkToFit, value, _positionID, _address));
            }
        }

        int _xfId=int.MinValue;
        const string xfIdPath = "@xfid";
        public int XfId 
        {
            get
            {
                return _styles.CellXfs[Index].XfId;
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.XfId, value, _positionID, _address));
            }
        }
        internal override string Id
        {
            get 
            { 
                return Numberformat.Id + "|" + Font.Id + "|" + Fill.Id + "|" + Border.Id + "|" + VerticalAlignment + "|" + HorizontalAlignment + "|" + WrapText.ToString() + "|" + ReadingOrder.ToString() + "|" + XfId.ToString(); 
            }
        }

    }
}
