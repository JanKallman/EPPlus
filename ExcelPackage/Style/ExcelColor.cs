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
    public class ExcelColor :  StyleBase
    {
        eStyleClass _cls;
        ExcelColorXml _sourceColor;
        StyleBase _parent;
        internal ExcelColor(ExcelStyles styles, OfficeOpenXml.XmlHelper.ChangedEventHandler ChangedEvent, int worksheetID, string address, eStyleClass cls, StyleBase parent) : 
            base(styles, ChangedEvent, worksheetID, address)
            
        {
            _parent = parent;
            _cls = cls;
        }

        public string Theme
        {
            get
            {
                return GetSource().Theme;
            }
        }
        decimal _tint;
        public decimal Tint
        {
            get
            {
                return GetSource().Tint;
            }
        }
        string _rgb;
        public string Rgb
        {
            get
            {
                return GetSource().Rgb;
            }
            internal set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(_cls, eStyleProperty.Color, value, _positionID, _address));
            }
        }
        int _indexed;
        public int Indexed
        {
            get
            {
                return GetSource().Indexed;
            }
        }
        public void SetColor(System.Drawing.Color color)
        {
            Rgb = color.ToArgb().ToString("X");
        }


        internal override string Id
        {
            get 
            {
                return Theme + Tint + Rgb + Indexed;
            }
        }
        private ExcelColorXml GetSource()
        {
            Index = _parent.Index < 0 ? 0 : _parent.Index;
            switch(_cls)
            {
                case eStyleClass.FillBackgroundColor:
                    return _styles.Fills[Index].BackgroundColor;
                    break;
                case eStyleClass.FillPatternColor:
                    return _styles.Fills[Index].PatternColor;
                    break;
                case eStyleClass.Font:
                    return _styles.Fonts[Index].Color;
                    break;
                case eStyleClass.BorderLeft:
                    return _styles.Borders[Index].Left.Color;
                    break;
                case eStyleClass.BorderTop:
                    return _styles.Borders[Index].Top.Color;
                    break;
                case eStyleClass.BorderRight:
                    return _styles.Borders[Index].Right.Color;
                    break;
                case eStyleClass.BorderBottom:
                    return _styles.Borders[Index].Bottom.Color;
                    break;
                case eStyleClass.BorderDiagonal:
                    return _styles.Borders[Index].Diagonal.Color;
                    break;
                default:
                    throw(new Exception("Invalid style-class for Color"));
            }
        }
        internal override void SetIndex(int index)
        {
            _parent.Index = index;
        }
    }
}
