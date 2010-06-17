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
using System.Drawing;

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// Cell style Font
    /// </summary>
    public sealed class ExcelFont : StyleBase
    {
        internal ExcelFont(ExcelStyles styles, OfficeOpenXml.XmlHelper.ChangedEventHandler ChangedEvent, int PositionID, string address, int index) :
            base(styles, ChangedEvent, PositionID, address)

        {
            Index = index;
        }
        /// <summary>
        /// The name of the font
        /// </summary>
        public string Name
        {
            get
            {
                return _styles.Fonts[Index].Name;
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Font, eStyleProperty.Name, value, _positionID, _address));
            }
        }
        /// <summary>
        /// The Size of the font
        /// </summary>
        public float Size
        {
            get
            {
                return _styles.Fonts[Index].Size;
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Font, eStyleProperty.Size, value, _positionID, _address));
            }
        }
        /// <summary>
        /// Font family
        /// </summary>
        public int Family
        {
            get
            {
                return _styles.Fonts[Index].Family;
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Font, eStyleProperty.Family, value, _positionID, _address));
            }
        }
        /// <summary>
        /// Cell color
        /// </summary>
        public ExcelColor Color
        {
            get
            {
                return new ExcelColor(_styles, _ChangedEvent, _positionID, _address, eStyleClass.Font, this);
            }
        }
        /// <summary>
        /// Scheme
        /// </summary>
        public string Scheme
        {
            get
            {
                return _styles.Fonts[Index].Scheme;
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Font, eStyleProperty.Scheme, value, _positionID, _address));
            }
        }
        /// <summary>
        /// Font-bold
        /// </summary>
        public bool Bold
        {
            get
            {
                return _styles.Fonts[Index].Bold;
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Font, eStyleProperty.Bold, value, _positionID, _address));
            }
        }
        /// <summary>
        /// Font-italic
        /// </summary>
        public bool Italic
        {
            get
            {
                return _styles.Fonts[Index].Italic;
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Font, eStyleProperty.Italic, value, _positionID, _address));
            }
        }
        /// <summary>
        /// Font-Strikeout
        /// </summary>
        public bool Strike
        {
            get
            {
                return _styles.Fonts[Index].Strike;
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Font, eStyleProperty.Strike, value, _positionID, _address));
            }
        }
        /// <summary>
        /// Font-Underline
        /// </summary>
        public bool UnderLine
        {
            get
            {
                return _styles.Fonts[Index].UnderLine;
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Font, eStyleProperty.Underline, value, _positionID, _address));
            }
        }
        /// <summary>
        /// Font-Vertical Align
        /// </summary>
        public ExcelVerticalAlignmentFont VerticalAlign
        {
            get
            {
                return (ExcelVerticalAlignmentFont)Enum.Parse(typeof(ExcelVerticalAlignmentFont), _styles.Fonts[Index].VerticalAlign, true);
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Font, eStyleProperty.VerticalAlign, value, _positionID, _address));
            }
        }
        /// <summary>
        /// Set the font from a Font object
        /// </summary>
        /// <param name="Font"></param>
        public void SetFromFont(Font Font)
        {
            Name = Font.Name;
            //Family=fnt.FontFamily.;
            Size = (int)Font.Size;
            Strike = Font.Strikeout;
            Bold = Font.Bold;
            UnderLine = Font.Underline;
            Italic = Font.Italic;
        }

        internal override string Id
        {
            get 
            {
                return Name + Size.ToString() + Family.ToString() + Scheme.ToString() + Bold.ToString()[0] + Italic.ToString()[0] + Strike.ToString()[0] + UnderLine.ToString()[0] + VerticalAlign;
            }
        }
    }
}
