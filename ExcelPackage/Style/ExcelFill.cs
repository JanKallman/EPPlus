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

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// The background fill of a cell
    /// </summary>
    public sealed class ExcelFill : StyleBase
    {
        internal ExcelFill(ExcelStyles styles, OfficeOpenXml.XmlHelper.ChangedEventHandler ChangedEvent, int PositionID, string address, int index) :
            base(styles, ChangedEvent, PositionID, address)

        {
            Index = index;
        }
        /// <summary>
        /// The pattern of the fill
        /// </summary>
        public ExcelFillStyle PatternType
        {
            get
            {
                if (Index == int.MinValue)
                {
                    return ExcelFillStyle.None;
                }
                else
                {
                    return _styles.Fills[Index].PatternType;
                }
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Fill, eStyleProperty.PatternType, value, _positionID, _address));
            }
        }
        ExcelColor _patternColor = null;
        /// <summary>
        /// The color of the pattern
        /// </summary>
        public ExcelColor PatternColor
        {
            get
            {
                if (_patternColor == null)
                {
                    _patternColor = new ExcelColor(_styles, _ChangedEvent, _positionID, _address, eStyleClass.FillPatternColor, this);
                }
                return _patternColor;
            }
        }
        ExcelColor _backgroundColor = null;
        /// <summary>
        /// The background color
        /// </summary>
        public ExcelColor BackgroundColor
        {
            get
            {
                if (_backgroundColor == null)
                {
                    _backgroundColor = new ExcelColor(_styles, _ChangedEvent, _positionID, _address, eStyleClass.FillBackgroundColor, this);
                }
                return _backgroundColor;
                
            }
        }

        internal override string Id
        {
            get { return PatternType +PatternColor.Id+BackgroundColor.Id; }
        }
    }
}
