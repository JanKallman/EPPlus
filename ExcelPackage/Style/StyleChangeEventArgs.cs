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
    internal enum eStyleClass
    {
        Numberformat,
        Font,    
        Border,
        BorderTop,
        BorderLeft,
        BorderBottom,
        BorderRight,
        BorderDiagonal,
        Fill,
        FillBackgroundColor,
        FillPatternColor,
        NamedStyle,
        Style
    };
    internal enum eStyleProperty
    {
        Format,
        Name,
        Size,
        Bold,
        Italic,
        Strike,
        Color,
        Family,
        Scheme,
        Underline,
        HorizontalAlign,
        VerticalAlign,
        Border,
        NamedStyle,
        Style,
        PatternType,
        ReadingOrder,
        WrapText,
        TextRotation,
        Locked,
        Hidden,
        ShrinkToFit,
        BorderDiagonalUp,
        BorderDiagonalDown,
        XfId,
        Indent
    }
    internal class StyleChangeEventArgs : EventArgs
    {
        internal StyleChangeEventArgs(eStyleClass styleclass, eStyleProperty styleProperty, object value, int positionID, string address)
        {
            StyleClass = styleclass;
            StyleProperty=styleProperty;
            Value = value;
            Address = address;
            PositionID = positionID;
        }
        internal eStyleClass StyleClass;
        internal eStyleProperty StyleProperty;
        //internal string PropertyName;
        internal object Value;
        internal int PositionID { get; set; }
        //internal string Address;
        internal string Address;
    }
}
