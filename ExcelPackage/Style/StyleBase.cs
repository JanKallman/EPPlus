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
    public enum ExcelBorderStyle
    {
        None,
        Hair,
        Dotted,
        DashDot,
        Thin,
        DashDotDot,
        Dashed,
        MediumDashDotDot,
        MediumDashed,
        MediumDashDot,
        Thick,
        Medium,
        Double
    };
    public enum ExcelHorizontalAlignment
    {
        Left,
        Center,
        CenterContinuous,
        Right,
        Fill,
        Distributed,
        Justify
    }
    public enum ExcelVerticalAlignment
    {
        Top,
        Center,
        Bottom,
        Distributed,
        Justify
    }
    public enum ExcelFillStyle
    {
        None,
        Solid,
        DarkGray,
        MediumGray,
        LightGray,
        Gray125,
        Gray0625,
        DarkVertical,
        DarkHorizontal,
        DarkDown,
        DarkUp,
        DarkGrid,
        DarkTrellis,
        LightVertical,
        LightHorizontal,
        LightDown,
        LightUp,
        LightGrid,
        LightTrellis
    }
    public abstract class StyleBase
    {
        protected ExcelStyles _styles;
        internal OfficeOpenXml.XmlHelper.ChangedEventHandler _ChangedEvent;
        protected int _positionID;
        protected string _address;
        internal StyleBase(ExcelStyles styles, OfficeOpenXml.XmlHelper.ChangedEventHandler ChangedEvent, int PositionID, string Address)
        {
            _styles = styles;
            _ChangedEvent = ChangedEvent;
            _address = Address;
            _positionID = PositionID;
        }
        internal int Index { get; set;}
        internal abstract string Id {get;}

        internal virtual void SetIndex(int index)
        {
            Index = index;
        }
    }
}
