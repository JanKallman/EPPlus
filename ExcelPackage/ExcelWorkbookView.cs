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
 * Jan Källman		                Initial Release		        2010-11-02
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
namespace OfficeOpenXml
{
    /// <summary>
    /// Access to workbook view properties
    /// </summary>
    public class ExcelWorkbookView : XmlHelper
    {
        #region ExcelWorksheetView Constructor
        /// <summary>
        /// Creates a new ExcelWorkbookView which provides access to all the 
        /// view states of the worksheet.
        /// </summary>
        /// <param name="ns"></param>
        /// <param name="node"></param>
        internal ExcelWorkbookView(XmlNamespaceManager ns, XmlNode node, ExcelWorkbook wb) :
            base(ns, node)
		{
            SchemaNodeOrder = wb.SchemaNodeOrder;
		}
		#endregion
        const string LEFT_PATH="d:bookViews/d:workbookView/@xWindow";
        /// <summary>
        /// Position of the upper left corner of the workbook window. In twips.
        /// </summary>
        public int Left
        { 
            get
            {
                return GetXmlNodeInt(LEFT_PATH);
            }
            internal set
            {
                SetXmlNodeString(LEFT_PATH,value.ToString());
            }
        }
        const string TOP_PATH="d:bookViews/d:workbookView/@yWindow";
        /// <summary>
        /// Position of the upper left corner of the workbook window. In twips.
        /// </summary>
        public int Top
        { 
            get
            {
                return GetXmlNodeInt(TOP_PATH);
            }
            internal set
            {
                SetXmlNodeString(TOP_PATH, value.ToString());
            }
        }
        const string WIDTH_PATH="d:bookViews/d:workbookView/@windowWidth";
        /// <summary>
        /// Width of the workbook window. In twips.
        /// </summary>
        public int Width
        { 
            get
            {
                return GetXmlNodeInt(WIDTH_PATH);
            }
            internal set
            {
                SetXmlNodeString(WIDTH_PATH, value.ToString());
            }
        }
        const string HEIGHT_PATH="d:bookViews/d:workbookView/@windowHeight";
        /// <summary>
        /// Height of the workbook window. In twips.
        /// </summary>
        public int Height
        { 
            get
            {
                return GetXmlNodeInt(HEIGHT_PATH);
            }
            internal set
            {
                SetXmlNodeString(HEIGHT_PATH, value.ToString());
            }
        }
        const string MINIMIZED_PATH="d:bookViews/d:workbookView/@minimized";
        /// <summary>
        /// If true the the workbook window is minimized.
        /// </summary>
        public bool Minimized
        {
            get
            {
                return GetXmlNodeBool(MINIMIZED_PATH);
            }
            set
            {
                SetXmlNodeString(MINIMIZED_PATH, value.ToString());
            }
        }
        const string SHOWVERTICALSCROLL_PATH = "d:bookViews/d:workbookView/@showVerticalScroll";
        /// <summary>
        /// Show the vertical scrollbar
        /// </summary>
        public bool ShowVerticalScrollBar
        {
            get
            {
                return GetXmlNodeBool(SHOWVERTICALSCROLL_PATH,true);
            }
            set
            {
                SetXmlNodeBool(SHOWVERTICALSCROLL_PATH, value, true);
            }
        }
        const string SHOWHORIZONTALSCR_PATH = "d:bookViews/d:workbookView/@showHorizontalScroll";
        /// <summary>
        /// Show the horizontal scrollbar
        /// </summary>
        public bool ShowHorizontalScrollBar
        {
            get
            {
                return GetXmlNodeBool(SHOWHORIZONTALSCR_PATH, true);
            }
            set
            {
                SetXmlNodeBool(SHOWHORIZONTALSCR_PATH, value, true);
            }
        }
        const string SHOWSHEETTABS_PATH = "d:bookViews/d:workbookView/@showSheetTabs";
        /// <summary>
        /// Show the sheet tabs
        /// </summary>
        public bool ShowSheetTabs
        {
            get
            {
                return GetXmlNodeBool(SHOWSHEETTABS_PATH, true);
            }
            set
            {
                SetXmlNodeBool(SHOWSHEETTABS_PATH, value, true);
            }
        }
        /// <summary>
        /// Set the window position in twips
        /// </summary>
        /// <param name="left"></param>
        /// <param name="top"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        public void SetWindowSize(int left, int top, int width, int height)
        {
            Left = left;
            Top = top;
            Width = width;
            Height = height;
        }
    }
}
    