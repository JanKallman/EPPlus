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

/* 
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * ExcelPackage provides server-side generation of Excel 2007 spreadsheets.
 * See http://www.codeplex.com/ExcelPackage for details.
 * 
 * Copyright 2007 © Dr John Tunnicliffe 
 * mailto:dr.john.tunnicliffe@btinternet.com
 * All rights reserved.
 * 
 * ExcelPackage is an Open Source project provided under the 
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
 */
/*
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * John Tunnicliffe		Initial Release		01-Jan-2007
 * ******************************************************************************
 */
using System;
using System.Xml;

namespace OfficeOpenXml
{
	/// <summary>
	/// Represents the different view states of the worksheet
	/// </summary>
	public class ExcelWorksheetView : XmlHelper
	{
		// TODO: implement the different view states of the worksheet
		private ExcelWorksheet _xlWorksheet;
		private XmlElement _sheetView;

		#region ExcelWorksheetView Constructor
		/// <summary>
		/// Creates a new ExcelWorksheetView which provides access to all the 
		/// view states of the worksheet.
		/// </summary>
		/// <param name="xlWorksheet"></param>
		protected internal ExcelWorksheetView(XmlNamespaceManager ns, XmlNode node,  ExcelWorksheet xlWorksheet) :
            base(ns, node)
		{
			_xlWorksheet = xlWorksheet;
		}
		#endregion

		#region SheetViewElement
		/// <summary>
		/// Returns a reference to the sheetView element
		/// </summary>
		protected internal XmlElement SheetViewElement
		{
			get 
			{
				return (XmlElement)TopNode;
			}
		}
		#endregion
		#region TabSelected
        private XmlElement _selectionNode = null;
        private XmlElement SelectionNode
        {
            get
            {
                _selectionNode = SheetViewElement.SelectSingleNode("//d:selection", _xlWorksheet.NameSpaceManager) as XmlElement;
                if (_selectionNode == null)
                {
                    _selectionNode = _xlWorksheet.WorksheetXml.CreateElement("selection", ExcelPackage.schemaMain);
                    SheetViewElement.AppendChild(_selectionNode);
                }
                return _selectionNode;
            }
        }
        #endregion
        #region Public functions

        /// <summary>
        /// Indicates if the worksheet is selected within the workbook
        /// </summary>
        public bool TabSelected
        {
            get
            {
                return GetXmlNodeBool("@tabSelected");
            }
            set
            {
                if (value)
                {
                    //    // ensure no other worksheet has its tabSelected attribute set to 1
                    foreach (ExcelWorksheet sheet in _xlWorksheet.xlPackage.Workbook.Worksheets)
                        sheet.View.TabSelected = false;

                    SheetViewElement.SetAttribute("tabSelected", "1");
                    XmlElement bookView = _xlWorksheet.Workbook.WorkbookXml.SelectSingleNode("//d:workbookView", _xlWorksheet.NameSpaceManager) as XmlElement;
                    if (bookView != null)
                    {
                        bookView.SetAttribute("activeTab", (_xlWorksheet.PositionID - 1).ToString());
                    }
                }
                else
                    SetXmlNode("@tabSelected", "0");

            }
        }

        const string _selectionRangePath = "d:selection/@sqref";
        /// <summary>
        /// Selected Cells.Used in combination with ActiveCell
        /// </summary>        
        public string SelectedRange
        {
            get 
            {
                string address=GetXmlNode(_selectionRangePath);
                if (address == "")
                {
                    return "A1";
                }
                return address;
                //if(_selectionNode==null)
                //{
                //    return "A1";
                //}
                //return SelectionNode.GetAttribute("sqref");
            }
            set
            {                
                int fromCol, fromRow, toCol, toRow;
                ExcelCellBase.GetRowColFromAddress(value, out fromRow, out fromCol, out toRow, out toCol);
                //SelectionNode.SetAttribute("sqref",value);
                SetXmlNode(_selectionRangePath, value);
                if (SelectionNode.GetAttribute("activeCell") == "")
                {

                    ActiveCell = ExcelCellBase.GetAddress(fromRow, fromCol);
                }
                else
                {
                   //TODO:Add fix for out of range here
                }
            }
        }
        const string _activeCellPath = "d:selection/@activeCell";
        /// <summary>
        /// Set the active cell. Must be set within the SelectedRange.
        /// </summary>
        public string ActiveCell
        {
            get
            {
                string address = GetXmlNode(_activeCellPath);
                if (address == "")
                {
                    return "A1";
                }
                return address;
            }
            set
            {
                int fromCol, fromRow, toCol, toRow;
                ExcelCellBase.GetRowColFromAddress(value, out fromRow, out fromCol, out toRow, out toCol);
                SetXmlNode(_activeCellPath, value);
                if (SelectionNode.GetAttribute("sqref") == "")
                {

                    SelectedRange = ExcelCellBase.GetAddress(fromRow, fromCol);
                }
                else
                {
                    //TODO:Add fix for out of range here
                }
            }
        }

		/// <summary>
		/// Sets the view mode of the worksheet to pageLayout
		/// </summary>
		public bool PageLayoutView
		{
			get
			{
                return GetXmlNodeBool("@view");
			}
			set
			{
                if (value)
                    SetXmlNode("@view", "pageLayout"); //  SheetViewElement.SetAttribute("view", "pageLayout");
                else
                    SheetViewElement.RemoveAttribute("view");
			}
		}
        /// <summary>
        /// Show gridlines in the worksheet
        /// </summary>
        public bool ShowGridLines 
        {
            get
            {
                return GetXmlNodeBool("@showGridLines");
            }
            set
            {
                SetXmlNode("@showGridLines", value ? "1" : "0");
            }
        }
        /// <summary>
        /// Scale 
        /// </summary>
        public int ZoomScale
        {
            get
            {
                return GetXmlNodeInt("@zoomScale");
            }
            set
            {
                SetXmlNode("@zoomScale", value.ToString());
            }
        }
        #endregion
    }
}
