/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Jan Källman		    Initial Release		       2009-10-01
 * Jan Källman		    License changed GPL-->LGPL 2011-12-27
 *******************************************************************************/
using System;
using System.Xml;

namespace OfficeOpenXml
{
	/// <summary>
	/// Represents the different view states of the worksheet
	/// </summary>
	public class ExcelWorksheetView : XmlHelper
	{
        /// <summary>
        /// The worksheet panes after a freeze or split.
        /// </summary>
        public class ExcelWorksheetPanes : XmlHelper
        {
            XmlElement _selectionNode = null;
            internal ExcelWorksheetPanes(XmlNamespaceManager ns, XmlNode topNode) :
                base(ns, topNode)
            {
                if(topNode.Name=="selection")
                {
                    _selectionNode=topNode as XmlElement;
                }
            }

            const string _activeCellPath = "@activeCell";
            /// <summary>
            /// Set the active cell. Must be set within the SelectedRange.
            /// </summary>
            public string ActiveCell
            {
                get
                {
                    string address = GetXmlNodeString(_activeCellPath);
                    if (address == "")
                    {
                        return "A1";
                    }
                    return address;
                }
                set
                {
                    int fromCol, fromRow, toCol, toRow;
                    if(_selectionNode==null) CreateSelectionElement();
                    ExcelCellBase.GetRowColFromAddress(value, out fromRow, out fromCol, out toRow, out toCol);
                    SetXmlNodeString(_activeCellPath, value);
                    if (((XmlElement)TopNode).GetAttribute("sqref") == "")
                    {

                        SelectedRange = ExcelCellBase.GetAddress(fromRow, fromCol);
                    }
                    else
                    {
                        //TODO:Add fix for out of range here
                    }
                }
            }

            private void CreateSelectionElement()
            {
 	            _selectionNode=TopNode.OwnerDocument.CreateElement("selection", ExcelPackage.schemaMain);
                TopNode.AppendChild(_selectionNode);
                TopNode=_selectionNode;             
            }
            const string _selectionRangePath = "@sqref";
            /// <summary>
            /// Selected Cells.Used in combination with ActiveCell
            /// </summary>        
            public string SelectedRange
            {
                get
                {
                    string address = GetXmlNodeString(_selectionRangePath);
                    if (address == "")
                    {
                        return "A1";
                    }
                    return address;
                }
                set
                {
                    int fromCol, fromRow, toCol, toRow;
                    if(_selectionNode==null) CreateSelectionElement();
                    ExcelCellBase.GetRowColFromAddress(value, out fromRow, out fromCol, out toRow, out toCol);
                    SetXmlNodeString(_selectionRangePath, value);
                    if (((XmlElement)TopNode).GetAttribute("activeCell") == "")
                    {

                        ActiveCell = ExcelCellBase.GetAddress(fromRow, fromCol);
                    }
                    else
                    {
                        //TODO:Add fix for out of range here
                    }
                }
            }
        }
		private ExcelWorksheet _worksheet;

		#region ExcelWorksheetView Constructor
		/// <summary>
		/// Creates a new ExcelWorksheetView which provides access to all the view states of the worksheet.
		/// </summary>
        /// <param name="ns"></param>
        /// <param name="node"></param>
        /// <param name="xlWorksheet"></param>
		internal ExcelWorksheetView(XmlNamespaceManager ns, XmlNode node,  ExcelWorksheet xlWorksheet) :
            base(ns, node)
		{
            _worksheet = xlWorksheet;
            SchemaNodeOrder = new string[] { "sheetViews", "sheetView", "pane", "selection" };
            Panes = LoadPanes(); 
		}

		#endregion
        private ExcelWorksheetPanes[] LoadPanes()
        {
            XmlNodeList nodes = TopNode.SelectNodes("//d:selection", NameSpaceManager);
            if(nodes.Count==0)
            {
                return new ExcelWorksheetPanes[] { new ExcelWorksheetPanes(NameSpaceManager, TopNode) };
            }
            else
            {
                ExcelWorksheetPanes[] panes = new ExcelWorksheetPanes[nodes.Count];
                int i=0;
                foreach(XmlElement elem in nodes)
                {
                    panes[i++] = new ExcelWorksheetPanes(NameSpaceManager, elem);
                }
                return panes;
            }
        }
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
                _selectionNode = SheetViewElement.SelectSingleNode("//d:selection", _worksheet.NameSpaceManager) as XmlElement;
                if (_selectionNode == null)
                {
                    _selectionNode = _worksheet.WorksheetXml.CreateElement("selection", ExcelPackage.schemaMain);
                    SheetViewElement.AppendChild(_selectionNode);
                }
                return _selectionNode;
            }
        }
        #endregion
        #region Public Methods & Properties
        /// <summary>
        /// The active cell.
        /// </summary>
        public string ActiveCell
        {
            get
            {
                return Panes[Panes.GetUpperBound(0)].ActiveCell;
            }
            set 
            {
                Panes[Panes.GetUpperBound(0)].ActiveCell = value;
            }
        }
        /// <summary>
        /// Selected Cells in the worksheet.Used in combination with ActiveCell
        /// </summary>
        public string SelectedRange
        {
            get
            {
                return Panes[Panes.GetUpperBound(0)].SelectedRange;
            }
            set
            {
                Panes[Panes.GetUpperBound(0)].SelectedRange = value;
            }
        }
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
                    foreach (ExcelWorksheet sheet in _worksheet._package.Workbook.Worksheets)
                        sheet.View.TabSelected = false;

                    SheetViewElement.SetAttribute("tabSelected", "1");
                    XmlElement bookView = _worksheet.Workbook.WorkbookXml.SelectSingleNode("//d:workbookView", _worksheet.NameSpaceManager) as XmlElement;
                    if (bookView != null)
                    {
                        bookView.SetAttribute("activeTab", (_worksheet.PositionID - 1).ToString());
                    }
                }
                else
                    SetXmlNodeString("@tabSelected", "0");

            }
        }

		/// <summary>
		/// Sets the view mode of the worksheet to pagelayout
		/// </summary>
		public bool PageLayoutView
		{
			get
			{
                return GetXmlNodeString("@view") == "pageLayout";
			}
			set
			{
                if (value)
                    SetXmlNodeString("@view", "pageLayout");
                else
                    SheetViewElement.RemoveAttribute("view");
			}
		}
        /// <summary>
        /// Sets the view mode of the worksheet to pagebreak
        /// </summary>
        public bool PageBreakView
        {
            get
            {
                return GetXmlNodeString("@view") == "pageBreakPreview";
            }
            set
            {
                if (value)
                    SetXmlNodeString("@view", "pageBreakPreview");
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
                SetXmlNodeString("@showGridLines", value ? "1" : "0");
            }
        }
        /// <summary>
        /// Show the Column/Row headers (containg column letters and row numbers)
        /// </summary>
        public bool ShowHeaders
        {
            get
            {
                return GetXmlNodeBool("@showRowColHeaders");
            }
            set
            {
                SetXmlNodeString("@showRowColHeaders", value ? "1" : "0");
            }
        }
        /// <summary>
        /// Window zoom magnification for current view representing percent values.
        /// </summary>
        public int ZoomScale
        {
            get
            {
                return GetXmlNodeInt("@zoomScale");
            }
            set
            {
                if (value < 10 || value > 400)
                {
                    throw new ArgumentOutOfRangeException("Zoome scale out of range (10-400)");
                }
                SetXmlNodeString("@zoomScale", value.ToString());
            }
        }
        /// <summary>
        /// Flag indicating whether the sheet is in 'right to left' display mode. When in this mode,Column A is on the far right, Column B ;is one column left of Column A, and so on. Also,information in cells is displayed in the Right to Left format.
        /// </summary>
        public bool RightToLeft
        {
            get
            {
                return GetXmlNodeBool("@rightToLeft");
            }
            set
            {
                SetXmlNodeString("@rightToLeft", value == true ? "1" : "0");
            }
        }
        internal bool WindowProtection 
        {
            get
            {
                return GetXmlNodeBool("@windowProtection",false);
            }
            set
            {
                SetXmlNodeBool("@windowProtection",value,false);
            }
        }
        /// <summary>
        /// Reference to the panes
        /// </summary>
        public ExcelWorksheetPanes[] Panes
        {
            get;
            internal set;
        }
        string _paneNodePath = "d:pane";
        string _selectionNodePath = "d:selection";
        /// <summary>
        /// Freeze the columns/rows to left and above the cell
        /// </summary>
        /// <param name="Row"></param>
        /// <param name="Column"></param>
        public void FreezePanes(int Row, int Column)
        {
            //TODO:fix this method to handle splits as well.
            if (Row == 1 && Column == 1) UnFreezePanes();
            string sqRef = SelectedRange, activeCell = ActiveCell;
            
            XmlElement paneNode = TopNode.SelectSingleNode(_paneNodePath, NameSpaceManager) as XmlElement;
            if (paneNode == null)
            {
                CreateNode(_paneNodePath);
                paneNode = TopNode.SelectSingleNode(_paneNodePath, NameSpaceManager) as XmlElement;
            }
            paneNode.RemoveAll();   //Clear all attributes
            if (Column > 1) paneNode.SetAttribute("xSplit", (Column - 1).ToString());
            if (Row > 1) paneNode.SetAttribute("ySplit", (Row - 1).ToString());
            paneNode.SetAttribute("topLeftCell", ExcelCellBase.GetAddress(Row, Column));
            paneNode.SetAttribute("state", "frozen");

            RemoveSelection();

            if (Row > 1 && Column==1)
            {
                paneNode.SetAttribute("activePane", "bottomLeft");
                XmlElement sel=TopNode.OwnerDocument.CreateElement("selection", ExcelPackage.schemaMain);
                sel.SetAttribute("pane", "bottomLeft");
                if (activeCell != "") sel.SetAttribute("activeCell", activeCell);
                if (sqRef != "") sel.SetAttribute("sqref", sqRef);
                sel.SetAttribute("sqref", sqRef);
                TopNode.InsertAfter(sel, paneNode);
            }
            else if (Column > 1 && Row == 1)
            {
                paneNode.SetAttribute("activePane", "topRight");
                XmlElement sel = TopNode.OwnerDocument.CreateElement("selection", ExcelPackage.schemaMain);
                sel.SetAttribute("pane", "topRight");
                if (activeCell != "") sel.SetAttribute("activeCell", activeCell);
                if (sqRef != "") sel.SetAttribute("sqref", sqRef);
                TopNode.InsertAfter(sel, paneNode);
            }
            else
            {
                paneNode.SetAttribute("activePane", "bottomRight");
                XmlElement sel1 = TopNode.OwnerDocument.CreateElement("selection", ExcelPackage.schemaMain);
                sel1.SetAttribute("pane", "topRight");
                string cell = ExcelCellBase.GetAddress(1, Column);
                sel1.SetAttribute("activeCell", cell);
                sel1.SetAttribute("sqref", cell);
                paneNode.ParentNode.InsertAfter(sel1, paneNode);

                XmlElement sel2 = TopNode.OwnerDocument.CreateElement("selection", ExcelPackage.schemaMain);
                cell = ExcelCellBase.GetAddress(Row, 1);
                sel2.SetAttribute("pane", "bottomLeft");
                sel2.SetAttribute("activeCell", cell);
                sel2.SetAttribute("sqref", cell);
                sel1.ParentNode.InsertAfter(sel2, sel1);

                XmlElement sel3 = TopNode.OwnerDocument.CreateElement("selection", ExcelPackage.schemaMain);
                sel3.SetAttribute("pane", "bottomRight");
                if(activeCell!="") sel3.SetAttribute("activeCell", activeCell);                
                if(sqRef!="") sel3.SetAttribute("sqref", sqRef);
                sel2.ParentNode.InsertAfter(sel3, sel2);

            }
            Panes=LoadPanes();
        }
        private void RemoveSelection()
        {
            //Find selection nodes and remove them            
            XmlNodeList selections = TopNode.SelectNodes(_selectionNodePath, NameSpaceManager);
            foreach (XmlNode sel in selections)
            {
                sel.ParentNode.RemoveChild(sel);
            }
        }
        /// <summary>
        /// Unlock all rows and columns to scroll freely
        /// /// </summary>
        public void UnFreezePanes()
        {
            string sqRef = SelectedRange, activeCell = ActiveCell;

            XmlElement paneNode = TopNode.SelectSingleNode(_paneNodePath, NameSpaceManager) as XmlElement;
            if (paneNode != null)
            {
                paneNode.ParentNode.RemoveChild(paneNode);
            }
            RemoveSelection();

            Panes=LoadPanes();

            SelectedRange = sqRef;
            ActiveCell = activeCell;
        }
        #endregion
    }
}
