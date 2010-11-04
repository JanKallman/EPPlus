using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml
{
    public class ExcelWorkbookView : XmlHelper
    {
        #region ExcelWorksheetView Constructor
        /// <summary>
        /// Creates a new ExcelWorksheetView which provides access to all the 
        /// view states of the worksheet.
        /// </summary>
        /// <param name="ns"></param>
        /// <param name="node"></param>
        internal ExcelWorkbookView(XmlNamespaceManager ns, XmlNode node) :
            base(ns, node)
		{
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
    