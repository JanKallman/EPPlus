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
 * Jan Källman		                Initial Release		        2009-10-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
public enum eShapeStyle
{
    AccentBorderCallout1,
    AccentBorderCallout2,
    AccentBorderCallout3,
    AccentCallout1,
    AccentCallout2,
    AccentCallout3,
    ActionButtonBackPrevious,
    ActionButtonBeginning,
    ActionButtonBlank,
    ActionButtonDocument,
    ActionButtonEnd,
    ActionButtonForwardNext,
    ActionButtonHelp,
    ActionButtonHome,
    ActionButtonInformation,
    ActionButtonMovie,
    ActionButtonReturn,
    ActionButtonSound,
    Arc,
    BentArrow,
    BentConnector2,
    BentConnector3,
    BentConnector4,
    BentConnector5,
    BentUpArrow,
    Bevel,
    BlockArc,
    BorderCallout1,
    BorderCallout2,
    BorderCallout3,
    BracePair,
    BracketPair,
    Callout1,
    Callout2,
    Callout3,
    Can,
    ChartPlus,
    ChartStar,
    ChartX,
    Chevron,
    Chord,
    CircularArrow,
    Cloud,
    CloudCallout,
    Corner,
    CornerTabs,
    Cube,
    CurvedConnector2,
    CurvedConnector3,
    CurvedConnector4,
    CurvedConnector5,
    CurvedDownArrow,
    CurvedLeftArrow,
    CurvedRightArrow,
    CurvedUpArrow,
    Decagon,
    DiagStripe,
    Diamond,
    Dodecagon,
    Donut,
    DoubleWave,
    DownArrow,
    DownArrowCallout,
    Ellipse,
    EllipseRibbon,
    EllipseRibbon2,
    FlowChartAlternateProcess,
    FlowChartCollate,
    FlowChartConnector,
    FlowChartDecision,
    FlowChartDelay,
    FlowChartDisplay,
    FlowChartDocument,
    FlowChartExtract,
    FlowChartInputOutput,
    FlowChartInternalStorage,
    FlowChartMagneticDisk,
    FlowChartMagneticDrum,
    FlowChartMagneticTape,
    FlowChartManualInput,
    FlowChartManualOperation,
    FlowChartMerge,
    FlowChartMultidocument,
    FlowChartOfflineStorage,
    FlowChartOffpageConnector,
    FlowChartOnlineStorage,
    FlowChartOr,
    FlowChartPredefinedProcess,
    FlowChartPreparation,
    FlowChartProcess,
    FlowChartPunchedCard,
    FlowChartPunchedTape,
    FlowChartSort,
    FlowChartSummingJunction,
    FlowChartTerminator,
    FoldedCorner,
    Frame,
    Funnel,
    Gear6,
    Gear9,
    HalfFrame,
    Heart,
    Heptagon,
    Hexagon,
    HomePlate,
    HorizontalScroll,
    IrregularSeal1,
    IrregularSeal2,
    LeftArrow,
    LeftArrowCallout,
    LeftBrace,
    LeftBracket,
    LeftCircularArrow,
    LeftRightArrow,
    LeftRightArrowCallout,
    LeftRightCircularArrow,
    LeftRightRibbon,
    LeftRightUpArrow,
    LeftUpArrow,
    LightningBolt,
    Line,
    LineInv,
    MathDivide,
    MathEqual,
    MathMinus,
    MathMultiply,
    MathNotEqual,
    MathPlus,
    Moon,
    NonIsoscelesTrapezoid,
    NoSmoking,
    NotchedRightArrow,
    Octagon,
    Parallelogram,
    Pentagon,
    Pie,
    PieWedge,
    Plaque,
    PlaqueTabs,
    Plus,
    QuadArrow,
    QuadArrowCallout,
    Rect,
    Ribbon,
    Ribbon2,
    RightArrow,
    RightArrowCallout,
    RightBrace,
    RightBracket,
    Round1Rect,
    Round2DiagRect,
    Round2SameRect,
    RoundRect,
    RtTriangle,
    SmileyFace,
    Snip1Rect,
    Snip2DiagRect,
    Snip2SameRect,
    SnipRoundRect,
    SquareTabs,
    Star10,
    Star12,
    Star16,
    Star24,
    Star32,
    Star4,
    Star5,
    Star6,
    Star7,
    Star8,
    StraightConnector1,
    StripedRightArrow,
    Sun,
    SwooshArrow,
    Teardrop,
    Trapezoid,
    Triangle,
    UpArrow,
    UpArrowCallout,
    UpDownArrow,
    UpDownArrowCallout,
    UturnArrow,
    Wave,
    WedgeEllipseCallout,
    WedgeRectCallout,
    WedgeRoundRectCallout,
    VerticalScroll
}
public enum eTextAlignment
{
    Left,
    Center,
    Right,
    Distributed,
    Justified,
    JustifiedLow,
    ThaiDistributed
}
/// <summary>
/// Fillstyle.
/// </summary>
public enum eFillStyle
{
    NoFill,
    SolidFill,
    GradientFill,
    PatternFill,
    BlipFill,
    GroupFill
}
namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// An Excel shape.
    /// </summary>
    public sealed class ExcelShape : ExcelDrawing
    {
        internal ExcelShape(ExcelDrawings drawings, XmlNode node) :
            base(drawings, node, "xdr:sp/xdr:nvSpPr/xdr:cNvPr/@name")
        {
            init();
        }
        internal ExcelShape(ExcelDrawings drawings, XmlNode node, eShapeStyle style) :
            base(drawings, node, "xdr:sp/xdr:nvSpPr/xdr:cNvPr/@name")
        {
            init();
            XmlElement shapeNode = node.OwnerDocument.CreateElement("xdr", "sp", ExcelPackage.schemaSheetDrawings);
            shapeNode.SetAttribute("macro", "");
            shapeNode.SetAttribute("textlink", "");
            node.AppendChild(shapeNode);

            shapeNode.InnerXml = ShapeStartXml();
            node.AppendChild(shapeNode.OwnerDocument.CreateElement("xdr", "clientData", ExcelPackage.schemaSheetDrawings));
        }
        private void init()
        {
            SchemaNodeOrder = new string[] { "prstGeom", "ln", "pPr", "defRPr", "solidFill", "uFill", "latin", "cs", "r", "rPr", "t" };
        }
        #region "public methods"
        const string ShapeStylePath = "xdr:sp/xdr:spPr/a:prstGeom/@prst";
        /// <summary>
        /// Shape style
        /// </summary>
        public eShapeStyle Style
        {
            get
            {
                string v = GetXmlNodeString(ShapeStylePath);
                try
                {
                    return (eShapeStyle)Enum.Parse(typeof(eShapeStyle), v, true);
                }
                catch
                {
                    throw (new Exception(string.Format("Invalid shapetype {0}", v)));
                }
            }
            set
            {
                string v = value.ToString();
                v = v.Substring(0, 1).ToLower() + v.Substring(1, v.Length - 1);
                SetXmlNodeString(ShapeStylePath, v);
            }
        }
        ExcelDrawingFill _fill = null;
        /// <summary>
        /// Fill
        /// </summary>
        public ExcelDrawingFill Fill
        {
            get
            {
                if (_fill == null)
                {
                    _fill = new ExcelDrawingFill(NameSpaceManager, TopNode, "xdr:sp/xdr:spPr");
                }
                return _fill;
            }
        }
        ExcelDrawingBorder _border = null;
        /// <summary>
        /// Border
        /// </summary>
        public ExcelDrawingBorder Border        
        {
            get
            {
                if (_border == null)
                {
                    _border = new ExcelDrawingBorder(NameSpaceManager, TopNode, "xdr:sp/xdr:spPr/a:ln");
                }
                return _border;
            }
        }
        string[] paragraphNodeOrder = new string[] { "pPr", "defRPr", "solidFill", "uFill", "latin", "cs", "r", "rPr", "t" };
        const string PARAGRAPH_PATH = "xdr:sp/xdr:txBody/a:p";
        ExcelTextFont _font=null;
        public ExcelTextFont Font
        {
            get
            {
                if (_font == null)
                {
                    XmlNode node = TopNode.SelectSingleNode(PARAGRAPH_PATH, NameSpaceManager);
                    if(node==null)
                    {
                        Text="";    //Creates the node p element
                        node = TopNode.SelectSingleNode(PARAGRAPH_PATH, NameSpaceManager);
                    }
                    _font = new ExcelTextFont(NameSpaceManager, TopNode, "xdr:sp/xdr:txBody/a:p/a:pPr/a:defRPr", paragraphNodeOrder);
                }
                return _font;
            }
        }
        const string TextPath = "xdr:sp/xdr:txBody/a:p/a:r/a:t";
        /// <summary>
        /// Text inside the shape
        /// </summary>
        public string Text
        {
            get
            {
                return GetXmlNodeString(TextPath);
            }
            set
            {
                SetXmlNodeString(TextPath, value);
            }

        }
        string lockTextPath = "xdr:sp/@fLocksText";
        /// <summary>
        /// Lock drawing
        /// </summary>
        public bool LockText
        {
            get
            {
                return GetXmlNodeBool(lockTextPath, true);
            }
            set
            {
                SetXmlNodeBool(lockTextPath, value);
            }
        }
        ExcelParagraphCollection _richText = null;
        /// <summary>
        /// Richtext collection. Used to format specific parts of the text
        /// </summary>
        public ExcelParagraphCollection RichText
        {
            get
            {
                if (_richText == null)
                {
                    //XmlNode node=TopNode.SelectSingleNode(PARAGRAPH_PATH, NameSpaceManager);
                    //if (node == null)
                    //{
                    //    CreateNode(PARAGRAPH_PATH);
                    //}
                        _richText = new ExcelParagraphCollection(NameSpaceManager, TopNode, PARAGRAPH_PATH, paragraphNodeOrder);
                }
                return _richText;
            }
        }
        const string TextAnchoringPath = "xdr:sp/xdr:txBody/a:bodyPr/@anchor";
        /// <summary>
        /// Text Anchoring
        /// </summary>
        public eTextAnchoringType TextAnchoring
        {
            get
            {
                return GetTextAchoringEnum(GetXmlNodeString(TextAnchoringPath));
            }
            set
            {
                SetXmlNodeString(TextAnchoringPath, GetTextAchoringText(value));
            }
        }
        const string TextAnchoringCtlPath = "xdr:sp/xdr:txBody/a:bodyPr/@anchorCtr";
        /// <summary>
        /// Specifies the centering of the text box.
        /// </summary>
        public bool TextAnchoringControl
        {
            get
            {
                return GetXmlNodeBool(TextAnchoringCtlPath);
            }
            set
            {
                if (value)
                {
                    SetXmlNodeString(TextAnchoringCtlPath, "1");
                }
                else
                {
                    SetXmlNodeString(TextAnchoringCtlPath, "0");
                }
            }
        }
        const string TEXT_ALIGN_PATH = "xdr:sp/xdr:txBody/a:p/a:pPr/@algn";
        /// <summary>
        /// How the text is aligned
        /// </summary>
        public eTextAlignment TextAlignment
        {
            get
            {
               switch(GetXmlNodeString(TEXT_ALIGN_PATH))
               {
                   case "ctr":
                       return eTextAlignment.Center;
                   case "r":
                       return eTextAlignment.Right;
                   case "dist":
                       return eTextAlignment.Distributed;
                   case "just":
                       return eTextAlignment.Justified;
                   case "justLow":
                       return eTextAlignment.JustifiedLow;
                   case "thaiDist":
                       return eTextAlignment.ThaiDistributed;
                   default: 
                       return eTextAlignment.Left;
               }
            }
            set
            {
                switch (value)
                {
                    case eTextAlignment.Right:
                        SetXmlNodeString(TEXT_ALIGN_PATH, "r");
                        break;
                    case eTextAlignment.Center:
                        SetXmlNodeString(TEXT_ALIGN_PATH, "ctr");
                        break;
                    case eTextAlignment.Distributed:
                        SetXmlNodeString(TEXT_ALIGN_PATH, "dist");
                        break;
                    case eTextAlignment.Justified:
                        SetXmlNodeString(TEXT_ALIGN_PATH, "just");
                        break;
                    case eTextAlignment.JustifiedLow:
                        SetXmlNodeString(TEXT_ALIGN_PATH, "justLow");
                        break;
                    case eTextAlignment.ThaiDistributed:
                        SetXmlNodeString(TEXT_ALIGN_PATH, "thaiDist");
                        break;
                    default:
                        DeleteNode(TEXT_ALIGN_PATH);
                        break;
                }                
            }
        }
        const string INDENT_ALIGN_PATH = "xdr:sp/xdr:txBody/a:p/a:pPr/@lvl";
        /// <summary>
        /// Indentation
        /// </summary>
        public int Indent
        {
            get
            {
                return GetXmlNodeInt(INDENT_ALIGN_PATH);
            }
            set
            {
                if (value < 0 || value > 8)
                {
                    throw(new ArgumentOutOfRangeException("Indent level must be between 0 and 8"));
                }
                SetXmlNodeString(INDENT_ALIGN_PATH, value.ToString());
            }
        }
        const string TextVerticalPath = "xdr:sp/xdr:txBody/a:bodyPr/@vert";
        /// <summary>
        /// Vertical text
        /// </summary>
        public eTextVerticalType TextVertical
        {
            get
            {
                return GetTextVerticalEnum(GetXmlNodeString(TextVerticalPath));
            }
            set
            {
                SetXmlNodeString(TextVerticalPath, GetTextVerticalText(value));
            }
        }
        #endregion
        #region "Private Methods"
        private string ShapeStartXml()
        {
            StringBuilder xml = new StringBuilder();
            xml.AppendFormat("<xdr:nvSpPr><xdr:cNvPr id=\"{0}\" name=\"{1}\" /><xdr:cNvSpPr /></xdr:nvSpPr><xdr:spPr><a:prstGeom prst=\"rect\"><a:avLst /></a:prstGeom></xdr:spPr><xdr:style><a:lnRef idx=\"2\"><a:schemeClr val=\"accent1\"><a:shade val=\"50000\" /></a:schemeClr></a:lnRef><a:fillRef idx=\"1\"><a:schemeClr val=\"accent1\" /></a:fillRef><a:effectRef idx=\"0\"><a:schemeClr val=\"accent1\" /></a:effectRef><a:fontRef idx=\"minor\"><a:schemeClr val=\"lt1\" /></a:fontRef></xdr:style><xdr:txBody><a:bodyPr vertOverflow=\"clip\" rtlCol=\"0\" anchor=\"ctr\" /><a:lstStyle /><a:p></a:p></xdr:txBody>", _id, Name);
            return xml.ToString();
        }
        private string GetTextAchoringText(eTextAnchoringType value)
        {
            switch (value)
            {
                case eTextAnchoringType.Bottom:
                    return "b";
                case eTextAnchoringType.Center:
                    return "ctr";
                case eTextAnchoringType.Distributed:
                    return "dist";
                case eTextAnchoringType.Justify:
                    return "just";
                default:
                    return "t";
            }
        }
        private string GetTextVerticalText(eTextVerticalType value)
        {
            switch (value)
            {
                case eTextVerticalType.EastAsianVertical:
                    return "eaVert";
                case eTextVerticalType.MongolianVertical:
                    return "mongolianVert";
                case eTextVerticalType.Vertical:
                    return "vert";
                case eTextVerticalType.Vertical270:
                    return "vert270";
                case eTextVerticalType.WordArtVertical:
                    return "wordArtVert";
                case eTextVerticalType.WordArtVerticalRightToLeft:
                    return "wordArtVertRtl";
                default:
                    return "horz";
            }
        }
        private eTextVerticalType GetTextVerticalEnum(string text)
        {
            switch (text)
            {
                case "eaVert":
                    return eTextVerticalType.EastAsianVertical;
                case "mongolianVert":
                    return eTextVerticalType.MongolianVertical;
                case "vert":
                    return eTextVerticalType.Vertical;
                case "vert270":
                    return eTextVerticalType.Vertical270;
                case "wordArtVert":
                    return eTextVerticalType.WordArtVertical;
                case "wordArtVertRtl":
                    return eTextVerticalType.WordArtVerticalRightToLeft;
                default:
                    return eTextVerticalType.Horizontal;
            }
        }
        private eTextAnchoringType GetTextAchoringEnum(string text)
        {
            switch (text)
            {
                case "b":
                    return eTextAnchoringType.Bottom;
                case "ctr":
                    return eTextAnchoringType.Center;
                case "dist":
                    return eTextAnchoringType.Distributed;
                case "just":
                    return eTextAnchoringType.Justify;
                default:
                    return eTextAnchoringType.Top;
            }
        }
        #endregion
        internal new string Id
        {
            get { return Name + Text; }
        }
    }
}
