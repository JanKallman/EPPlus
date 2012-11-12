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
 * Jan Källman		                Initial Release		        2009-12-22
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Collections;
using System.IO;
using System.Drawing;
using System.Linq;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils;
namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Collection for Drawing objects.
    /// </summary>
    public class ExcelDrawings : IEnumerable<ExcelDrawing>
    {
        private XmlDocument _drawingsXml=new XmlDocument();
        private Dictionary<string, int> _drawingNames;
        private List<ExcelDrawing> _drawings;
        internal class ImageCompare
        {
            internal byte[] image { get; set; }
            internal string relID { get; set; }

            internal bool Comparer(byte[] compareImg)
            {
                if (compareImg.Length != image.Length)
                {
                    return false;
                }

                for (int i = 0; i < image.Length; i++)
                {
                    if (image[i] != compareImg[i])
                    {
                        return false;
                    }
                }
                return true; //Equal
            }
        }
        //internal List<ImageCompare> _pics = new List<ImageCompare>();
        internal Dictionary<string, string> _hashes = new Dictionary<string, string>();
        internal ExcelPackage _package;
        internal Packaging.ZipPackageRelationship _drawingRelation = null;
        internal ExcelDrawings(ExcelPackage xlPackage, ExcelWorksheet sheet)
        {
                _drawingsXml = new XmlDocument();                
                _drawingsXml.PreserveWhitespace = false;
                _drawings = new List<ExcelDrawing>();
                _drawingNames = new Dictionary<string,int>();
                _package = xlPackage;
                Worksheet = sheet;
                XmlNode node = sheet.WorksheetXml.SelectSingleNode("//d:drawing", sheet.NameSpaceManager);
                CreateNSM();
                if (node != null)
                {
                    _drawingRelation = sheet.Part.GetRelationship(node.Attributes["r:id"].Value);
                    _uriDrawing = UriHelper.ResolvePartUri(sheet.WorksheetUri, _drawingRelation.TargetUri);

                    _part = xlPackage.Package.GetPart(_uriDrawing);
                    XmlHelper.LoadXmlSafe(_drawingsXml, _part.GetStream()); 

                    AddDrawings();
                }
         }
        internal ExcelWorksheet Worksheet { get; set; }
        /// <summary>
        /// A reference to the drawing xml document
        /// </summary>
        public XmlDocument DrawingXml
        {
            get
            {
                return _drawingsXml;
            }
        }
        private void AddDrawings()
        {
            XmlNodeList list = _drawingsXml.SelectNodes("//xdr:twoCellAnchor", NameSpaceManager);

            foreach (XmlNode node in list)
            {
                ExcelDrawing dr = ExcelDrawing.GetDrawing(this, node);
                _drawings.Add(dr);
                if (!_drawingNames.ContainsKey(dr.Name.ToLower()))
                {
                    _drawingNames.Add(dr.Name.ToLower(), _drawings.Count - 1);
                }
            }
        }


        #region NamespaceManager
        /// <summary>
        /// Creates the NamespaceManager. 
        /// </summary>
        private void CreateNSM()
        {
            NameTable nt = new NameTable();
            _nsManager = new XmlNamespaceManager(nt);
            _nsManager.AddNamespace("a", ExcelPackage.schemaDrawings);
            _nsManager.AddNamespace("xdr", ExcelPackage.schemaSheetDrawings);
            _nsManager.AddNamespace("c", ExcelPackage.schemaChart);
            _nsManager.AddNamespace("r", ExcelPackage.schemaRelationships);
        }
        /// <summary>
        /// Provides access to a namespace manager instance to allow XPath searching
        /// </summary>
        XmlNamespaceManager _nsManager=null;
        public XmlNamespaceManager NameSpaceManager
        {
            get
            {
                return _nsManager;
            }
        }
        #endregion
        #region IEnumerable Members

        public IEnumerator GetEnumerator()
        {
            return (_drawings.GetEnumerator());
        }
        #region IEnumerable<ExcelDrawing> Members

        IEnumerator<ExcelDrawing> IEnumerable<ExcelDrawing>.GetEnumerator()
        {
            return (_drawings.GetEnumerator());
        }

        #endregion

        /// <summary>
        /// Returns the drawing at the specified position.  
        /// </summary>
        /// <param name="PositionID">The position of the drawing. 0-base</param>
        /// <returns></returns>
        public ExcelDrawing this[int PositionID]
        {
            get
            {
                return (_drawings[PositionID]);
            }
        }

        /// <summary>
        /// Returns the drawing matching the specified name
        /// </summary>
        /// <param name="Name">The name of the worksheet</param>
        /// <returns></returns>
        public ExcelDrawing this[string Name]
        {
            get
            {
                if (_drawingNames.ContainsKey(Name.ToLower()))
                {
                    return _drawings[_drawingNames[Name.ToLower()]];
                }
                else
                {
                    return null;
                }
            }
        }
        public int Count
        {
            get
            {
                if (_drawings == null)
                {
                    return 0;
                }
                else
                {
                    return _drawings.Count;
                }
            }
        }
        Packaging.ZipPackagePart _part=null;
        internal Packaging.ZipPackagePart Part
        {
            get
            {
                return _part;
            }        
        }
        Uri _uriDrawing=null;
        public Uri UriDrawing
        {
            get
            {
                return _uriDrawing;
            }
        }
        #endregion
        #region Add functions
            /// <summary>
            /// Add a new chart to the worksheet.
            /// Do not support Bubble-, Radar-, Stock- or Surface charts. 
            /// </summary>
            /// <param name="Name"></param>
            /// <param name="ChartType">Type of chart</param>
            /// <param name="PivotTableSource">The pivottable source for a pivotchart</param>    
            /// <returns>The chart</returns>
            public ExcelChart AddChart(string Name, eChartType ChartType, ExcelPivotTable PivotTableSource)
            {
                if(_drawingNames.ContainsKey(Name.ToLower()))
                {
                    throw new Exception("Name already exists in the drawings collection");
                }

                if (ChartType == eChartType.Bubble ||
                    ChartType == eChartType.Bubble3DEffect ||
                    ChartType == eChartType.Radar ||
                    ChartType == eChartType.RadarFilled ||
                    ChartType == eChartType.RadarMarkers ||
                    ChartType == eChartType.StockHLC ||
                    ChartType == eChartType.StockOHLC ||
                    ChartType == eChartType.StockVOHLC ||
                    ChartType == eChartType.Surface ||
                    ChartType == eChartType.SurfaceTopView ||
                    ChartType == eChartType.SurfaceTopViewWireframe ||
                    ChartType == eChartType.SurfaceWireframe)
                {
                    throw(new NotImplementedException("Chart type is not supported in the current version"));
                }

                XmlElement drawNode = CreateDrawingXml();

                ExcelChart chart = ExcelChart.GetNewChart(this, drawNode, ChartType, null, PivotTableSource);
                chart.Name = Name;
                _drawings.Add(chart);
                _drawingNames.Add(Name.ToLower(), _drawings.Count - 1);
                return chart;
            }
            /// <summary>
            /// Add a new chart to the worksheet.
            /// Do not support Bubble-, Radar-, Stock- or Surface charts. 
            /// </summary>
            /// <param name="Name"></param>
            /// <param name="ChartType">Type of chart</param>
            /// <returns>The chart</returns>
            public ExcelChart AddChart(string Name, eChartType ChartType)
            {
                return AddChart(Name, ChartType, null);
            }
            /// <summary>
            /// Add a picure to the worksheet
            /// </summary>
            /// <param name="Name"></param>
            /// <param name="image">An image. Allways saved in then JPeg format</param>
            /// <returns></returns>
            public ExcelPicture AddPicture(string Name, Image image)
            {
               return AddPicture(Name, image, null);
            }
            /// <summary>
            /// Add a picure to the worksheet
            /// </summary>
            /// <param name="Name"></param>
            /// <param name="image">An image. Allways saved in then JPeg format</param>
            /// <param name="Hyperlink">Picture Hyperlink</param>
            /// <returns></returns>
            public ExcelPicture AddPicture(string Name, Image image, Uri Hyperlink)
            {
                if (image != null)
                {
                    if (_drawingNames.ContainsKey(Name.ToLower()))
                    {
                        throw new Exception("Name already exists in the drawings collection");
                    }
                    XmlElement drawNode = CreateDrawingXml();
                    drawNode.SetAttribute("editAs", "oneCell");
                    ExcelPicture pic = new ExcelPicture(this, drawNode, image, Hyperlink);
                    pic.Name = Name;
                    _drawings.Add(pic);
                    _drawingNames.Add(Name.ToLower(), _drawings.Count - 1);
                    return pic;
                }
                throw (new Exception("AddPicture: Image can't be null"));
            }
            /// <summary>
            /// Add a picure to the worksheet
            /// </summary>
            /// <param name="Name"></param>
            /// <param name="ImageFile">The image file</param>
            /// <returns></returns>
            public ExcelPicture AddPicture(string Name, FileInfo ImageFile)
            {
               return AddPicture(Name, ImageFile, null);
            }
            /// <summary>
            /// Add a picure to the worksheet
            /// </summary>
            /// <param name="Name"></param>
            /// <param name="ImageFile">The image file</param>
            /// <param name="Hyperlink">Picture Hyperlink</param>
            /// <returns></returns>
            public ExcelPicture AddPicture(string Name, FileInfo ImageFile, Uri Hyperlink)
            {
               if (ImageFile != null)
               {
                  if (_drawingNames.ContainsKey(Name.ToLower()))
                  {
                     throw new Exception("Name already exists in the drawings collection");
                  }
                  XmlElement drawNode = CreateDrawingXml();
                  drawNode.SetAttribute("editAs", "oneCell");
                  ExcelPicture pic = new ExcelPicture(this, drawNode, ImageFile, Hyperlink);
                  pic.Name = Name;
                  _drawings.Add(pic);
                  _drawingNames.Add(Name.ToLower(), _drawings.Count - 1);
                  return pic;
               }
               throw (new Exception("AddPicture: ImageFile can't be null"));
            }

        /// <summary>
        /// Add a new shape to the worksheet
        /// </summary>
        /// <param name="Name">Name</param>
        /// <param name="Style">Shape style</param>
        /// <returns>The shape object</returns>
    
        public ExcelShape AddShape(string Name, eShapeStyle Style)
            {
                if (_drawingNames.ContainsKey(Name.ToLower()))
                {
                    throw new Exception("Name already exists in the drawings collection");
                }
                XmlElement drawNode = CreateDrawingXml();
                ExcelShape shape = new ExcelShape(this, drawNode, Style);
                shape.Name = Name;
                shape.Style = Style;
                _drawings.Add(shape);
                _drawingNames.Add(Name.ToLower(), _drawings.Count - 1);
                return shape;
            }
            private XmlElement CreateDrawingXml()
            {
                if (DrawingXml.OuterXml == "")
                {
                    DrawingXml.LoadXml(string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><xdr:wsDr xmlns:xdr=\"{0}\" xmlns:a=\"{1}\" />", ExcelPackage.schemaSheetDrawings, ExcelPackage.schemaDrawings));
                    _uriDrawing = new Uri(string.Format("/xl/drawings/drawing{0}.xml", Worksheet.SheetID),UriKind.Relative);

                    Packaging.ZipPackage package = Worksheet._package.Package;
                    _part = package.CreatePart(_uriDrawing, "application/vnd.openxmlformats-officedocument.drawing+xml", _package.Compression);

                    StreamWriter streamChart = new StreamWriter(_part.GetStream(FileMode.Create, FileAccess.Write));
                    DrawingXml.Save(streamChart);
                    streamChart.Close();
                    package.Flush();

                    _drawingRelation = Worksheet.Part.CreateRelationship(UriHelper.GetRelativeUri(Worksheet.WorksheetUri, _uriDrawing), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/drawing");
                    XmlElement e = Worksheet.WorksheetXml.CreateElement("drawing", ExcelPackage.schemaMain);
                    e.SetAttribute("id",ExcelPackage.schemaRelationships, _drawingRelation.Id);

                    Worksheet.WorksheetXml.DocumentElement.AppendChild(e);
                    package.Flush();                    
                }
                XmlNode colNode = _drawingsXml.SelectSingleNode("//xdr:wsDr", NameSpaceManager);
                XmlElement drawNode = _drawingsXml.CreateElement("xdr", "twoCellAnchor", ExcelPackage.schemaSheetDrawings);
                colNode.AppendChild(drawNode);

                //Add from position Element;
                XmlElement fromNode = _drawingsXml.CreateElement("xdr","from", ExcelPackage.schemaSheetDrawings);
                drawNode.AppendChild(fromNode);
                fromNode.InnerXml = "<xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff>";

                //Add to position Element;
                XmlElement toNode = _drawingsXml.CreateElement("xdr", "to", ExcelPackage.schemaSheetDrawings);
                drawNode.AppendChild(toNode);
                toNode.InnerXml = "<xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff>";
                return drawNode;
            }
        #endregion
        #region Remove methods
            /// <summary>
            /// Removes a drawing.
            /// </summary>
            /// <param name="Index">The index of the drawing</param>
            public void Remove(int Index)
        {
            var draw=_drawings[Index];
            draw.DeleteMe();
            for (int i = Index + 1; i < _drawings.Count; i++)
            {
                _drawingNames[_drawings[i].Name.ToLower()]--;
            }
            _drawingNames.Remove(draw.Name.ToLower());
            _drawings.Remove(draw);
        }
        /// <summary>
        /// Removes a drawing.
        /// </summary>
        /// <param name="Drawing">The drawing</param>
        public void Remove(ExcelDrawing Drawing)
        {
            Remove(_drawingNames[Drawing.Name.ToLower()]);
        }
        /// <summary>
        /// Removes a drawing.
        /// </summary>
        /// <param name="Name">The name of the drawing</param>
        public void Remove(string Name)
        {
            Remove(_drawingNames[Name.ToLower()]);
        }
        /// <summary>
        /// Removes all drawings from the collection
        /// </summary>
        public void Clear()
        {
            while (Count > 0)
            {
                Remove(0);
            }
        }
        #endregion
            internal void AdjustWidth(int[,] pos)
            {
                var ix = 0;
                //Now set the size for all drawings depending on the editAs property.
                foreach (OfficeOpenXml.Drawing.ExcelDrawing d in this)
                {
                    if (d.EditAs != Drawing.eEditAs.TwoCell)
                    {
                        if (d.EditAs == Drawing.eEditAs.Absolute)
                        {
                            d.SetPixelLeft(pos[ix, 0]);
                        }
                        d.SetPixelWidth(pos[ix, 1]);
                        
                    }
                    ix++;
                }
            }
            internal void AdjustHeight(int[,] pos)
            {
                var ix = 0;
                //Now set the size for all drawings depending on the editAs property.
                foreach (OfficeOpenXml.Drawing.ExcelDrawing d in this)
                {
                    if (d.EditAs != Drawing.eEditAs.TwoCell)
                    {
                        if (d.EditAs == Drawing.eEditAs.Absolute)
                        {
                            d.SetPixelTop(pos[ix, 0]);
                        }
                        d.SetPixelHeight(pos[ix, 1]);

                    }
                    ix++;
                }
            }
            internal int[,] GetDrawingWidths()
            {
                int[,] pos = new int[Count, 2];
                int ix = 0;
                //Save the size for all drawings
                foreach (ExcelDrawing d in this)
                {
                    pos[ix, 0] = d.GetPixelLeft();
                    pos[ix++, 1] = d.GetPixelWidth();
                }
                return pos;
            }
            internal int[,] GetDrawingHeight()
            {
                int[,] pos = new int[Count, 2];
                int ix = 0;
                //Save the size for all drawings
                foreach (ExcelDrawing d in this)
                {
                    pos[ix, 0] = d.GetPixelTop();
                    pos[ix++, 1] = d.GetPixelHeight();
                }
                return pos;
            }
    }
}
