using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Collections;
using System.IO.Packaging;

namespace OfficeOpenXml.Drawing.Vml
{
    internal class ExcelVmlDrawings : IEnumerable
    {
        internal RangeCollection _drawings;
        internal ExcelVmlDrawings(ExcelPackage pck, ExcelWorksheet ws, Uri uri)
        {
            VmlDrawingXml = new XmlDocument();
            VmlDrawingXml.PreserveWhitespace = false;
            
            NameTable nt=new NameTable();
            NameSpaceManager = new XmlNamespaceManager(nt);
            NameSpaceManager.AddNamespace("v", ExcelPackage.schemaMicrosoftVml);
            NameSpaceManager.AddNamespace("o", ExcelPackage.schemaMicrosoftOffice);
            NameSpaceManager.AddNamespace("x", ExcelPackage.schemaMicrosoftExcel);
            Uri = uri;
            if (uri == null)
            {
                VmlDrawingXml.LoadXml(CreateVmlDrawings());
                _drawings = new RangeCollection(new List<IRangeID>());
                Part = null;
            }
            else
            {
                Part=pck.Package.GetPart(uri);
                VmlDrawingXml.Load(Part.GetStream());
                AddDrawingsFromXml(ws);
            }
        }

        private void AddDrawingsFromXml(ExcelWorksheet ws)
        {
            var nodes = VmlDrawingXml.SelectNodes("//v:shape", NameSpaceManager);
            var list = new List<IRangeID>();
            foreach (XmlNode node in nodes)
            {
            	var rowNode = node.SelectSingleNode("x:ClientData/x:Row", NameSpaceManager);
				var colNode = node.SelectSingleNode("x:ClientData/x:Column", NameSpaceManager);
				if (rowNode != null && colNode != null)
				{
					var row = int.Parse(rowNode.InnerText) + 1;
					var col = int.Parse(colNode.InnerText) + 1;
					list.Add(new ExcelVmlDrawing(node, ws.Cells[row, col], NameSpaceManager));
				}
				else
				{
					list.Add(new ExcelVmlDrawing(node, ws.Cells[1, 1], NameSpaceManager));
				}
            }
            _drawings = new RangeCollection(list);
        }

        private string CreateVmlDrawings()
        {
            string vml=string.Format("<xml xmlns:v=\"{0}\" xmlns:o=\"{1}\" xmlns:x=\"{2}\">", 
                ExcelPackage.schemaMicrosoftVml, 
                ExcelPackage.schemaMicrosoftOffice, 
                ExcelPackage.schemaMicrosoftExcel);
            
             vml+="<o:shapelayout v:ext=\"edit\">";
             vml+="<o:idmap v:ext=\"edit\" data=\"1\"/>";
             vml+="</o:shapelayout>";

             vml+="<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\">";
             vml+="<v:stroke joinstyle=\"miter\" />";
             vml+="<v:path gradientshapeok=\"t\" o:connecttype=\"rect\" />";
             vml+="</v:shapetype>";
             vml+= "</xml>";

            return vml;
        }
        internal ExcelVmlDrawing Add(ExcelRangeBase cell)
        {
            XmlNode node = AddDrawing(cell.Start.Row, cell.Start.Column);
            var draw = new ExcelVmlDrawing(node, cell, NameSpaceManager);
            _drawings.Add(draw);
            return draw;
        }
        private XmlNode AddDrawing(int row, int col)
        {
            var node = VmlDrawingXml.CreateElement("v", "shape", ExcelPackage.schemaMicrosoftVml);
            VmlDrawingXml.DocumentElement.AppendChild(node);
            node.SetAttribute("id", _drawings.GetNewId());
            node.SetAttribute("type", "#_x0000_t202");
            node.SetAttribute("style", "position:absolute;z-index:1; visibility:hidden");
            //node.SetAttribute("style", "position:absolute; margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:1; visibility:hidden"); 
            node.SetAttribute("fillcolor", "#ffffe1");
            node.SetAttribute("insetmode",ExcelPackage.schemaMicrosoftOffice,"auto");

            string vml = "<v:fill color2=\"#ffffe1\" />";
            vml += "<v:shadow on=\"t\" color=\"black\" obscured=\"t\" />";
            vml += "<v:path o:connecttype=\"none\" />";
            vml += "<v:textbox style=\"mso-direction-alt:auto\">";
            vml += "<div style=\"text-align:left\" />";
            vml += "</v:textbox>";
            vml += "<x:ClientData ObjectType=\"Note\">";
            vml += "<x:MoveWithCells />";
            vml += "<x:SizeWithCells />";
            vml += string.Format("<x:Anchor>{0}, 15, {1}, 2, {2}, 31, {3}, 1</x:Anchor>", col, row - 1, col + 2, row + 3);
            vml += "<x:AutoFill>False</x:AutoFill>";
            vml += string.Format("<x:Row>{0}</x:Row>", row - 1); ;
            vml += string.Format("<x:Column>{0}</x:Column>", col - 1);
            vml += "</x:ClientData>";

            node.InnerXml = vml;
            return node;
        }
        internal ExcelVmlDrawing this[ulong rangeID]
        {
            get
            {
                return _drawings[rangeID] as ExcelVmlDrawing;
            }
        }
        internal bool ContainsKey(ulong rangeID)
        {
            return _drawings.ContainsKey(rangeID);
        }
        internal int Count
        {
            get
            {
                return _drawings.Count;
            }
        }
        internal XmlDocument VmlDrawingXml { get; set; }
        internal Uri Uri { get; set; }
        internal string RelId { get; set; }
        internal PackagePart Part { get; set; }
        internal XmlNamespaceManager NameSpaceManager { get; set; }

        #region IEnumerable Members

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _drawings;
        }

        #endregion
    }
}
