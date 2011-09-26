using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Collections;
using System.IO.Packaging;
using System.Globalization;

namespace OfficeOpenXml.Drawing.Vml
{
    public class ExcelVmlDrawingPictureCollection : ExcelVmlDrawingBaseCollection, IEnumerable
    {
        internal List<ExcelVmlDrawingPicture> _images;
        ExcelPackage _pck;
        ExcelWorksheet _ws;
        internal ExcelVmlDrawingPictureCollection(ExcelPackage pck, ExcelWorksheet ws, Uri uri) :
            base(pck, ws, uri)
        {            
            _pck = pck;
            _ws = ws;
            if (uri == null)
            {
                VmlDrawingXml.LoadXml(CreateVmlDrawings());
                _images = new List<ExcelVmlDrawingPicture>();
            }
            else
            {
                AddDrawingsFromXml();
            }
        }

        private void AddDrawingsFromXml()
        {
            var nodes = VmlDrawingXml.SelectNodes("//v:shape", NameSpaceManager);
            _images = new List<ExcelVmlDrawingPicture>();
            foreach (XmlNode node in nodes)
            {
                var img = new ExcelVmlDrawingPicture(node, NameSpaceManager, _ws);
                var rel = Part.GetRelationship(img.RelId);
                img.ImageUri = PackUriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
                _images.Add(img);
            }
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
        internal ExcelVmlDrawingPicture Add(string id, Uri uri, string name, double width, double height)
        {
            XmlNode node = AddImage(id, uri, name, width, height);
            var draw = new ExcelVmlDrawingPicture(node, NameSpaceManager, _ws);
            draw.ImageUri = uri;
            _images.Add(draw);
            return draw;
        }
        private XmlNode AddImage(string id, Uri targeUri, string Name, double width, double height)
        {
            var node = VmlDrawingXml.CreateElement("v", "shape", ExcelPackage.schemaMicrosoftVml);
            VmlDrawingXml.DocumentElement.AppendChild(node);
            node.SetAttribute("id", id);
            node.SetAttribute("o:type", "#_x0000_t75");
            node.SetAttribute("style", string.Format("position:absolute;margin-left:0;margin-top:0;width:{0}pt;height:{1}pt;z-index:1", width.ToString(CultureInfo.InvariantCulture), height.ToString(CultureInfo.InvariantCulture)));
            //node.SetAttribute("fillcolor", "#ffffe1");
            //node.SetAttribute("insetmode", ExcelPackage.schemaMicrosoftOffice, "auto");

            node.InnerXml = string.Format("<v:imagedata o:relid=\"\" o:title=\"{0}\"/><o:lock v:ext=\"edit\" rotation=\"t\"/>",  Name);
            return node;
        }
        /// <summary>
        /// Indexer
        /// </summary>
        /// <param name="Index">Index</param>
        /// <returns>The VML Drawing Picture object</returns>
        public ExcelVmlDrawingPicture this[int Index]
        {
            get
            {
                return _images[Index] as ExcelVmlDrawingPicture;
            }
        }
        public int Count
        {
            get
            {
                return _images.Count;
            }
        }

        int _nextID = 0;
        /// <summary>
        /// returns the next drawing id.
        /// </summary>
        /// <returns></returns>
        internal string GetNewId()
        {
            if (_nextID == 0)
            {
                foreach (ExcelVmlDrawingComment draw in this)
                {
                    if (draw.Id.Length > 3 && draw.Id.StartsWith("vml"))
                    {
                        int id;
                        if (int.TryParse(draw.Id.Substring(3, draw.Id.Length - 3), out id))
                        {
                            if (id > _nextID)
                            {
                                _nextID = id;
                            }
                        }
                    }
                }
            }
            _nextID++;
            return "vml" + _nextID.ToString();
        }
        #region IEnumerable Members

        IEnumerator IEnumerable.GetEnumerator() 
        {
            return _images.GetEnumerator();
        }

        #endregion
    }
}
