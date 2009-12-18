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
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.IO;
using System.IO.Packaging;
using System.Drawing;
using System.Drawing.Imaging;

namespace OfficeOpenXml.Drawing
{
    public class ExcelPicture : ExcelDrawing
    {
        internal ExcelPicture(ExcelDrawings drawings, XmlNode node) :
            base(drawings, node, "xdr:pic/xdr:nvPicPr/xdr:cNvPr/@name")
        {
            XmlNode picNode = node.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip", drawings.NameSpaceManager);
            if (picNode != null)
            {
                PackageRelationship drawingRelation = drawings.Part.GetRelationship(picNode.Attributes["r:embed"].Value);
                UriPic = PackUriHelper.ResolvePartUri(drawings.UriDrawing, drawingRelation.TargetUri);

                PackagePart part = drawings.Part.Package.GetPart(UriPic);
                _image = Image.FromStream(part.GetStream());
            }
            else
            {
            }
        }
        internal ExcelPicture(ExcelDrawings drawings, XmlNode node, Image image) :
            base(drawings, node, "xdr:pic/xdr:nvPicPr/xdr:cNvPr/@name")
        {
            XmlElement picNode = node.OwnerDocument.CreateElement("xdr", "pic", ExcelPackage.schemaSheetDrawings);
            node.InsertAfter(picNode,node.SelectSingleNode("xdr:to",NameSpaceManager));
            picNode.InnerXml = PicStartXml();

            node.InsertAfter(node.OwnerDocument.CreateElement("xdr", "clientData", ExcelPackage.schemaSheetDrawings), picNode);

            Package package = drawings.Worksheet.xlPackage.Package;
            UriPic = GetNewUri(package, "/xl/media/image{0}.jpeg");
            Part = package.CreatePart(UriPic, "image/jpeg", CompressionOption.NotCompressed);

            //Set the Image and save it to the package.
            Image=image;
            SetPosDefaults(Image);
            //Create relationship
            PackageRelationship picRelation = drawings.Part.CreateRelationship(UriPic, TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");
            node.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip/@r:embed", NameSpaceManager).Value=picRelation.Id;
            package.Flush();
        }
        private void SetPosDefaults(Image image)
        {
            SetPixelWidth(image.Width, image.HorizontalResolution);
            SetPixelHeight(image.Height, image.VerticalResolution);
        }
        private string PicStartXml()
        {
            StringBuilder xml = new StringBuilder();
            xml.AppendFormat("<xdr:nvPicPr><xdr:cNvPr id=\"2\" descr=\"\" />");
            xml.Append("<xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\" /></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"\" cstate=\"print\" /><a:stretch><a:fillRect /> </a:stretch> </xdr:blipFill> <xdr:spPr> <a:xfrm> <a:off x=\"0\" y=\"0\" />  <a:ext cx=\"0\" cy=\"0\" /> </a:xfrm> <a:prstGeom prst=\"rect\"> <a:avLst /> </a:prstGeom> </xdr:spPr>");
            return xml.ToString();
        }

        Image _image = null;
        public Image Image 
        {
            get
            {
                return _image;
            }
            set
            {
                _image = value;
                _image.Save(Part.GetStream(FileMode.Create, FileAccess.Write), ImageFormat.Jpeg);   //Always JPEG here at this point. 
            }
        }
        internal Uri UriPic { get; set; }
        internal PackagePart Part;

        internal string Id
        {
            get { return Name; }
        }
    }
}
