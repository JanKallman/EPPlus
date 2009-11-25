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

namespace OfficeOpenXml.Drawing
{
    public class ExcelPicture : ExcelDrawing
    {
        internal ExcelPicture(ExcelDrawings drawings, XmlNode node) :
            base(drawings, node, "xdr:pic/xdr:nvPicPr/xdr:cNvPr/@name")
        {
            XmlNode chartNode = node.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip", drawings.NameSpaceManager);
            if (chartNode != null)
            {
                PackageRelationship drawingRelation = drawings.Part.GetRelationship(chartNode.Attributes["r:embed"].Value);
                UriPic = PackUriHelper.ResolvePartUri(drawings.UriDrawing, drawingRelation.TargetUri);

                PackagePart part = drawings.Part.Package.GetPart(UriPic);
                Image = Image.FromStream(part.GetStream());
            }
            else
            {
            }
        }
        public Image Image { get; set; }
        //const string namePath="xdr:pic/xdr:nvPicPr/xdr:cNvPr/@name";
        //public override string Name
        //{
        //    get
        //    {
        //        return GetXmlNode(namePath);
        //    }
        //    set
        //    {
        //        SetXmlNode(namePath, value);
        //    }
        //}
        internal Uri UriPic { get; set; }
        internal PackagePart Part;

        internal string Id
        {
            get { return Name; }
        }
    }
}
