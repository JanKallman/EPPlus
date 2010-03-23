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
 * Jan Källman		                Initial Release		        2009-10-01
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using OfficeOpenXml.Style;
namespace OfficeOpenXml.Style.XmlAccess
{
    public class ExcelBorderItemXml : StyleXmlHelper
    {
        internal ExcelBorderItemXml(XmlNamespaceManager nameSpaceManager) : base(nameSpaceManager)
        {
            _borderStyle=ExcelBorderStyle.None;
            _color = new ExcelColorXml(NameSpaceManager);
        }
        internal ExcelBorderItemXml(XmlNamespaceManager nsm, XmlNode topNode) :
            base(nsm, topNode)
        {
            if (topNode != null)
            {
                _borderStyle = GetBorderStyle(GetXmlNode("@style"));
                _color = new ExcelColorXml(nsm, topNode.SelectSingleNode(_colorPath, nsm));
                Exists = true;
            }
            else
            {
                Exists = false;
            }
        }

        private ExcelBorderStyle GetBorderStyle(string style)
        {
            if(style=="") return ExcelBorderStyle.None;
            string sInStyle = style.Substring(0, 1).ToUpper() + style.Substring(1, style.Length - 1);
            try
            {
                return (ExcelBorderStyle)Enum.Parse(typeof(ExcelBorderStyle), sInStyle);
            }
            catch
            {
                return ExcelBorderStyle.None;
            }

        }
        ExcelBorderStyle _borderStyle = ExcelBorderStyle.None;
        public ExcelBorderStyle Style
        {
            get
            {
                return _borderStyle;
            }
            set
            {
                _borderStyle = value;
                Exists = true;
            }
        }
        ExcelColorXml _color = null;
        const string _colorPath = "d:color";
        public ExcelColorXml Color
        {
            get
            {
                return _color;
            }
            internal set
            {
                _color = value;
            }
        }
        internal override string Id
        {
            get { return Style + Color.Id; }
        }

        internal ExcelBorderItemXml Copy()
        {
            ExcelBorderItemXml borderItem = new ExcelBorderItemXml(NameSpaceManager);
            borderItem.Style = _borderStyle;
            borderItem.Color = _color.Copy();
            return borderItem;
        }

        internal override XmlNode CreateXmlNode(XmlNode topNode)
        {
            TopNode = topNode;

            if (Style != ExcelBorderStyle.None)
            {
                SetXmlNode("@style", SetBorderString(Style));
                if (Color.Exists)
                {
                    CreateNode(_colorPath);
                    topNode.AppendChild(Color.CreateXmlNode(TopNode.SelectSingleNode(_colorPath,NameSpaceManager)));
                }
            }
            return TopNode;
        }

        private string SetBorderString(ExcelBorderStyle Style)
        {
            string newName=Enum.GetName(typeof(ExcelBorderStyle), Style);
            return newName.Substring(0, 1).ToLower() + newName.Substring(1, newName.Length - 1);
        }
        public bool Exists { get; private set; }
    }
}
