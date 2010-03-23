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
namespace OfficeOpenXml.Style.XmlAccess
{
    public class ExcelNamedStyleXml : StyleXmlHelper
    {
        ExcelStyles _styles;
        internal ExcelNamedStyleXml(XmlNamespaceManager nameSpaceManager, ExcelStyles styles)
            : base(nameSpaceManager)
        {
            _styles = styles;
            BuildInId = int.MinValue;
        }
        internal ExcelNamedStyleXml(XmlNamespaceManager NameSpaceManager, XmlNode topNode, ExcelStyles styles) :
            base(NameSpaceManager, topNode)
        {
            StyleXfId = GetXmlNodeInt(idPath);
            Name = GetXmlNode(namePath);
            BuildInId = GetXmlNodeInt(buildInIdPath);
            _styles = styles;
            _style = new ExcelStyle(styles, styles.NamedStylePropertyChange, -1, Name, _styleXfId);
        }
        internal override string Id
        {
            get
            {
                return Name;
            }
        }
        int _styleXfId=0;
        const string idPath = "@xfId";
        public int StyleXfId
        {
            get
            {
                return _styleXfId;
            }
            set
            {
                _styleXfId = value;
            }
        }
        int _xfId = int.MinValue;
        internal int XfId
        {
            get
            {
                return _xfId;
            }
            set
            {
                _xfId = value;
            }
        }
        const string buildInIdPath = "@builtinId";
        public int BuildInId { get; set; }
        const string namePath = "@name";
        string _name;
        public string Name
        {
            get
            {
                return _name;
            }
            internal set
            {
                _name = value;
            }
        }
        ExcelStyle _style = null;
        public ExcelStyle Style
        {
            get
            {
                return _style;
            }
            internal set
            {
                _style = value;
            }
        }

        internal override XmlNode CreateXmlNode(XmlNode topNode)
        {
            TopNode = topNode;
            SetXmlNode(namePath, _name);
            SetXmlNode("@xfId", _styles.CellStyleXfs[StyleXfId].newID.ToString());
            if (BuildInId>=0) SetXmlNode("@builtinId", BuildInId.ToString());
            return TopNode;            
        }
    }
}
