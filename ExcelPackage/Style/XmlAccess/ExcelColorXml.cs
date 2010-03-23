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
    public class ExcelColorXml : StyleXmlHelper
    {
        internal ExcelColorXml(XmlNamespaceManager nameSpaceManager)
            : base(nameSpaceManager)
        {
            _auto = "";
            _theme = "";
            _tint = 0;
            _rgb = "";
            _indexed = int.MinValue;
        }
        internal ExcelColorXml(XmlNamespaceManager nsm, XmlNode topNode) :
            base(nsm, topNode)
        {
            if(topNode==null)
            {
                _exists=false;
            }
            else
            {
                _exists = true;
                _auto = GetXmlNode("@auto");
                _theme = GetXmlNode("@theme");
                _tint = GetXmlNodeDecimal("@tint");
                _rgb = GetXmlNode("@rgb");
                _indexed = GetXmlNodeInt("@indexed");
            }
        }
        
        internal override string Id
        {
            get
            {
                return _auto + "|" + _theme + "|" + _tint + "|" + _rgb + "|" + _indexed;
            }
        }
        string _auto;
        public string Auto
        {
            get
            {
                return _auto;
            }
        }
        string _theme;
        public string Theme
        {
            get
            {
                return _theme;
            }
        }
        decimal _tint;
        public decimal Tint
        {
            get
            {
                return _tint;
            }
        }
        string _rgb;
        public string Rgb
        {
            get
            {
                return _rgb;
            }
            set
            {
                _rgb = value;
                _exists=true;
            }
        }
        int _indexed;
        public int Indexed
        {
            get
            {
              return _indexed;
            }
        }
        public void SetColor(System.Drawing.Color color)
        {
            //XmlNode node = TopNode.SelectSingleNode(_parentPath, NameSpaceManager);
            _theme = "";
            _tint = decimal.MaxValue;
            _indexed=int.MinValue;
            _rgb = color.ToArgb().ToString("X");
        }

        internal ExcelColorXml Copy()
        {
            return new ExcelColorXml(NameSpaceManager) {_indexed=Indexed, _tint=Tint, _rgb=Rgb, _theme=Theme, _auto=Auto, _exists=Exists };
        }

        internal override XmlNode CreateXmlNode(XmlNode topNode)
        {
            TopNode = topNode;
            if(_rgb!="")
            {
                SetXmlNode("@rgb", _rgb);
            }
            else if (_indexed >= 0)
            {
                SetXmlNode("@indexed", _indexed.ToString());
            }
            else if (_auto != "")
            {
                SetXmlNode("@auto", _auto);
            }
            else
            {
                SetXmlNode("@theme", _theme.ToString());
                SetXmlNode("@tint", _tint.ToString());
            }
            return TopNode;
        }

        bool _exists;
        internal bool Exists
        {
            get
            {
                return _exists;
            }
        }
    }
}
