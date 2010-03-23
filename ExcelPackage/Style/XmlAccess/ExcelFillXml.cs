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
    public class ExcelFillXml : StyleXmlHelper 
    {
        internal ExcelFillXml(XmlNamespaceManager nameSpaceManager)
            : base(nameSpaceManager)
        {
            _fillPatternType = ExcelFillStyle.None;
            _backgroundColor = new ExcelColorXml(NameSpaceManager);
            _patternColor = new ExcelColorXml(NameSpaceManager);
        }
        internal ExcelFillXml(XmlNamespaceManager nsm, XmlNode topNode):
            base(nsm, topNode)
        {
            PatternType = GetPatternType(GetXmlNode(fillPatternTypePath));
            _backgroundColor = new ExcelColorXml(nsm, topNode.SelectSingleNode(_backgroundColorPath, nsm));
            _patternColor = new ExcelColorXml(nsm, topNode.SelectSingleNode(_patternColorPath, nsm));
        }

        private ExcelFillStyle GetPatternType(string patternType)
        {
            if (patternType == "") return ExcelFillStyle.None;
            patternType = patternType.Substring(0, 1).ToUpper() + patternType.Substring(1, patternType.Length - 1);
            try
            {
                return (ExcelFillStyle)Enum.Parse(typeof(ExcelFillStyle), patternType);
            }
            catch
            {
                return ExcelFillStyle.None;
            }
        }
        internal override string Id
        {
            get
            {
                return PatternType + PatternColor.Id + BackgroundColor.Id;
            }
        }
        #region Public Properties
        const string fillPatternTypePath = "d:patternFill/@patternType";
        ExcelFillStyle _fillPatternType;
        public ExcelFillStyle PatternType
        {
            get
            {
                return _fillPatternType;
            }
            set
            {
                _fillPatternType=value;
            }
        }
        ExcelColorXml _patternColor = null;
        const string _patternColorPath = "d:patternFill/d:bgColor";
        public ExcelColorXml PatternColor
        {
            get
            {
                return _patternColor;
            }
            internal set
            {
                _patternColor = value;
            }
        }
        ExcelColorXml _backgroundColor = null;
        const string _backgroundColorPath = "d:patternFill/d:fgColor";
        public ExcelColorXml BackgroundColor
        {
            get
            {
                return _backgroundColor;
            }
            internal set
            {
                _backgroundColor=value;
            }
        }
        #endregion


        //internal Fill Copy()
        //{
        //    Fill newFill = new Fill(NameSpaceManager, TopNode.Clone());
        //    return newFill;
        //}

        internal ExcelFillXml Copy()
        {
            ExcelFillXml newFill = new ExcelFillXml(NameSpaceManager);
            newFill.PatternType = _fillPatternType;
            newFill.BackgroundColor = _backgroundColor.Copy();
            newFill.PatternColor = _patternColor.Copy();
            return newFill;
        }

        internal override XmlNode CreateXmlNode(XmlNode topNode)
        {
            TopNode = topNode;
            SetXmlNode(fillPatternTypePath, SetPatternString(_fillPatternType));
            if (PatternType != ExcelFillStyle.None)
            {
                XmlNode pattern = topNode.SelectSingleNode(fillPatternTypePath, NameSpaceManager);
                if (BackgroundColor.Exists)
                {
                    CreateNode(_backgroundColorPath);
                    BackgroundColor.CreateXmlNode(topNode.SelectSingleNode(_backgroundColorPath, NameSpaceManager));
                    if (PatternColor.Exists)
                    {
                        CreateNode(_patternColorPath);
                        //topNode.SelectSingleNode(_bgColorPath, NameSpaceManager).AppendChild(PatternColor.CreateXmlNode(TopNode.OwnerDocument.CreateElement("bgColor", ExcelPackage.schemaMain)));
                        topNode.AppendChild(PatternColor.CreateXmlNode(topNode.SelectSingleNode(_patternColorPath, NameSpaceManager)));
                    }
                }
            }
            return topNode;
        }

        private string SetPatternString(ExcelFillStyle pattern)
        {
            string newName = Enum.GetName(typeof(ExcelFillStyle), pattern);
            return newName.Substring(0, 1).ToLower() + newName.Substring(1, newName.Length - 1);
        }
    }
}
