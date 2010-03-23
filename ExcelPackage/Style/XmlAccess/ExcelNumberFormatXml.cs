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
    public class ExcelNumberFormatXml : StyleXmlHelper
    {
        public ExcelNumberFormatXml(XmlNamespaceManager nameSpaceManager) : base(nameSpaceManager)
        {
            
        }        
        public ExcelNumberFormatXml(XmlNamespaceManager nameSpaceManager, bool buildIn): base(nameSpaceManager)
        {
            BuildIn = buildIn;
        }
        public ExcelNumberFormatXml(XmlNamespaceManager nsm, XmlNode topNode) :
            base(nsm, topNode)
        {
            _numFmtId = GetXmlNodeInt("@numFmtId");
            _format = GetXmlNode("@formatCode");
        }
        public bool BuildIn { get; private set; }
        int _numFmtId;
//        const string idPath = "@numFmtId";
        /// <summary>
        /// Id for number format
        /// 
        /// Build in ID's
        /// 
        /// 0   General 
        /// 1   0 
        /// 2   0.00 
        /// 3   #,##0 
        /// 4   #,##0.00 
        /// 9   0% 
        /// 10  0.00% 
        /// 11  0.00E+00 
        /// 12  # ?/? 
        /// 13  # ??/?? 
        /// 14  mm-dd-yy 
        /// 15  d-mmm-yy 
        /// 16  d-mmm 
        /// 17  mmm-yy 
        /// 18  h:mm AM/PM 
        /// 19  h:mm:ss AM/PM 
        /// 20  h:mm 
        /// 21  h:mm:ss 
        /// 22  m/d/yy h:mm 
        /// 37  #,##0 ;(#,##0) 
        /// 38  #,##0 ;[Red](#,##0) 
        /// 39  #,##0.00;(#,##0.00) 
        /// 40  #,##0.00;[Red](#,##0.00) 
        /// 45  mm:ss 
        /// 46  [h]:mm:ss 
        /// 47  mmss.0 
        /// 48  ##0.0E+0 
        /// 49  @
        /// </summary>            
        public int NumFmtId
        {
            get
            {
                return _numFmtId;
            }
            set
            {
                _numFmtId = value;
            }
        }
        internal override string Id
        {
            get
            {
                return _format;
            }
        }
        const string fmtPath = "@formatCode";
        string _format = string.Empty;
        public string Format
        {
            get
            {
                return _format;
            }
            set
            {
                _numFmtId = GetFromBuildIdFromFormat(value);
                _format = value;
            }
        }
        private string GetFromBuildInFromID(int _numFmtId)
        {
            switch (_numFmtId)
            {
                case 0:
                    return "General";
                case 1:
                    return "0";
                case 2:
                    return "0.00";
                case 3:
                    return "#,##0";
                case 4:
                    return "#,##0.00";
                case 9:
                    return "0%";
                case 10:
                    return "0.00%";
                case 11:
                    return "0.00E+00";
                case 12:
                    return "# ?/?";
                case 13:
                    return "# ??/??";
                case 14:
                    return "mm-dd-yy";
                case 15:
                    return "d-mmm-yy";
                case 16:
                    return "d-mmm";
                case 17:
                    return "mmm-yy";
                case 18:
                    return "h:mm AM/PM";
                case 19:
                    return "h:mm:ss AM/PM";
                case 20:
                    return "h:mm";
                case 21:
                    return "h:mm:ss";
                case 22:
                    return "m/d/yy h:mm";
                case 37:
                    return "#,##0 ;(#,##0)";
                case 38:
                    return "#,##0 ;[Red](#,##0)";
                case 39:
                    return "#,##0.00;(#,##0.00)";
                case 40:
                    return "#,##0.00;[Red](#,#)";
                case 45:
                    return "mm:ss";
                case 46:
                    return "[h]:mm:ss";
                case 47:
                    return "mmss.0";
                case 48:
                    return "##0.0";
                case 49:
                    return "@";
                default:
                    return string.Empty;
            }
        }
        private int GetFromBuildIdFromFormat(string format)
        {
            switch (format)
            {
                case "General":
                    return 0;
                case "0":
                    return 1;
                case "0.00":
                    return 2;
                case "#,##0":
                    return 3;
                case "#,##0.00":
                    return 4;
                case "0%":
                    return 9;
                case "0.00%":
                    return 10;
                case "0.00E+00":
                    return 11;
                case "# ?/?":
                    return 12;
                case "# ??/??":
                    return 13;
                case "mm-dd-yy":
                    return 14;
                case "d-mmm-yy":
                    return 15;
                case "d-mmm":
                    return 16;
                case "mmm-yy":
                    return 17;
                case "h:mm AM/PM":
                    return 18;
                case "h:mm:ss AM/PM":
                    return 19;
                case "h:mm":
                    return 20;
                case "h:mm:ss":
                    return 21;
                case "m/d/yy h:mm":
                    return 22;
                case "#,##0 ;(#,##0)":
                    return 37;
                case "#,##0 ;[Red](#,##0)":
                    return 38;
                case "#,##0.00;(#,##0.00)":
                    return 39;
                case "#,##0.00;[Red](#,#)":
                    return 40;
                case "mm:ss":
                    return 45;
                case "[h]:mm:ss":
                    return 46;
                case "mmss.0":
                    return 47;
                case "##0.0":
                    return 48;
                case "@":
                    return 49;
                default:
                    return int.MinValue;
            }
        }
        internal string GetNewID(int NumFmtId, string Format)
        {
            
            if (NumFmtId < 0)
            {
                NumFmtId = GetFromBuildIdFromFormat(Format);                
            }
            return NumFmtId.ToString();
        }

        internal static void AddBuildIn(XmlNamespaceManager NameSpaceManager, ExcelStyleCollection<ExcelNumberFormatXml> NumberFormats)
        {
            NumberFormats.Add("General",new ExcelNumberFormatXml(NameSpaceManager,true){NumFmtId=0,Format="General"});
            NumberFormats.Add("0", new ExcelNumberFormatXml(NameSpaceManager,true) { NumFmtId = 1, Format = "0" });
            NumberFormats.Add("0.00", new ExcelNumberFormatXml(NameSpaceManager,true) { NumFmtId = 2, Format = "0.00" });
            NumberFormats.Add("#,##0", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 3, Format = "#,##0" });
            NumberFormats.Add("#,##0.00", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 4, Format = "#,##0.00" });
            NumberFormats.Add("0%", new ExcelNumberFormatXml(NameSpaceManager,true) { NumFmtId = 9, Format = "0%" });
            NumberFormats.Add("0.00%", new ExcelNumberFormatXml(NameSpaceManager,true) { NumFmtId = 10, Format = "0.00%" });
            NumberFormats.Add("0.00E+00", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 11, Format = "0.00E+00" });
            NumberFormats.Add("# ?/?", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 12, Format = "# ?/?" });
            NumberFormats.Add("# ??/??", new ExcelNumberFormatXml(NameSpaceManager,true) { NumFmtId = 13, Format = "# ??/??" });
            NumberFormats.Add("mm-dd-yy", new ExcelNumberFormatXml(NameSpaceManager,true) { NumFmtId = 14, Format = "mm-dd-yy" });
            NumberFormats.Add("d-mmm-yy", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 15, Format = "d-mmm-yy" });
            NumberFormats.Add("d-mmm", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 16, Format = "d-mmm" });
            NumberFormats.Add("mmm-yy", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 17, Format = "mmm-yy" });
            NumberFormats.Add("h:mm AM/PM", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 18, Format = "h:mm AM/PM" });
            NumberFormats.Add("h:mm:ss AM/PM", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 19, Format = "h:mm:ss AM/PM" });
            NumberFormats.Add("h:mm", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 20, Format = "h:mm" });
            NumberFormats.Add("h:mm:dd", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 21, Format = "h:mm:dd" });
            NumberFormats.Add("m/d/yy h:mm", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 22, Format = "m/d/yy h:mm" });
            NumberFormats.Add("#,##0 ;(#,##0)", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 37, Format = "#,##0 ;(#,##0)" });
            NumberFormats.Add("#,##0 ;[Red](#,##0)", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 38, Format = "#,##0 ;[Red](#,##0)" });
            NumberFormats.Add("#,##0.00;(#,##0.00)", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 39, Format = "#,##0.00;(#,##0.00)" });
            NumberFormats.Add("#,##0.00;[Red](#,#)", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 40, Format = "#,##0.00;[Red](#,#)" });
            NumberFormats.Add("mm:ss", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 45, Format = "mm:ss" });
            NumberFormats.Add("[h]:mm:ss", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 46, Format = "[h]:mm:ss" });
            NumberFormats.Add("mmss.0", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 47, Format = "mmss.0" });
            NumberFormats.Add("##0.0", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 48, Format = "##0.0" });
            NumberFormats.Add("@", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 49, Format = "@" });

            NumberFormats.NextId = 164; //Start for custom formats.
        }

        internal override XmlNode CreateXmlNode(XmlNode topNode)
        {
            TopNode = topNode;
            SetXmlNode("@numFmtId", NumFmtId.ToString());
            SetXmlNode("@formatCode", Format);
            return TopNode;
        }
    }
}
