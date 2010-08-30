/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * All rights reserved.
 * 
 * EPPlus is an Open Source project provided under the 
 * GNU General Public License (GPL) as published by the 
 * Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
 * 
 * See http://epplus.codeplex.com/ for details
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
 * Code change notes:
 * 
 * Author							Change						Date
 *******************************************************************************
 * Jan Källman		Added		30-AUG-2010
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Table
{
    /// <summary>
    /// Table style Enum
    /// </summary>
    public enum TableStyles
    {
        Custom,
        Light1,
        Light2,
        Light3,
        Light4,
        Light5,
        Light6,
        Light7,
        Light8,
        Light9,
        Light10,
        Light11,
        Light12,
        Light13,
        Light14,
        Light15,
        Light16,
        Light17,
        Light18,
        Light19,
        Light20,
        Light21,
        Medium1,
        Medium2,
        Medium3,
        Medium4,
        Medium5,
        Medium6,
        Medium7,
        Medium8,
        Medium9,
        Medium10,
        Medium11,
        Medium12,
        Medium13,
        Medium14,
        Medium15,
        Medium16,
        Medium17,
        Medium18,
        Medium19,
        Medium20,
        Medium21,
        Medium22,
        Medium23,
        Medium24,
        Medium25,
        Medium26,
        Medium27,
        Medium28,
        Dark1,
        Dark2,
        Dark3,
        Dark4,
        Dark5,
        Dark6,
        Dark7,
        Dark8,
        Dark9,
        Dark10,
        Dark11,
        //PivotStyleLight1,
        //PivotStyleLight2,
        //PivotStyleLight3,
        //PivotStyleLight4,
        //PivotStyleLight5,
        //PivotStyleLight6,
        //PivotStyleLight7,
        //PivotStyleLight8,
        //PivotStyleLight9,
        //PivotStyleLight10,
        //PivotStyleLight11,
        //PivotStyleLight12,
        //PivotStyleLight13,
        //PivotStyleLight14,
        //PivotStyleLight15,
        //PivotStyleLight16,
        //PivotStyleLight17,
        //PivotStyleLight18,
        //PivotStyleLight19,
        //PivotStyleLight20,
        //PivotStyleLight21,
        //PivotStyleMedium1,
        //PivotStyleMedium2,
        //PivotStyleMedium3,
        //PivotStyleMedium4,
        //PivotStyleMedium5,
        //PivotStyleMedium6,
        //PivotStyleMedium7,
        //PivotStyleMedium8,
        //PivotStyleMedium9,
        //PivotStyleMedium10,
        //PivotStyleMedium11,
        //PivotStyleMedium12,
        //PivotStyleMedium13,
        //PivotStyleMedium14,
        //PivotStyleMedium15,
        //PivotStyleMedium16,
        //PivotStyleMedium17,
        //PivotStyleMedium18,
        //PivotStyleMedium19,
        //PivotStyleMedium20,
        //PivotStyleMedium21,
        //PivotStyleMedium22,
        //PivotStyleMedium23,
        //PivotStyleMedium24,
        //PivotStyleMedium25,
        //PivotStyleMedium26,
        //PivotStyleMedium27,
        //PivotStyleMedium28
        //PivotStyleDark1,
        //PivotStyleDark2,
        //PivotStyleDark3,
        //PivotStyleDark4,
        //PivotStyleDark5,
        //PivotStyleDark6,
        //PivotStyleDark7,
        //PivotStyleDark8,
        //PivotStyleDark9,
        //PivotStyleDark10,
        //PivotStyleDark11    
    }
    public class ExcelTable : XmlHelper
    {
        public ExcelTable (XmlNamespaceManager ns, ExcelWorksheet sheet, ExcelAddressBase address, string name) : base(ns)
	    {
            WorkSheet = sheet;
            Address = address;            
            TableXml = new XmlDocument();
            TableXml.LoadXml(GetStartXml(name));    
            TopNode = TableXml.DocumentElement;
        }

        private string GetStartXml(string name)
        {
            string xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?>";
            xml += string.Format("<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"1\" name=\"{0}\" displayName=\"{0}\" ref=\"{1}\" totalsRowShown=\"0\">", name, Address.Address);
            xml += string.Format("<autoFilter ref=\"{0}\" />", Address.Address);

            int cols=Address._toCol-Address._fromCol;
            xml += string.Format("<tableColumns count=\"0\">",cols);
            for(int i=0;i<cols;i++)
            {
                xml += string.Format("<tableColumn id=\"{0}\" name=\"Column{0}\" />", i);
            }
            xml += "</tableColumns>";
            xml += "<tableStyleInfo name=\"TableStyleMedium6\" showFirstColumn=\"0\" showLastColumn=\"0\" showRowStripes=\"1\" showColumnStripes=\"0\" /> ";
            xml += "</table>";

            return xml;
        }
        /// <summary>
        /// The Xml document
        /// </summary>
        public XmlDocument TableXml
        {
            get;
            set;
        }
        const string NAME_PATH="d:table/@name";
        const string DISPLAY_NAME_PATH = "d:table/@displayName";
        /// <summary>
        /// Name of the table object in Excel
        /// </summary>
        public string Name
        {
            get
            {
                return GetXmlNodeString(NAME_PATH);
            }
            set
            {
                SetXmlNodeString(NAME_PATH, value);
                SetXmlNodeString(DISPLAY_NAME_PATH, value);
            }
        }
        /// <summary>
        /// The worksheet of the table
        /// </summary>
        public ExcelWorksheet WorkSheet
        {
            get;
            set;
        }
        /// <summary>
        /// The address of the table
        /// </summary>
        public ExcelAddressBase Address
        {
            get;
            set;
        }
        TableStyles _tableStyle = TableStyles.Medium6;
        /// <summary>
        /// The table style. If this property is cusom the style from the StyleName propery is used.
        /// </summary>
        public TableStyles Style
        {
            get
            {
                return _tableStyle;
            }
            set
            {
                _tableStyle=value;
                if (value != TableStyles.Custom)
                {
                    SetXmlNodeString(StyleName, "TableStyle" + value.ToString());
                }
            }
        }
        const string HEADERROWCOUNT_PATH = "d:table/@headerRowCount";
        public int HeaderRowCount
        {
            get
            {
                return GetXmlNodeInt(HEADERROWCOUNT_PATH);
            }
            set
            {
                SetXmlNodeString(HEADERROWCOUNT_PATH, value.ToString());
            }
        }
        const string TOTALSROWCOUNT_PATH = "d:table/@totalsRowCount";
        public int TotalsRowCount
        {
            get
            {
                return GetXmlNodeInt(TOTALSROWCOUNT_PATH);
            }
            set
            {
                SetXmlNodeString(TOTALSROWCOUNT_PATH, value.ToString());
            }
        }
        const string STYLENAME_PATH = "d:table/d:tableStyleInfo/@name";
        public string StyleName
        {
            get
            {
                return GetXmlNodeString(StyleName);
            }
            set
            {
                if (value.StartsWith("TableStyle"))
                {
                    try
                    {
                        _tableStyle = (TableStyles)Enum.Parse(typeof(TableStyles), value.Substring(10,value.Length-10), true);
                    }
                    catch
                    {
                        _tableStyle = TableStyles.Custom;
                    }
                }
                else
                {
                    _tableStyle = TableStyles.Custom;
                }
                SetXmlNodeString(StyleName,value);
            }
        }
        const string SHOWFIRSTCOLUMN_PATH = "d:table/d:tableStyleInfo/@showFirstColumn";
        public bool ShowFirstColumn
        {
            get
            {
                return GetXmlNodeBool(SHOWFIRSTCOLUMN_PATH);
            }
            set
            {
                SetXmlNodeBool(SHOWFIRSTCOLUMN_PATH, value, false);
            }   
        }
        const string SHOWLASTCOLUMN_PATH = "d:table/d:tableStyleInfo/@showLastColumn";
        public bool ShowLastColumn
        {
            get
            {
                return GetXmlNodeBool(SHOWLASTCOLUMN_PATH);
            }
            set
            {
                SetXmlNodeBool(SHOWLASTCOLUMN_PATH, value, false);
            }
        }
        const string SHOWROWSTRIPES_PATH = "d:table/d:tableStyleInfo/@showRowStripes";
        public bool ShowRowStripes
        {
            get
            {
                return GetXmlNodeBool(SHOWROWSTRIPES_PATH);
            }
            set
            {
                SetXmlNodeBool(SHOWROWSTRIPES_PATH, value, false);
            }
        }
        const string SHOWCOLUMNSTRIPES_PATH = "d:table/d:tableStyleInfo/@showColumnStripes";
        public bool ShowColumnStripes
        {
            get
            {
                return GetXmlNodeBool(SHOWCOLUMNSTRIPES_PATH);
            }
            set
            {
                SetXmlNodeBool(SHOWCOLUMNSTRIPES_PATH, value, false);
            }
        }
    }
}
