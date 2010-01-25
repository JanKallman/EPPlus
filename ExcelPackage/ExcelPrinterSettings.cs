/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * All rights reserved.
 * 
 * EPPlus is an Open Source project provided under the 
 * GNU General Public License (GPL) as published by the 
 * Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
 * 
 * EPPlus provides server-side generation of Excel 2007 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 * EPPlus is a fork of the ExcelPackage project
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
 *******************************************************************************
 * Jan Källman		Added		10-SEP-2009
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml
{
    public enum eOrientation
    {
        Portrait,
        Landscape
    }
    /// <summary>
    /// Printer settings
    /// </summary>
    public class ExcelPrinterSettings : XmlHelper
    {
        bool _marginsCreated = false;
        public ExcelPrinterSettings(XmlNamespaceManager ns) :
            base(ns)
        {

        }
        public ExcelPrinterSettings(XmlNamespaceManager ns, XmlNode topNode) :
            base(ns, topNode)
        {

        }
        const string _leftMarginPath = "d:pageMargins/@left";
        /// <summary>
        /// Left margin
        /// </summary>
        public decimal LeftMargin 
        {
            get
            {
                return GetXmlNodeDecimal(_leftMarginPath);
            }
            set
            {
               CreateMargins();
               SetXmlNode(_leftMarginPath, value.ToString());
            }
        }
        const string _rightMarginPath = "d:pageMargins/@right";
        /// <summary>
        /// Right margin
        /// </summary>
        public decimal RightMargin 
        {
            get
            {
                return GetXmlNodeDecimal(_rightMarginPath);
            }
            set
            {
                CreateMargins();
                SetXmlNode(_rightMarginPath, value.ToString());
            }
        }
        const string _topMarginPath = "d:pageMargins/@top";
        /// <summary>
        /// Top margin
        /// </summary>
        public decimal TopMargin 
        {
            get
            {
                return GetXmlNodeDecimal(_topMarginPath);
            }
            set
            {
                CreateMargins();
                SetXmlNode(_topMarginPath, value.ToString());
            }
        }
        const string _bottomMarginPath = "d:pageMargins/@bottom";
        /// <summary>
        /// Bottom margin
        /// </summary>
        public decimal BottomMargin 
        {
            get
            {
                return GetXmlNodeDecimal(_bottomMarginPath);
            }
            set
            {
                CreateMargins();
                SetXmlNode(_bottomMarginPath, value.ToString());
            }
        }
        const string _headerMarginPath = "d:pageMargins/@header";
        /// <summary>
        /// Header margin
        /// </summary>
        public decimal HeaderMargin 
        {
            get
            {
                return GetXmlNodeDecimal(_headerMarginPath);
            }
            set
            {
                CreateMargins();
                SetXmlNode(_headerMarginPath, value.ToString());
            }
        }
        const string _footerMarginPath = "d:pageMargins/@footer";
        /// <summary>
        /// Footer margin
        /// </summary>
        public decimal FooterMargin 
        {
            get
            {
                return GetXmlNodeDecimal(_footerMarginPath);
            }
            set
            {
                CreateMargins();
                SetXmlNode(_footerMarginPath, value.ToString());
            }
        }
        const string _orientationPath = "d:pageSetup/@orientation";
        /// <summary>
        /// Orientation 
        /// Portrait or Landscape
        /// </summary>
        public eOrientation Orientation
        {
            get
            {
                return (eOrientation)Enum.Parse(typeof(eOrientation), GetXmlNode(_orientationPath), true);
            }
            set
            {
                SetXmlNode(_orientationPath, value.ToString().ToLower());
            }
        }
        const string _fitToWidthPath = "d:pageSetup/@fitToWidth";
        /// <summary>
        /// Fit to Width in pages. 
        /// Set FitToPage to true when using this one. 
        /// 0 is automatic
        /// </summary>
        public int FitToWidth
        {
            get
            {
                return GetXmlNodeInt(_fitToWidthPath);
            }
            set
            {
                SetXmlNode(_fitToWidthPath, value.ToString());
            }
        }
        const string _fitToHeightPath = "d:pageSetup/@fitToHeight";
        /// <summary>
        /// Fit to height in pages. 
        /// Set FitToPage to true when using this one. 
        /// 0 is automatic
        /// </summary>
        public int FitToHeight
        {
            get
            {
                return GetXmlNodeInt(_fitToHeightPath);
            }
            set
            {
                SetXmlNode(_fitToHeightPath, value.ToString());
            }
        }
        const string _scalePath = "d:pageSetup/@scale";
        /// <summary>
        /// Print scale
        /// </summary>
        public int Scale
        {
            get
            {
                return GetXmlNodeInt(_scalePath);
            }
            set
            {
                SetXmlNode(_scalePath, value.ToString());
            }
        }
        const string _fitToPagePath = "d:sheetPr/d:pageSetUpPr/@fitToPage";
        /// <summary>
        /// Fit To Page.
        /// </summary>
        public bool FitToPage
        {
            get
            {
                return GetXmlNodeBool(_fitToPagePath);
            }
            set
            {
                SetXmlNode(_fitToPagePath, value ? "1" : "0");
            }
        }
        /// <summary>
        /// All or none of the margin attributes must exist. Create all att ones.
        /// </summary>
        private void CreateMargins()
        {
            if (_marginsCreated==false && TopNode.SelectSingleNode(_leftMarginPath, NameSpaceManager) == null) 
            {
                _marginsCreated=true;
                LeftMargin = 0.7087M;
                RightMargin = 0.7087M;
                TopMargin = 0.7480M;
                BottomMargin = 0.7480M;
                HeaderMargin = 0.315M;
                FooterMargin = 0.315M;
            }
        }
    }
}
