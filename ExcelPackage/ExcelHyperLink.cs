/* 
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * EPPlus provides server-side generation of Excel 2007 spreadsheets.
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
 * Jan Källman		                Added this class		        2010-01-24
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml
{
    /// <summary>
    /// HyperlinkClass
    /// </summary>
    public class ExcelHyperLink : Uri
    {
        /// <summary>
        /// A new hyperlink with the specified URI
        /// </summary>
        /// <param name="uriString">The URI</param>
        public ExcelHyperLink(string uriString) :
            base(uriString)
        {

        }
        /// <summary>
        /// A new hyperlink with the specified URI. This syntax is obsolete
        /// </summary>
        /// <param name="uriString">The URI</param>
        /// <param name="dontEscape"></param>
        public ExcelHyperLink(string uriString, bool dontEscape) :
            base(uriString, dontEscape)
        {

        }
        /// <summary>
        /// A new hyperlink with the specified URI and kind
        /// </summary>
        /// <param name="uriString">The URI</param>
        /// <param name="uriKind">Kind (absolute/relative or indeterminate)</param>
        public ExcelHyperLink(string uriString, UriKind uriKind) :
            base(uriString, uriKind)
        {

        }
        /// <summary>
        /// Sheet internal reference
        /// </summary>
        /// <param name="referenceAddress">Address</param>
        /// <param name="display">Displayed text</param>
        public ExcelHyperLink(string referenceAddress, string display) :
            base("xl://internal")   //URI is not used on internal links so put a dummy uri here.
        {
            _referenceAddress = referenceAddress;
            _display = display;
        }

        string _referenceAddress = null;
        public string ReferenceAddress
        {
            get
            {
                return _referenceAddress;
            }
            set
            {
                _referenceAddress = value;
            }
        }
        string _display = "";
        /// <summary>
        /// Displayed text
        /// </summary>
        public string Display
        {
            get
            {
                return _display;
            }
            set
            {
                _display = value;
            }
        }
        int _colSpann = 0;
        /// <summary>
        /// If the hyperlink spans multiple columns
        /// </summary>
        public int ColSpann
        {
            get
            {
                return _colSpann;
            }
            set
            {
                _colSpann = value;
            }
        }
        int _rowSpann = 0;
        /// <summary>
        /// If the hyperlink spans multiple rows
        /// </summary>
        public int RowSpann
        {
            get
            {
                return _rowSpann;
            }
            set
            {
                _rowSpann = value;
            }
        }
    }
}
