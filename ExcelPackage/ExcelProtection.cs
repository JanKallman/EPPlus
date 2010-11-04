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
 *  
 * Code change notes:
 * 
 * Author							Change						Date
 *******************************************************************************
 * Jan Källman		Added		10-AUG-2010
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
namespace OfficeOpenXml
{
    /// <summary>
    /// Sets protection on the workbook level
    ///<seealso cref="ExcelEncryption"/> 
    ///<seealso cref="ExcelSheetProtection"/> 
    /// </summary>
    public class ExcelProtection : XmlHelper
    {
        public ExcelProtection(XmlNamespaceManager ns, XmlNode topNode) :
            base(ns, topNode)
        {
        }
        const string workbookPasswordPath = "d:workbookProtection/@workbookPassword";
        /// <summary>
        /// Sets a password for the workbook. This does not encrypt the workbook. 
        /// </summary>
        /// <param name="Password">The password. </param>
        public void SetPassword(string Password)
        {
            if(string.IsNullOrEmpty(Password))
            {
                DeleteNode(workbookPasswordPath);
            }
            else
            {
                SetXmlNodeString(workbookPasswordPath, ((int)EncryptedPackageHandler.CalculatePasswordHash(Password)).ToString("x"));
            }
        }
        const string lockStructurePath = "d:workbookProtection/@lockStructure";
        /// <summary>
        /// Locks the structure,which prevents users from adding or deleting worksheets or from displaying hidden worksheets.
        /// </summary>
        public bool LockStructure
        {
            get
            {
                return GetXmlNodeBool(lockStructurePath, false);
            }
            set
            {
                SetXmlNodeBool(lockStructurePath, value,  false);
            }
        }
        const string lockWindowsPath = "d:workbookProtection/@lockWindows";
        /// <summary>
        /// Locks the position of the workbook window.
        /// </summary>
        public bool LockWindows
        {
            get
            {
                return GetXmlNodeBool(lockWindowsPath, false);
            }
            set
            {
                SetXmlNodeBool(lockWindowsPath, value, false);
            }
        }
        const string lockRevisionPath = "d:workbookProtection/@lockRevision";

        /// <summary>
        /// Lock the workbook for revision
        /// </summary>
        public bool LockRevision
        {
            get
            {
                return GetXmlNodeBool(lockRevisionPath, false);
            }
            set
            {
                SetXmlNodeBool(lockRevisionPath, value, false);
            }
        }
    }
}
