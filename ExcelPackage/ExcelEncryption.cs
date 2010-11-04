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
 ***************************************************************************************
 * This class is created with the help of the MS-OFFCRYPTO PDF documentation... http://msdn.microsoft.com/en-us/library/cc313071(office.12).aspx
 * Decrypytion library for Office Open XML files(Lyquidity) and Sminks very nice example 
 * on "Reading compound documents in c#" on Stackoverflow. Many thanks!
 ***************************************************************************************
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

namespace OfficeOpenXml
{
    /// <summary>
    /// Encryption Algorithm
    /// </summary>
    public enum EncryptionAlgorithm
    {
        /// <summary>
        /// // 128-bit AES. Default
        /// </summary>
        AES128,
        /// <summary>
        /// // 192-bit AES.
        /// </summary>
        AES192,
        /// <summary>
        /// // 256-bit AES. 
        /// </summary>
        AES256
    }
    /// <summary>
    /// How and if the workbook is encrypted
    ///<seealso cref="ExcelProtection"/> 
    ///<seealso cref="ExcelSheetProtection"/> 
    /// </summary>
    public class ExcelEncryption
    {
        /// <summary>
        /// Constructor
        /// </summary>
        public ExcelEncryption()
        {
            Algorithm = EncryptionAlgorithm.AES128;
        }
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="encryptionAlgorithm">Algorithm used to encrypt the package. Default is AES128</param>
        public ExcelEncryption(EncryptionAlgorithm encryptionAlgorithm)
        {
            Algorithm = encryptionAlgorithm;
        }        
        bool _isEncrypted = false;
        /// <summary>
        /// Is the package encrypted
        /// </summary>
        public bool IsEncrypted 
        {
            get
            {
                return _isEncrypted;
            }
            set
            {
                _isEncrypted = value;
                if (_isEncrypted)
                {
                    if (_password == null) _password = "";
                }
                else
                {
                    _password = null;
                }
            }
        }
        string _password=null;
        /// <summary>
        /// The password used to encrypt the workbook.
        /// </summary>
        public string Password 
        {
            get
            {
                return _password;
            }
            set
            {
                _password = value;
                _isEncrypted = (value != null);
            }
        }
        /// <summary>
        /// Algorithm used for encrypting the package. Default is AES 128-bit
        /// </summary>
        public EncryptionAlgorithm Algorithm { get; set; }
    }
}
