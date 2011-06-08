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
 *
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

using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security;

// General Information about an assembly is controlled through the following 
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly: AssemblyTitle("EPPlus")]
[assembly: AssemblyDescription("Allows Excel 2007 files to be created on the server. See epplus.codeplex.com")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("")]
[assembly: AssemblyProduct("EPPlus")]
[assembly: AssemblyCopyright("Copyright 2009- ©Jan Källman. Parts of the Interface comes from ExcelPackage-project(© Dr John Tunnicliffe dr.john.tunnicliffe@btinternet.com)")]
[assembly: AssemblyTrademark("The GNU General Public License2 (GPL2)")]
[assembly: AssemblyCulture("")]
[assembly: ComVisible(false)]
[assembly: InternalsVisibleTo("EPPlusTest, PublicKey=00240000048000009400000006020000002400005253413100040000010001001dd11308ec93a6ebcec727e183a8972dc6f95c23ecc34aa04f40cbfc9c17b08b4a0ea5c00dcd203bace44d15a30ce8796e38176ae88e960ceff9cc439ab938738ba0e603e3d155fc298799b391c004fc0eb4393dd254ce25db341eb43303e4c488c9500e126f1288594f0710ec7d642e9c72e76dd860649f1c48249c00e31fba")]

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid("9dd43b8d-c4fe-4a8b-ad6e-47ef83bbbb01")]

// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version 
//      Build Number
//      Revision
//
// You can specify all the values or you can default the Revision and Build Numbers 
// by using the '*' as shown below:
[assembly: AssemblyVersion("2.9.0.1")]
[assembly: AssemblyFileVersion("2.9.0.1")]
[assembly: AllowPartiallyTrustedCallers]