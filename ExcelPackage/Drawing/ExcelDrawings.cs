/* 
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * EPPlus provides server-side generation of Excel 2007 spreadsheets.
 * EPPlus is a fork of the ExcelPackage project
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
using System.IO.Packaging;
using System.Xml;
using System.Collections;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Enables access to the Drawings.
    /// VERY basic support for Charts, Shapes and Pictures in present version.
    /// </summary>
    public class ExcelDrawings : IEnumerable
    {
        private XmlDocument _drawingsXml=new XmlDocument();
        private Dictionary<int, ExcelDrawing> _drawings;
        public ExcelDrawings(ExcelPackage xlPackage, ExcelWorksheet sheet)
        {
                _drawingsXml = new XmlDocument();
                _drawingsXml.PreserveWhitespace = true;
                _drawings = new Dictionary<int, ExcelDrawing>();

                XmlNode node = sheet.WorksheetXml.SelectSingleNode("//d:drawing", sheet.NameSpaceManager);
                CreateNSM();
                if (node != null)
                {
                    PackageRelationship drawingRelation = sheet.Part.GetRelationship(node.Attributes["r:id"].Value);
                    _uriDrawing = PackUriHelper.ResolvePartUri(sheet.WorksheetUri, drawingRelation.TargetUri);

                    _part = xlPackage.Package.GetPart(_uriDrawing);
                    _drawingsXml.Load(_part.GetStream());

                    AddDrawings();
                }
         }
        public XmlDocument DrawingXml
        {
            get
            {
                return _drawingsXml;
            }
        }
        private void AddDrawings()
        {
            XmlNodeList list = _drawingsXml.SelectNodes("//xdr:twoCellAnchor", NameSpaceManager);

            int i = 0;
            foreach (XmlNode node in list)
            {
                ExcelDrawing dr = ExcelDrawing.GetDrawing(this, node);
                _drawings.Add(++i, dr);
            }
        }


        #region NamespaceManager
        /// <summary>
        /// Creates the NamespaceManager. 
        /// </summary>
        private void CreateNSM()
        {
            NameTable nt = new NameTable();
            _nsManager = new XmlNamespaceManager(nt);
            _nsManager.AddNamespace("a", ExcelPackage.schemaDrawings);
            _nsManager.AddNamespace("xdr", ExcelPackage.schemaSheetDrawings);
            _nsManager.AddNamespace("c", ExcelPackage.schemaChart);
            _nsManager.AddNamespace("r", ExcelPackage.schemaRelationships);
        }
        /// <summary>
        /// Provides access to a namespace manager instance to allow XPath searching
        /// </summary>
        XmlNamespaceManager _nsManager=null;
        public XmlNamespaceManager NameSpaceManager
        {
            get
            {
                return _nsManager;
            }
        }
        #endregion

        #region IEnumerable Members

        public IEnumerator GetEnumerator()
        {
            return (_drawings.Values.GetEnumerator());
        }
        /// <summary>
        /// Returns the worksheet at the specified position.  
        /// </summary>
        /// <param name="PositionID">The position of the worksheet. 1-base</param>
        /// <returns></returns>
        public ExcelDrawing this[int PositionID]
        {
            get
            {
                return (_drawings[PositionID]);
            }
        }

        /// <summary>
        /// Returns the worksheet matching the specified name
        /// </summary>
        /// <param name="Name">The name of the worksheet</param>
        /// <returns></returns>
        public ExcelDrawing this[string Name]
        {
            get
            {
                foreach (ExcelDrawing drawing in _drawings.Values)
                {
                    if (drawing.Name == Name)
                        return drawing;
                }
                return null;
                //throw new Exception(string.Format("ExcelWorksheets Error: Worksheet '{0}' not found!",Name));
            }
        }
        public int Count
        {
            get
            {
                if (_drawings == null)
                {
                    return 0;
                }
                else
                {
                    return _drawings.Count;
                }
            }
        }
        PackagePart _part=null;
        internal PackagePart Part
        {
            get
            {
                return _part;
            }        
        }
        Uri _uriDrawing=null;
        public Uri UriDrawing
        {
            get
            {
                return _uriDrawing;
            }
        }
        #endregion

    }
}
