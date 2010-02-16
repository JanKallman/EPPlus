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

/* 
* You may amend and distribute as you like, but don't remove this header!
* 
* ExcelPackage provides server-side generation of Excel 2007 spreadsheets.
* See http://www.codeplex.com/ExcelPackage for details.
* 
* Copyright 2007 © Dr John Tunnicliffe 
* mailto:dr.john.tunnicliffe@btinternet.com
* All rights reserved.
* 
* ExcelPackage is an Open Source project provided under the 
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
*/

/*
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * John Tunnicliffe		Initial Release		01-Jan-2007
 * ******************************************************************************
 */
using System;
using System.Xml;
using System.IO;
using System.IO.Packaging;

namespace OfficeOpenXml
{
    /// <summary>
	/// Represents an Excel 2007 XLSX file package.  Opens the file and provides access
	/// to all the components (workbook, worksheets, properties etc.).
	/// </summary>
	public class ExcelPackage : IDisposable
	{
        internal const bool preserveWhitespace=true;
		#region Properties
		/// <summary>
		/// Provides access to the main schema used by all Excel components
		/// </summary>
		protected internal const string schemaMain = @"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
		/// <summary>
		/// Provides access to the relationship schema
		/// </summary>
		protected internal const string schemaRelationships = @"http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        protected internal const string schemaDrawings = @"http://schemas.openxmlformats.org/drawingml/2006/main";
        protected internal const string schemaSheetDrawings = @"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

        protected internal const string schemaChart = @"http://schemas.openxmlformats.org/drawingml/2006/chart";                                                        
        protected internal const string schemaHyperlink = @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
                                                           
        private Package _package;
		private string _outputFolderPath;

		private ExcelWorkbook _workbook;
        public const int MaxColumns = 16384;
        public const int MaxRows = 16777216;
		#endregion

		#region ExcelPackage Constructors
        /// <summary>
        /// Creates a new instance of the ExcelPackage. Output is accessed through the Stream property.
        /// </summary>
        public ExcelPackage()
        {
            ConstructNewFile();
        }
        /// <summary>
		/// Creates a new instance of the ExcelPackage class based on a existing file or creates a new file. 
		/// </summary>
		/// <param name="newFile">If newFile exists, it is opened.  Otherwise it is created from scratch.</param>
        public ExcelPackage(FileInfo newFile)
		{
            File = newFile;
            ConstructNewFile();
        }

		/// <summary>
		/// Creates a new instance of the ExcelPackage class based on a existing template.
		/// WARNING: If newFile exists, it is deleted!
		/// </summary>
		/// <param name="newFile">The name of the Excel file to be created</param>
		/// <param name="template">The name of the Excel template to use as the basis of the new Excel file</param>
		public ExcelPackage(FileInfo newFile, FileInfo template)
		{
            File = newFile;
            CreateFromTemplate(template);
		}   

        /// <summary>
        /// Create a new file frp, a template
        /// </summary>
        /// <param name="newFile"></param>
        /// <param name="template"></param>
        /// <returns></returns>
        private void CreateFromTemplate(FileInfo template)
        {
            if (template.Exists)
            {
                _stream = new MemoryStream();
                byte[] b = System.IO.File.ReadAllBytes(template.FullName);
                _stream.Write(b, 0, b.Length);
                _package = Package.Open(_stream, FileMode.Open, FileAccess.ReadWrite);
            }
            else
                throw new Exception("ExcelPackage Error: Passed invalid TemplatePath to Excel Template");
            //return newFile;
        }
        
        private void ConstructNewFile()
        {
            _stream = new MemoryStream();
            if (File!=null && File.Exists)
            {
                byte[] b = System.IO.File.ReadAllBytes(File.FullName);
                _stream.Write(b, 0, b.Length);
                _package = Package.Open(_stream, FileMode.Open, FileAccess.ReadWrite);
            }
            else
            {
                _package = Package.Open(_stream, FileMode.Create, FileAccess.ReadWrite);
                CreateBlankWb();
            }   
            //else
            //{
            //    _outputFolderPath = newFile.DirectoryName;
            //    if (newFile.Exists)
            //        // open the existing package
            //        _package = Package.Open(newFile.FullName, FileMode.Open, FileAccess.ReadWrite);
            //    else
            //    {
            //        // create a new package and add the main workbook.xml part
            //        _package = Package.Open(newFile.FullName, FileMode.Create, FileAccess.ReadWrite);

            //        CreateBlankWb();
            //    }
            //}
        }

        private void CreateBlankWb()
        {
            // save a temporary part to create the default application/xml content type
            Uri uriDefaultContentType = new Uri("/default.xml", UriKind.Relative);
            PackagePart partTemp = _package.CreatePart(uriDefaultContentType, "application/xml", CompressionOption.Maximum);

            XmlDocument workbook = Workbook.WorkbookXml; // this will create the workbook xml in the package

            // create the relationship to the main part
            _package.CreateRelationship(Workbook.WorkbookUri, TargetMode.Internal, schemaRelationships + "/officeDocument");

            // remove the temporary part that created the default xml content type
            _package.DeletePart(uriDefaultContentType);
        }

        MemoryStream _stream=null;
        /// <summary>
        /// Creates a new instance of the ExcelPackage class based on a existing template.
        /// </summary>
        /// <param name="template">The name of the Excel template to use as the basis of the new Excel file</param>
        /// <param name="useStream">if true use a strem. If false create a file in the temp dir with a random name</param>
        public ExcelPackage(FileInfo template, bool useStream)
        {
            CreateFromTemplate(template);
            if (useStream == false)
            {
                File = new FileInfo(Path.GetTempPath() + Guid.NewGuid().ToString() + ".xlsx");
            }
        }
        #endregion

		#region Public Properties
		/// <summary>
		/// Setting DebugMode to true will cause the Save method to write the 
		/// raw XML components to the same folder as the output Excel file
		/// </summary>
		public bool DebugMode = false;

		/// <summary>
		/// Returns a reference to the file package
		/// </summary>
		public Package Package { get { return (_package); } }

		/// <summary>
		/// Returns a reference to the workbook component within the package.
		/// All worksheets and cells can be accessed through the workbook.
		/// </summary>
		public ExcelWorkbook Workbook
		{
			get
			{
                if (_workbook == null)
                {
                    //  Create a NamespaceManager to handle the default namespace, 
                    //  and create a prefix for the default namespace:
                    NameTable nt = new NameTable();
                    var ns = new XmlNamespaceManager(nt);
                    ns.AddNamespace("d", ExcelPackage.schemaMain);
                    _workbook = new ExcelWorkbook(this, ns);

                    _workbook.GetDefinedNames();

                }
                return (_workbook);
			}
		}
		#endregion
		
		#region WriteDebugFile
		/// <summary>
		/// Writes a debug file to the output folder, but only if DebugMode = true
		/// </summary>
		/// <param name="XmlDoc">The XmlDocument to save to the file system</param>
		/// <param name="subFolder">The subfolder in which the file is to be saved</param>
		/// <param name="FileName">The name of the file to save.</param>
		protected internal void WriteDebugFile(XmlDocument XmlDoc, string subFolder, string FileName)
		{
			if (DebugMode)
			{
				DirectoryInfo dir = new DirectoryInfo(_outputFolderPath + "/" + subFolder);
				if (!dir.Exists)
					dir.Create();

				FileInfo file = new FileInfo(_outputFolderPath + "/" + subFolder + "/" + FileName);
				if (file.Exists)
				{
					file.IsReadOnly = false;
					file.Delete();
				}
				XmlDoc.Save(file.FullName);
			}
		}
		#endregion

		
		///// <summary>
		///// Returns the Uri to a parent part (e.g. workbook.xml) 
		///// </summary>
		///// <param name="Relationship">The relationship the </param>
		///// <returns></returns>
		//protected internal Uri GetMainUri(string Relationship)
		//{
		//  Uri uriMain = null;
		//  //  Get the Uri to the main part
		//  Uri uriParent = new Uri("/", UriKind.Relative);
		//  PackageRelationship relationship = GetMainRelationship(Relationship);
		//  if (relationship != null)
		//    uriMain = PackUriHelper.ResolvePartUri(uriParent, relationship.TargetUri);
		//  return (uriMain);
		//}

		///// <summary>
		///// 
		///// </summary>
		///// <param name="Relationship"></param>
		///// <returns></returns>
		//protected internal PackageRelationship GetMainRelationship(string Relationship)
		//{
		//  PackageRelationship relMain = null;
		//  foreach (PackageRelationship relationship in _package.GetRelationshipsByType(schemaRelationships + "/" + Relationship))
		//  {
		//    relMain = relationship;
		//    break;  //  There should only be one main part
		//  }
		//  return (relMain);
		//}

		#region GetSharedUri
		/// <summary>
		/// Obtains the Uri to a shared part (e.g. sharedstrings.xml)
		/// </summary>
		/// <param name="uriParent">Uri to the parent component</param>
		/// <param name="Relationship">The relationship to the parent component</param>
		/// <returns>The Uri to a shared part</returns>
		protected internal Uri GetSharedUri(Uri uriParent, string Relationship)
		{
			Uri uriShared = null;
			PackagePart partParent = _package.GetPart(uriParent);
			//  Get the Uri to the shared part
			foreach (System.IO.Packaging.PackageRelationship relationship in partParent.GetRelationshipsByType(schemaRelationships + "/" + Relationship))
			{
				uriShared = PackUriHelper.ResolvePartUri(uriParent, relationship.TargetUri);
				break;  //  There should only be one shared resource
			}
			return (uriShared);
		}
		#endregion

		#region AddSchemaAttribute
		/// <summary>
		/// Adds additional schema attributes to the root element
		/// </summary>
		/// <param name="root">The root element</param>
		/// <param name="nameSpace">The namespace of the schema</param>
		/// <param name="schema">The schema to apply</param>
		protected internal static void AddSchemaAttribute(XmlElement root, string schema, string nameSpace)
		{
			XmlAttribute nsAttribute = root.OwnerDocument.CreateAttribute("xmlns", nameSpace, @"http://www.w3.org/2000/xmlns/");
			nsAttribute.Value = schema;
			root.Attributes.Append(nsAttribute);
		}

		/// <summary>
		/// Adds additional schema attributes to the root element
		/// </summary>
		/// <param name="root">The root element</param>
		/// <param name="schema">The schema to apply</param>
		protected internal static void AddSchemaAttribute(XmlElement root, string schema)
		{
			XmlAttribute nsAttribute = root.OwnerDocument.CreateAttribute("xmlns");
			nsAttribute.Value = schema;
			root.Attributes.Append(nsAttribute);
		}
		#endregion

		#region SavePart
		/// <summary>
		/// Saves the XmlDocument into the package at the specified Uri.
		/// </summary>
		/// <param name="uriPart">The Uri of the component</param>
		/// <param name="xmlPart">The XmlDocument to save</param>
		protected internal void SavePart(Uri uriPart, XmlDocument xmlPart)
		{
			PackagePart partPack = _package.GetPart(uriPart);
			xmlPart.Save(partPack.GetStream(FileMode.Create, FileAccess.Write));
		}
		#endregion

		#region Dispose
		/// <summary>
		/// Closes the package.
		/// </summary>
		public void Dispose()
		{
			_package.Close();
		}
		#endregion

		#region Save  // ExcelPackage save
		/// <summary>
		/// Saves all the components back into the package.
		/// This method recursively calls the Save method on all sub-components.
        /// We close the package after the save is done.
		/// </summary>
		public void Save()
		{
            try
            {
                Workbook.Save();
                if (File != null)
                {                    
                    if (System.IO.File.Exists(File.FullName))
                    {
                        try
                        {
                            System.IO.File.Delete(File.FullName);
                        }
                        catch (Exception ex)
                        {
                            throw(new Exception( string.Format("Error overwriting file {0}", File.FullName), ex));
                        }
                    }
                    if (Stream is MemoryStream)
                    {
                        _package.Close();
                        var fi = new FileStream(File.FullName, FileMode.Create);
                        fi.Write(((MemoryStream)Stream).GetBuffer(), 0, (int)Stream.Length);
                        fi.Close();
                    }
                    else
                    {
                        System.IO.File.WriteAllBytes(File.FullName, GetAsByteArray(false));
                    }
                }
                _package.Close();
            }
            catch(Exception ex)
            {
                throw (new Exception(string.Format("Error saving file {0}"), ex));
            }
        }
        /// <summary>
        /// Saves the workbook to a new file
        /// Package is closed after it has been saved
        /// </summary>
        public void SaveAs(FileInfo file)
        {
            File = file;
            Save();
        }
        FileInfo _file=null;
        public FileInfo File
        {
            get
            {
                return _file;
            }
            set
            {
                _file = value;
                _outputFolderPath = _file.DirectoryName;
            }
        }
        public Stream Stream
        {
            get
            {
                return _stream;
            }
        }
		#endregion

		#region GetXmlFromUri
		/// <summary>
		/// Obtains the XmlDocument from the package referenced by the Uri
		/// </summary>
		/// <param name="uriPart">The Uri to the component</param>
		/// <returns>The XmlDocument of the component</returns>
		protected internal XmlDocument GetXmlFromUri(Uri uriPart)
		{
			XmlDocument xlPart = new XmlDocument();
			PackagePart packPart = _package.GetPart(uriPart);
			xlPart.Load(packPart.GetStream());
			return (xlPart);
		}
		#endregion

        /// <summary>
        /// Saves and returns the Excel files as a bytearray
        /// Can only be used when working with a stream. That is .. new ExcelPackage() or new ExcelPackage("file", true)
        /// </summary>
        /// <example>      
        /// Example how to return a document from a Webserver...
        /// <code> 
        ///  ExcelPackage package=new ExcelPackage();
        ///  /**** ... Create the document ****/
        ///  Byte[] bin = package.GetAsByteArray();
        ///  Response.ContentType = "Application/vnd.ms-Excel";
        ///  Response.AddHeader("content-disposition", "attachment;  filename=TheFile.xlsx");
		///  Response.BinaryWrite(bin);
        /// </code>
        /// </example>
        /// <returns></returns>
        public byte[] GetAsByteArray()
        {
           return GetAsByteArray(true);
        }
        internal  byte[] GetAsByteArray(bool save)
        {
            if(save) Workbook.Save();
            _package.Close();
            Byte[] byRet = new byte[Stream.Length];
            long pos = Stream.Position;            
            Stream.Seek(0, SeekOrigin.Begin);
            Stream.Read(byRet, 0, (int)Stream.Length);

            Stream.Seek(pos, SeekOrigin.Begin);
            Stream.Close();
            return byRet;
        }

    }
}