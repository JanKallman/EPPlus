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
 * Starnuto Di Topo & Jan Källman   Added stream constructors 
 *                                  and Load method Save as 
 *                                  stream                      2010-03-14
 * *******************************************************************************/
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
        internal const bool preserveWhitespace=false;
        Stream _stream = null;
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

        //Office properties
        protected internal const string schemaCore = @"http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        protected internal const string schemaExtended = @"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
        protected internal const string schemaCustom = @"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
        protected internal const string schemaDc = @"http://purl.org/dc/elements/1.1/";
        protected internal const string schemaDcTerms = @"http://purl.org/dc/terms/";
        protected internal const string schemaDcmiType = @"http://purl.org/dc/dcmitype/";
        protected internal const string schemaXsi = @"http://www.w3.org/2001/XMLSchema-instance";
        protected internal const string schemaVt = @"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

                                                           
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
            Init();
            ConstructNewFile();
        }
        /// <summary>
		/// Creates a new instance of the ExcelPackage class based on a existing file or creates a new file. 
		/// </summary>
		/// <param name="newFile">If newFile exists, it is opened.  Otherwise it is created from scratch.</param>
        public ExcelPackage(FileInfo newFile)
		{
            Init();
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
            Init();
            File = newFile;
            CreateFromTemplate(template);
		}
        /// <summary>
        /// Creates a new instance of the ExcelPackage class based on a existing template.
        /// </summary>
        /// <param name="template">The name of the Excel template to use as the basis of the new Excel file</param>
        /// <param name="useStream">if true use a strem. If false create a file in the temp dir with a random name</param>
        public ExcelPackage(FileInfo template, bool useStream)
        {
            Init();
            CreateFromTemplate(template);
            if (useStream == false)
            {
                File = new FileInfo(Path.GetTempPath() + Guid.NewGuid().ToString() + ".xlsx");
            }
        }
        /// <summary>
        /// Creates a new instance of the Excelpackage class based on a stream
        /// </summary>
        /// <param name="Stream">The stream object can be empty or contain a package. For example use Response.OutputStream to output the workbook to a webclient</param>
        public ExcelPackage(Stream newStream)
        {
            if (!(newStream.CanRead && newStream.CanWrite))
            {
                throw new Exception("The stream must be read/write");
            }
            Init();
            if (newStream.Length > 0)
            {
                _stream = newStream;
                if (newStream.CanSeek)
                {
                    _stream.Seek(0, SeekOrigin.Begin);
                }
                _package = Package.Open(_stream, FileMode.Open, FileAccess.ReadWrite);
            }
            else
            {
                CreateBlankWb();
            }
        }
        /// <summary>
        /// Creates a new instance of the Excelpackage class based on a stream
        /// </summary>
        /// <param name="newStream">This stream is copied to the output stream at load</param>
        /// <param name="templateStream">The output stream. For example Response.OutputStream to output the sheet to a webclient</param>
        public ExcelPackage(Stream newStream, Stream templateStream)
        {
            if (newStream.Length > 0)
            {
                throw(new Exception("The output stream must be empty. Length > 0"));
            }
            else if (!(newStream.CanRead && newStream.CanWrite))
            {
                throw new Exception("The stream must be read/write");
            }
            Init();
            Load(templateStream, newStream);
        }        
        #endregion
        /// <summary>
        /// Init values here
        /// </summary>
        private void Init()
        {
            Compression = CompressionOption.Normal;
        }
        /// <summary>
        /// Create a new file frp, a template
        /// </summary>
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
            if (File != null && File.Exists)
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
        }
        private void CreateBlankWb()
        {
            // save a temporary part to create the default application/xml content type
            Uri uriDefaultContentType = new Uri("/default.xml", UriKind.Relative);
            PackagePart partTemp = _package.CreatePart(uriDefaultContentType, "application/xml", Compression);

            XmlDocument workbook = Workbook.WorkbookXml; // this will create the workbook xml in the package

            // create the relationship to the main part
            _package.CreateRelationship(Workbook.WorkbookUri, TargetMode.Internal, schemaRelationships + "/officeDocument");

            // remove the temporary part that created the default xml content type
            _package.DeletePart(uriDefaultContentType);
        }
		#region Public Properties
		/// <summary>
		/// Setting DebugMode to true will cause the Save method to write the 
		/// raw XML components to the same folder as the output Excel file
		/// </summary>
		public bool DebugMode = false;

		/// <summary>
		/// Returns a reference to the package
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
                    ns.AddNamespace(string.Empty, ExcelPackage.schemaMain);
                    ns.AddNamespace("d", ExcelPackage.schemaMain);
                    ns.AddNamespace("vt", schemaVt);
                    // extended properties (app.xml)
                    ns.AddNamespace("xp", schemaExtended);
                    // custom properties
                    ns.AddNamespace("ctp", schemaCustom);
                    // core properties
                    ns.AddNamespace("cp", schemaCore);
                    // core property namespaces
                    ns.AddNamespace("dc", schemaDc);
                    ns.AddNamespace("dcterms", schemaDcTerms);
                    ns.AddNamespace("dcmitype", schemaDcmiType);
                    ns.AddNamespace("xsi", schemaXsi);

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
        /// <summary>
        /// Copies the Package to the Outstream
        /// Package is closed after it has been saved
        /// </summary>
        public void SaveAs(Stream OutputStream)
        {
            Save();
            CopyStream(_stream, ref OutputStream);
        }
        FileInfo _file = null;
        /// <summary>
        /// The output file. Null if no file is used
        /// </summary>
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
        /// <summary>
        /// The output stream
        /// </summary>
        public Stream Stream
        {
            get
            {
                return _stream;
            }
        }
		#endregion
        /// <summary>
        /// Compression option for the package
        /// </summary>
        public CompressionOption Compression { get; set; }
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
        /// <summary>
        /// Loads the specified package data from a stream.
        /// </summary>
        /// <param name="input">The input.</param>
        public void Load(Stream input)
        {
            Load(input, new MemoryStream());
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="input"></param>
        /// <param name="output"></param>
        private void Load(Stream input, Stream output)
        {
            this._stream = output;
            CopyStream(input, ref this._stream);
            this._package = Package.Open(this._stream, FileMode.Open, FileAccess.ReadWrite);
        }
        /// <summary>
        /// Copies the input stream to the output stream.
        /// </summary>
        /// <param name="inputStream">The input stream.</param>
        /// <param name="outputStream">The output stream.</param>
        private static void CopyStream(Stream inputStream, ref Stream outputStream)
        {
            if (!inputStream.CanRead)
            {
                throw (new Exception("Can not read from inputstream"));
            }
            if (!outputStream.CanWrite)
            {
                throw (new Exception("Can not write to outputstream"));
            }
            if (inputStream.CanSeek)
            {
                inputStream.Seek(0, SeekOrigin.Begin);
            }

            int bufferLength = 8096;
            Byte[] buffer = new Byte[bufferLength];
            int bytesRead = inputStream.Read(buffer, 0, bufferLength);
            // write the required bytes
            while (bytesRead > 0)
            {
                outputStream.Write(buffer, 0, bytesRead);
                bytesRead = inputStream.Read(buffer, 0, bufferLength);
            }
        }
    }
}