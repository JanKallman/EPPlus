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

namespace OfficeOpenXml
{
	#region class ExcelHeaderFooterText
	/// <summary>
	/// Helper class for ExcelHeaderFooter - simply stores the three header or footer
	/// text strings. 
	/// </summary>
	public class ExcelHeaderFooterText
	{
		/// <summary>
		/// Sets the text to appear on the left hand side of the header (or footer) on the worksheet.
		/// </summary>
		public string LeftAlignedText = null;
		/// <summary>
		/// Sets the text to appear in the center of the header (or footer) on the worksheet.
		/// </summary>
		public string CenteredText = null;
		/// <summary>
		/// Sets the text to appear on the right hand side of the header (or footer) on the worksheet.
		/// </summary>
		public string RightAlignedText = null;
	}
	#endregion

	#region ExcelHeaderFooter
	/// <summary>
	/// Represents the Header and Footer on an Excel Worksheet
	/// </summary>
	public class ExcelHeaderFooter
	{
		#region Static Properties
		/// <summary>
		/// Use this to insert the page number into the header or footer of the worksheet
		/// </summary>
		public const string PageNumber = @"&P";
		/// <summary>
		/// Use this to insert the number of pages into the header or footer of the worksheet
		/// </summary>
		public const string NumberOfPages = @"&N";
		/// <summary>
		/// Use this to insert the name of the worksheet into the header or footer of the worksheet
		/// </summary>
		public const string SheetName = @"&A";
		/// <summary>
		/// Use this to insert the full path to the folder containing the workbook into the header or footer of the worksheet
		/// </summary>
		public const string FilePath = @"&Z";
		/// <summary>
		/// Use this to insert the name of the workbook file into the header or footer of the worksheet
		/// </summary>
		public const string FileName = @"&F";
		/// <summary>
		/// Use this to insert the current date into the header or footer of the worksheet
		/// </summary>
		public const string CurrentDate = @"&D";
		/// <summary>
		/// Use this to insert the current time into the header or footer of the worksheet
		/// </summary>
		public const string CurrentTime = @"&T";
		#endregion

		#region ExcelHeaderFooter Private Properties
		private XmlElement _headerFooterNode;
		private ExcelHeaderFooterText _oddHeader;
		private ExcelHeaderFooterText _oddFooter;
		private ExcelHeaderFooterText _evenHeader;
		private ExcelHeaderFooterText _evenFooter;
		private ExcelHeaderFooterText _firstHeader;
		private ExcelHeaderFooterText _firstFooter;
		private System.Nullable<bool> _alignWithMargins = null;
		private System.Nullable<bool> _differentOddEven = null;
		private System.Nullable<bool> _differentFirst = null;
		#endregion

		#region ExcelHeaderFooter Constructor
		/// <summary>
		/// ExcelHeaderFooter Constructor
		/// For internal use only!
		/// </summary>
		/// <param name="HeaderFooterNode"></param>
		protected internal ExcelHeaderFooter(XmlElement HeaderFooterNode)
		{
			if (HeaderFooterNode.Name != "headerFooter")
				throw new Exception("ExcelHeaderFooter Error: Passed invalid headerFooter node");
			else
			{
				_headerFooterNode = HeaderFooterNode;
				// TODO: populate structure based on XML content
			}
		}
		#endregion

		#region alignWithMargins
		/// <summary>
		/// Gets/sets the alignWithMargins attribute
		/// </summary>
		public bool AlignWithMargins
		{
			get
			{
				if (_alignWithMargins == null)
				{
					_alignWithMargins = false;
					XmlAttribute attr = (XmlAttribute)_headerFooterNode.Attributes.GetNamedItem("alignWithMargins");
					if (attr != null)
					{
						if (attr.Value == "1")
							_alignWithMargins = true;
					}
				}
				return _alignWithMargins.Value;
			}
			set
			{
				_alignWithMargins = value;
				XmlAttribute attr = (XmlAttribute)_headerFooterNode.Attributes.GetNamedItem("alignWithMargins");
				if (attr == null)
				{
					attr = _headerFooterNode.Attributes.Append(_headerFooterNode.OwnerDocument.CreateAttribute("alignWithMargins"));
				}
				if (value)
					attr.Value = "1";
				else
					attr.Value = "0";
			}
		}
		#endregion

		#region differentOddEven
		/// <summary>
		/// Gets/sets the flag that tells Excel to display different headers and footers on odd and even pages.
		/// </summary>
		public bool differentOddEven
		{
			get
			{
				if (_differentOddEven == null)
				{
					_differentOddEven = false;
					XmlAttribute attr = (XmlAttribute)_headerFooterNode.Attributes.GetNamedItem("differentOddEven");
					if (attr != null)
					{
						if (attr.Value == "1")
							_differentOddEven = true;
					}
				}
				return _differentOddEven.Value;
			}
			set
			{
				_differentOddEven = value;
				XmlAttribute attr = (XmlAttribute)_headerFooterNode.Attributes.GetNamedItem("differentOddEven");
				if (attr == null)
				{
					attr = _headerFooterNode.Attributes.Append(_headerFooterNode.OwnerDocument.CreateAttribute("differentOddEven"));
				}
				if (value)
					attr.Value = "1";
				else
					attr.Value = "0";
			}
		}
		#endregion

		#region differentFirst
		/// <summary>
		/// Gets/sets the flag that tells Excel to display different headers and footers on the first page of the worksheet.
		/// </summary>
		public bool differentFirst
		{
			get
			{
				if (_differentFirst == null)
				{
					_differentFirst = false;
					XmlAttribute attr = (XmlAttribute)_headerFooterNode.Attributes.GetNamedItem("differentFirst");
					if (attr != null)
					{
						if (attr.Value == "1")
							_differentFirst = true;
					}
				}
				return _differentFirst.Value;
			}
			set
			{
				_differentFirst = value;
				XmlAttribute attr = (XmlAttribute)_headerFooterNode.Attributes.GetNamedItem("differentFirst");
				if (attr == null)
				{
					attr = _headerFooterNode.Attributes.Append(_headerFooterNode.OwnerDocument.CreateAttribute("differentFirst"));
				}
				if (value)
					attr.Value = "1";
				else
					attr.Value = "0";
			}
		}
		#endregion

		#region ExcelHeaderFooter Public Properties
		/// <summary>
		/// Provides access to a ExcelHeaderFooterText class that allows you to set the values of the header on odd numbered pages of the document.
		/// If you want the same header on both odd and even pages, then only set values in this ExcelHeaderFooterText class.
		/// </summary>
		public ExcelHeaderFooterText oddHeader { get { if (_oddHeader == null) _oddHeader = new ExcelHeaderFooterText(); return _oddHeader; } }
		/// <summary>
		/// Provides access to a ExcelHeaderFooterText class that allows you to set the values of the footer on odd numbered pages of the document.
		/// If you want the same footer on both odd and even pages, then only set values in this ExcelHeaderFooterText class.
		/// </summary>
		public ExcelHeaderFooterText oddFooter { get { if (_oddFooter == null) _oddFooter = new ExcelHeaderFooterText(); return _oddFooter; } }
		// evenHeader and evenFooter set differentOddEven = true
		/// <summary>
		/// Provides access to a ExcelHeaderFooterText class that allows you to set the values of the header on even numbered pages of the document.
		/// </summary>
		public ExcelHeaderFooterText evenHeader { get { if (_evenHeader == null) _evenHeader = new ExcelHeaderFooterText(); differentOddEven = true; return _evenHeader; } }
		/// <summary>
		/// Provides access to a ExcelHeaderFooterText class that allows you to set the values of the footer on even numbered pages of the document.
		/// </summary>
		public ExcelHeaderFooterText evenFooter { get { if (_evenFooter == null) _evenFooter = new ExcelHeaderFooterText(); differentOddEven = true; return _evenFooter; } }
		// firstHeader and firstFooter set differentFirst = true
		/// <summary>
		/// Provides access to a ExcelHeaderFooterText class that allows you to set the values of the header on the first page of the document.
		/// </summary>
		public ExcelHeaderFooterText firstHeader { get { if (_firstHeader == null) _firstHeader = new ExcelHeaderFooterText(); differentFirst = true; return _firstHeader; } }
		/// <summary>
		/// Provides access to a ExcelHeaderFooterText class that allows you to set the values of the footer on the first page of the document.
		/// </summary>
		public ExcelHeaderFooterText firstFooter { get { if (_firstFooter == null) _firstFooter = new ExcelHeaderFooterText(); differentFirst = true; return _firstFooter; } }
		#endregion

		#region Save  //  ExcelHeaderFooter
		/// <summary>
		/// Saves the header and footer information to the worksheet XML
		/// </summary>
		protected internal void Save()
		{
			//  The header/footer elements must appear in this order, if they appear:
			//  <oddHeader />
			//  <oddFooter />
			//  <evenHeader />
			//  <evenFooter />
			//  <firstHeader />
			//  <firstFooter />

			XmlNode node;
			if (_oddHeader != null)
			{
				node = _headerFooterNode.AppendChild(_headerFooterNode.OwnerDocument.CreateElement("oddHeader", ExcelPackage.schemaMain));
				node.InnerText = GetHeaderFooterText(oddHeader);
			}
			if (_oddFooter != null)
			{
				node = _headerFooterNode.AppendChild(_headerFooterNode.OwnerDocument.CreateElement("oddFooter", ExcelPackage.schemaMain));
				node.InnerText = GetHeaderFooterText(oddFooter);
			}

			// only set evenHeader and evenFooter 
			if (differentOddEven)
			{
				if (_evenHeader != null)
				{
					node = _headerFooterNode.AppendChild(_headerFooterNode.OwnerDocument.CreateElement("evenHeader", ExcelPackage.schemaMain));
					node.InnerText = GetHeaderFooterText(evenHeader);
				}
				if (_evenFooter != null)
				{
					node = _headerFooterNode.AppendChild(_headerFooterNode.OwnerDocument.CreateElement("evenFooter", ExcelPackage.schemaMain));
					node.InnerText = GetHeaderFooterText(evenFooter);
				}
			}

			// only set firstHeader and firstFooter
			if (differentFirst)
			{
				if (_firstHeader != null)
				{
					node = _headerFooterNode.AppendChild(_headerFooterNode.OwnerDocument.CreateElement("firstHeader", ExcelPackage.schemaMain));
					node.InnerText = GetHeaderFooterText(firstHeader);
				}
				if (_firstFooter != null)
				{
					node = _headerFooterNode.AppendChild(_headerFooterNode.OwnerDocument.CreateElement("firstFooter", ExcelPackage.schemaMain));
					node.InnerText = GetHeaderFooterText(firstFooter);
				}
			}
		}

		/// <summary>
		/// Helper function for Save
		/// </summary>
		/// <param name="inStruct"></param>
		/// <returns></returns>
		protected internal string GetHeaderFooterText(ExcelHeaderFooterText inStruct)
		{
			string retValue = "";
			if (inStruct.LeftAlignedText != null)
				retValue += "&L" + inStruct.LeftAlignedText;
			if (inStruct.CenteredText != null)
				retValue += "&C" + inStruct.CenteredText;
			if (inStruct.RightAlignedText != null)
				retValue += "&R" + inStruct.RightAlignedText;
			return retValue;
		}
		#endregion
	}
	#endregion

}
