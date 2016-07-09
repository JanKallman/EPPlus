/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author				Change						Date
 * ******************************************************************************
 * Richard Tallent		Initial Release				2012-08-13
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.Style
{
	public class ExcelRichTextHtmlUtility
	{

		/// <summary>
		/// Provides basic HTML support by converting well-behaved HTML into appropriate RichText blocks.
		/// HTML support is limited, and does not include font colors, sizes, or typefaces at this time,
		/// and also does not support CSS style attributes. It does support line breaks using the BR tag.
		///
		/// This routine parses the HTML into RegEx pairings of an HTML tag and the text until the NEXT 
		/// tag (if any). The tag is parsed to determine the setting change to be applied to the last set
		/// of settings, and if the text is not blank, a new block is added to rich text.
		/// </summary>
		/// <param name="range"></param>
		/// <param name="html">The HTML to parse into RichText</param>
		/// <param name="defaultFontName"></param>
		/// <param name="defaultFontSize"></param>

		public static void SetRichTextFromHtml(ExcelRange range, string html, string defaultFontName, short defaultFontSize)
		{
			// Reset the cell value, just in case there is an existing RichText value.
			range.Value = "";

			// Sanity check for blank values, skips creating Regex objects for performance.
			if (String.IsNullOrEmpty(html))
			{
				range.IsRichText = false;
				return;
			}

			// Change all BR tags to line breaks. http://epplus.codeplex.com/discussions/238692/
			// Cells with line breaks aren't necessarily considered rich text, so this is performed
			// before parsing the HTML tags.
			html = System.Text.RegularExpressions.Regex.Replace(html, @"<br[^>]*>", "\r\n", RegexOptions.Compiled | RegexOptions.IgnoreCase);

			string tag;
			string text;
			ExcelRichText thisrt = null;
			bool isFirst = true;

			// Get all pairs of legitimate tags and the text between them. This loop will
			// only execute if there is at least one start or end tag.
			foreach (Match m in System.Text.RegularExpressions.Regex.Matches(html, @"<(/?[a-z]+)[^<>]*>([\s\S]*?)(?=</?[a-z]+[^<>]*>|$)", RegexOptions.Compiled | RegexOptions.IgnoreCase))
			{
				if (isFirst)
				{
					// On the very first match, set up the initial rich text object with
					// the defaults for the text BEFORE the match.
					range.IsRichText = true;
					thisrt = range.RichText.Add(CleanText(html.Substring(0, m.Index)));	// May be 0-length
					thisrt.Size = defaultFontSize;										// Set the default font size
					thisrt.FontName = defaultFontName;									// Set the default font name
					isFirst = false;
				}
				// Get the tag and the block of text until the NEXT tag or EOS. If there are HTML entities
				// encoded, unencode them, they should be passed to RichText as normal characters (other
				// than non-breaking spaces, which should be replaced with normal spaces, they break Excel.
				tag = m.Groups[1].Captures[0].Value;
				text = CleanText(m.Groups[2].Captures[0].Value);

				if (thisrt.Text == "")
				{
					// The most recent rich text block wasn't *actually* used last time around, so update
					// the text and keep it as the "current" block. This happens with the first block if
					// it starts with a tag, and may happen later if tags come one right after the other.
					thisrt.Text = text;
				}
				else
				{
					// The current rich text block has some text, so create a new one. RichText.Add()
					// automatically applies the settings from the previous block, other than vertical
					// alignment.
					thisrt = range.RichText.Add(text);
				}
				// Override the settings based on the current tag, keep all other settings.
				SetStyleFromTag(tag, thisrt);
			}

			if (thisrt == null)
			{
				// No HTML tags were found, so treat this as a normal text value.
				range.IsRichText = false;
				range.Value = CleanText(html);
			}
			else if (String.IsNullOrEmpty(thisrt.Text))
			{
				// Rich text was found, but the last node contains no text, so remove it. This can happen if,
				// say, the end of the string is an end tag or unsupported tag (common).
				range.RichText.Remove(thisrt);

				// Failsafe -- the HTML may be just tags, no text, in which case there may be no rich text
				// directives that remain. If that is the case, turn off rich text and treat this like a blank
				// cell value.
				if (range.RichText.Count == 0)
				{
					range.IsRichText = false;
					range.Value = "";
				}

			}

		}

		private static void SetStyleFromTag(string tag, ExcelRichText settings)
		{
			switch (tag.ToLower())
			{
				case "b":
				case "strong":
					settings.Bold = true;
					break;
				case "i":
				case "em":
					settings.Italic = true;
					break;
				case "u":
					settings.UnderLine = true;
					break;
				case "s":
				case "strike":
					settings.Strike = true;
					break;
				case "sup":
					settings.VerticalAlign = ExcelVerticalAlignmentFont.Superscript;
					break;
				case "sub":
					settings.VerticalAlign = ExcelVerticalAlignmentFont.Subscript;
					break;
				case "/b":
				case "/strong":
					settings.Bold = false;
					break;
				case "/i":
				case "/em":
					settings.Italic = false;
					break;
				case "/u":
					settings.UnderLine = false;
					break;
				case "/s":
				case "/strike":
					settings.Strike = false;
					break;
				case "/sup":
				case "/sub":
					settings.VerticalAlign = ExcelVerticalAlignmentFont.None;
					break;
				default:
					// unsupported HTML, no style change
					break;
			}
		}

		private static string CleanText(string s)
		{
			// Need to convert HTML entities (named or numbered) into actual Unicode characters
			s = System.Web.HttpUtility.HtmlDecode(s);
			// Remove any non-breaking spaces, kills Excel
			s = s.Replace("\u00A0", " ");
			return s;
		}

	}
}
