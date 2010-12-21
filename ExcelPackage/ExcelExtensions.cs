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
 * Michael Tran			Created		        2010-12-15
 *******************************************************************************/

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace OfficeOpenXml
{
	///<summary>
	/// Extensions for Excel
	///</summary>
	public static class ExcelExtensions
	{
		#region ExcelRange Extensions
		public static ExcelColumn Column(this ExcelRange excelRange)
		{
			return GetColumn(excelRange.Worksheet, excelRange._fromCol);
		}

		public static ExcelRow Row(this ExcelRange excelRange)
		{
			return GetRow(excelRange.Worksheet, excelRange._fromRow);
		}

		public static ExcelRange SetValue<TValue>(this ExcelRange excelRange, TValue value)
		{
			excelRange.Value = value;
			return excelRange;
		}
		#endregion

		#region Worksheet Extensions
		public static ExcelRange SetCellValue<TValue>(this ExcelWorksheet sheet, int cellRow, int cellColumn, TValue value)
		{
			return sheet.Cells[cellRow, cellColumn].SetValue(value);
		}

		public static ExcelRange SetCellValue<TValue>(this ExcelWorksheet sheet, string cellAddress, TValue value)
		{
			return sheet.Cells[cellAddress].SetValue(value);
		}

		public static ExcelRange SetCellValue<TValue>(this ExcelWorksheet sheet, int cellRowFrom, int cellRowTo, int cellColumnFrom, int cellColumnTo, TValue value)
		{
			return sheet.Cells[cellRowFrom, cellColumnFrom, cellRowTo, cellColumnTo].SetValue(value);
		}

		public static TValue GetCellValue<TValue>(this ExcelWorksheet sheet, int cellRow, int cellColumn)
		{
			return sheet.GetCellValue(cellRow, cellColumn, default(TValue));
		}

		public static TValue GetCellValue<TValue>(this ExcelWorksheet sheet, int cellRow, int cellColumn, TValue defaultValue)
		{
			return ConvertTo(sheet.Cell(cellRow, cellColumn).Value, defaultValue);
		}

		public static ExcelRange GetCell(this ExcelWorksheet sheet, int cellRow, int cellColumn)
		{
			return sheet.Cells[cellRow, cellColumn];
		}

		public static ExcelRange GetCell(this ExcelWorksheet sheet, int cellRowFrom, int cellColumnFrom, int cellRowTo, int cellColumnTo)
		{
			return sheet.Cells[cellRowFrom, cellColumnFrom, cellRowTo, cellColumnTo];
		}

		public static ExcelRow GetRow(this ExcelWorksheet sheet, int rowNumber)
		{
			return sheet.Row(rowNumber);
		}

		public static ExcelColumn GetColumn(this ExcelWorksheet sheet, int columnNumber)
		{
			return sheet.Column(columnNumber);
		}

		public static ExcelRange Add<TValue>(this ExcelWorksheet sheet, ExcelRange excelRange, InsertDirection direction, IEnumerable<TValue> values)
		{
			var rowNumber = excelRange._fromRow;
			var columnNumber = excelRange._fromCol;
			ExcelRange endCell = excelRange;
			foreach (var value in values)
			{
				var cell = sheet.GetCell(rowNumber, columnNumber);
				cell.SetValue(value);
				endCell = cell;
				if (direction == InsertDirection.Across)
					columnNumber++;
				else
					rowNumber++;
			}
			return sheet.GetCell(excelRange._fromRow, excelRange._toCol, endCell._fromRow, endCell._toCol);
		}

		#endregion

		#region ExcelColumn Extensions
		public static ExcelColumn Hide(this ExcelColumn column)
		{
			return column.Hide(true);
		}

		public static ExcelColumn Hide(this ExcelColumn column, bool hide)
		{
			column.Hidden = hide;
			return column;
		}
		#endregion

		#region ExcelRow Extensions
		public static ExcelRow Hide(this ExcelRow row)
		{
			return row.Hide(true);
		}

		public static ExcelRow Hide(this ExcelRow row, bool hide)
		{
			row.Hidden = hide;
			return row;
		}
		#endregion

		#region Private Convert Methods
		/// <summary>
		/// 	Converts an object to the specified target type or returns the default value.
		/// </summary>
		/// <typeparam name = "T"></typeparam>
		/// <param name = "value">The value.</param>
		/// <returns>The target type</returns>
		static T ConvertTo<T>(this object value)
		{
			return value.ConvertTo(default(T));
		}

		/// <summary>
		/// 	Converts an object to the specified target type or returns the default value.
		/// </summary>
		/// <typeparam name = "T"></typeparam>
		/// <param name = "value">The value.</param>
		/// <param name = "defaultValue">The default value.</param>
		/// <returns>The target type</returns>
		static T ConvertTo<T>(this object value, T defaultValue)
		{
			if (value != null)
			{
				try
				{
					var targetType = typeof(T);
					var valueType = value.GetType();

					if (valueType == targetType) return (T)value;

					var converter = TypeDescriptor.GetConverter(value);
					if (converter != null)
					{
						if (converter.CanConvertTo(targetType))
							return (T)converter.ConvertTo(value, targetType);
					}

					converter = TypeDescriptor.GetConverter(targetType);
					if (converter != null)
					{
						if (converter.CanConvertFrom(valueType))
							return (T)converter.ConvertFrom(value);
					}
				}
				catch (Exception e)
				{
					return defaultValue;
				}
			}
			return defaultValue;
		}
		#endregion
	}

	public enum InsertDirection
	{
		Across,
		Down
	}
}
