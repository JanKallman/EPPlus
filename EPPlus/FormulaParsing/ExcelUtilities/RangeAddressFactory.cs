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
 * Author							Change						Date
 * ******************************************************************************
 * Mats Alm   		                Added       		        2013-03-01 (Prior file history on https://github.com/swmal/ExcelFormulaParser)
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
    public class RangeAddressFactory
    {
        private readonly ExcelDataProvider _excelDataProvider;
        private readonly AddressTranslator _addressTranslator;
        private readonly IndexToAddressTranslator _indexToAddressTranslator;

        public RangeAddressFactory(ExcelDataProvider excelDataProvider)
            : this(excelDataProvider, new AddressTranslator(excelDataProvider), new IndexToAddressTranslator(excelDataProvider, ExcelReferenceType.RelativeRowAndColumn))
        {
           
            
        }

        public RangeAddressFactory(ExcelDataProvider excelDataProvider, AddressTranslator addressTranslator, IndexToAddressTranslator indexToAddressTranslator)
        {
            Require.That(excelDataProvider).Named("excelDataProvider").IsNotNull();
            Require.That(addressTranslator).Named("addressTranslator").IsNotNull();
            Require.That(indexToAddressTranslator).Named("indexToAddressTranslator").IsNotNull();
            _excelDataProvider = excelDataProvider;
            _addressTranslator = addressTranslator;
            _indexToAddressTranslator = indexToAddressTranslator;
        }

        public RangeAddress Create(int col, int row)
        {
            return Create(string.Empty, col, row);
        }

        public RangeAddress Create(string worksheetName, int col, int row)
        {
            return new RangeAddress()
            {
                Address = _indexToAddressTranslator.ToAddress(col, row),
                Worksheet = worksheetName,
                FromCol = col,
                ToCol = col,
                FromRow = row,
                ToRow = row
            };
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="worksheetName">will be used if no worksheet name is specified in <paramref name="address"/></param>
        /// <param name="address">address of a range</param>
        /// <returns></returns>
        public RangeAddress Create(string worksheetName, string address)
        {
            Require.That(address).Named("range").IsNotNullOrEmpty();
            //var addressInfo = ExcelAddressInfo.Parse(address);
            var adr = new ExcelAddressBase(address);

            var sheet = string.IsNullOrEmpty(adr.WorkSheet) ? worksheetName : adr.WorkSheet;
            var rangeAddress = new RangeAddress()
            {
                Address = adr.Address,
                Worksheet = sheet,
                FromRow = adr._fromRow,
                FromCol = adr._fromCol,
                ToRow = adr._toRow,
                ToCol = adr._toCol
            };

            //if (addressInfo.IsMultipleCells)
            //{
            //    HandleMultipleCellAddress(rangeAddress, addressInfo);
            //}
            //else
            //{
            //    HandleSingleCellAddress(rangeAddress, addressInfo);
            //}
            return rangeAddress;
        }

        public RangeAddress Create(string range)
        {
            Require.That(range).Named("range").IsNotNullOrEmpty();
            //var addressInfo = ExcelAddressInfo.Parse(range);
            var adr = new ExcelAddressBase(range);
            var rangeAddress = new RangeAddress()
            {
                Address = adr.Address,
                Worksheet = adr.WorkSheet ?? "",
                FromRow = adr._fromRow,
                FromCol = adr._fromCol,
                ToRow = adr._toRow,
                ToCol = adr._toCol
            };
           
            //if (addressInfo.IsMultipleCells)
            //{
            //    HandleMultipleCellAddress(rangeAddress, addressInfo);
            //}
            //else
            //{
            //    HandleSingleCellAddress(rangeAddress, addressInfo);
            //}
            return rangeAddress;
        }

        private void HandleSingleCellAddress(RangeAddress rangeAddress, ExcelAddressInfo addressInfo)
        {
            int col, row;
            _addressTranslator.ToColAndRow(addressInfo.StartCell, out col, out row);
            rangeAddress.FromCol = col;
            rangeAddress.ToCol = col;
            rangeAddress.FromRow = row;
            rangeAddress.ToRow = row;
        }

        private void HandleMultipleCellAddress(RangeAddress rangeAddress, ExcelAddressInfo addressInfo)
        {
            int fromCol, fromRow;
            _addressTranslator.ToColAndRow(addressInfo.StartCell, out fromCol, out fromRow);
            int toCol, toRow;
            _addressTranslator.ToColAndRow(addressInfo.EndCell, out toCol, out toRow, AddressTranslator.RangeCalculationBehaviour.LastPart);
            rangeAddress.FromCol = fromCol;
            rangeAddress.ToCol = toCol;
            rangeAddress.FromRow = fromRow;
            rangeAddress.ToRow = toRow;
        }
    }
}
