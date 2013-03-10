using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
    public class ExcelAddressInfo
    {
        private ExcelAddressInfo(string address) 
        {   
            var addressOnSheet = address;
            Worksheet = string.Empty;
            if (address.Contains("!"))
            {
                var worksheetArr = address.Split('!');
                Worksheet = worksheetArr[0];
                addressOnSheet = worksheetArr[1];
            }
            if (addressOnSheet.Contains(":"))
            {
                var rangeArr = addressOnSheet.Split(':');
                StartCell = rangeArr[0];
                EndCell = rangeArr[1];
            }
            else
            {
                StartCell = addressOnSheet;
            }
            AddressOnSheet = addressOnSheet;
        }

        public static ExcelAddressInfo Parse(string address)
        {
            Require.That(address).Named("address").IsNotNullOrEmpty();
            return new ExcelAddressInfo(address);
        }

        public string Worksheet { get; private set; }

        public bool WorksheetIsSpecified
        {
            get
            {
                return !string.IsNullOrEmpty(Worksheet);
            }
        }

        public bool IsMultipleCells 
        { 
            get 
            { 
                return !string.IsNullOrEmpty(EndCell); 
            } 
        }

        public string StartCell { get; private set; }

        public string EndCell { get; private set; }

        public string AddressOnSheet { get; private set; }
    }
}
