using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;
using OfficeOpenXml.DataValidation;
using System.Xml;
using System.Globalization;

namespace EPPlusTest.DataValidation
{
    public abstract class ValidationTestBase
    {
        protected ExcelPackage _package;
        protected ExcelWorksheet _sheet;
        protected XmlNode _dataValidationNode;
        protected XmlNamespaceManager _namespaceManager;
        protected CultureInfo _cultureInfo;

        public void SetupTestData()
        {
            _package = new ExcelPackage();
            _sheet = _package.Workbook.Worksheets.Add("test");
            _cultureInfo = new CultureInfo("en-US");
        }

        public void CleanupTestData()
        {
            _package = null;
            _sheet = null;
            _namespaceManager = null;
        }

        protected string GetTestOutputPath(string fileName)
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), fileName);
        }

        protected void SaveTestOutput(string fileName)
        {
            var path = GetTestOutputPath(fileName);
            if (File.Exists(path))
            {
                File.Delete(path);
            }
            _package.SaveAs(new FileInfo(path));
        }

        protected void LoadXmlTestData(string address, string validationType, string formula1Value)
        {
            var xmlDoc = new XmlDocument();
            _namespaceManager = new XmlNamespaceManager(xmlDoc.NameTable);
            _namespaceManager.AddNamespace("d", "urn:a");
            var sb = new StringBuilder();
            sb.AppendFormat("<dataValidation xmlns:d=\"urn:a\" type=\"{0}\" sqref=\"{1}\">", validationType, address);
            sb.AppendFormat("<d:formula1>{0}</d:formula1>", formula1Value);
            sb.Append("</dataValidation>");
            xmlDoc.LoadXml(sb.ToString());
            _dataValidationNode = xmlDoc.DocumentElement;
        }

        protected void LoadXmlTestData(string address, string validationType, string formula1Value, bool showErrorMsg, bool showInputMsg)
        {
            var xmlDoc = new XmlDocument();
            _namespaceManager = new XmlNamespaceManager(xmlDoc.NameTable);
            _namespaceManager.AddNamespace("d", "urn:a");
            var sb = new StringBuilder();
            sb.AppendFormat("<dataValidation xmlns:d=\"urn:a\" type=\"{0}\" sqref=\"{1}\" ", validationType, address);
            sb.AppendFormat(" showErrorMessage=\"{0}\" showInputMessage=\"{1}\">", showErrorMsg ? 1 : 0, showInputMsg ? 1 : 0);
            sb.AppendFormat("<d:formula1>{0}</d:formula1>", formula1Value);
            sb.Append("</dataValidation>");
            xmlDoc.LoadXml(sb.ToString());
            _dataValidationNode = xmlDoc.DocumentElement;
        }

        protected void LoadXmlTestData(string address, string validationType, string formula1Value, string prompt, string promptTitle, string error, string errorTitle)
        {
            var xmlDoc = new XmlDocument();
            _namespaceManager = new XmlNamespaceManager(xmlDoc.NameTable);
            _namespaceManager.AddNamespace("d", "urn:a");
            var sb = new StringBuilder();
            sb.AppendFormat("<dataValidation xmlns:d=\"urn:a\" type=\"{0}\" sqref=\"{1}\"", validationType, address);
            sb.AppendFormat(" prompt=\"{0}\" promptTitle=\"{1}\"", prompt, promptTitle);
            sb.AppendFormat(" error=\"{0}\" errorTitle=\"{1}\">", error, errorTitle);
            sb.AppendFormat("<d:formula1>{0}</d:formula1>", formula1Value);
            sb.Append("</dataValidation>");
            xmlDoc.LoadXml(sb.ToString());
            _dataValidationNode = xmlDoc.DocumentElement;
        }

    }
}
