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
 *  Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Jan Källman                      Added       		        2011-04-18
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using OfficeOpenXml;
using System.Data.SqlClient;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Drawing.Chart;
namespace EPPlusSamples
{
    /// <summary>
    /// This class shows how to use pivottables 
    /// </summary>
    public static class Sample12
    {
        public class SalesDTO
        {
            public string Title { get; set; }            
            public string FirstName { get; set; }
            public string MiddleName { get; set; }
            public string LastName { get; set; }
            public string Name
            {
                get
                {
                    return string.IsNullOrEmpty(MiddleName) ? FirstName + " " + LastName : FirstName + " " + MiddleName + " " + LastName;
                }
            }
            public DateTime OrderDate { get; set; }
            public decimal SubTotal { get; set; }
            public decimal Tax { get; set; }
            public decimal Freight { get; set; }
            public decimal Total
            {
                get
                {
                    return SubTotal + Tax + Freight;
                }
            }
        }
        public static string RunSample12(string sqlServerName, DirectoryInfo outputDir)
        {
            var list = new List<SalesDTO>();
            if (sqlServerName == "")
            {
                list = GetRandomData();
            }
            else
            {
                list = GetDataFromSQL(sqlServerName);
            }

            string file = outputDir.FullName + @"\sample12.xlsx";
            if (File.Exists(file)) File.Delete(file);
            FileInfo newFile = new FileInfo(file);

            using (ExcelPackage pck = new ExcelPackage(newFile))
            {
                // get the handle to the existing worksheet
                var wsData = pck.Workbook.Worksheets.Add("SalesData");

                var dataRange = wsData.Cells["A1"].LoadFromCollection(
                    from s in list 
                    orderby s.LastName, s.FirstName 
                    select s, 
                   true, OfficeOpenXml.Table.TableStyles.Medium2);                
                
                wsData.Cells[2, 6, dataRange.End.Row, 6].Style.Numberformat.Format = "mm-dd-yy";
                wsData.Cells[2, 7, dataRange.End.Row, 11].Style.Numberformat.Format = "#,##0";
                
                dataRange.AutoFitColumns();

                var wsPivot = pck.Workbook.Worksheets.Add("PivotSimple");
                var pivotTable1 = wsPivot.PivotTables.Add(wsPivot.Cells["A1"], dataRange, "PerEmploee");

                pivotTable1.RowFields.Add(pivotTable1.Fields[4]);
                var dataField = pivotTable1.DataFields.Add(pivotTable1.Fields[6]);
                dataField.Format="#,##0";
                pivotTable1.DataOnRows = true;

                var chart = wsPivot.Drawings.AddChart("PivotChart", eChartType.Pie, pivotTable1);
                chart.SetPosition(1, 0, 4, 0);
                chart.SetSize(600, 400);

                var wsPivot2 = pck.Workbook.Worksheets.Add("PivotDateGrp");
                var pivotTable2 = wsPivot2.PivotTables.Add(wsPivot2.Cells["A3"], dataRange, "PerEmploeeAndQuarter");

                pivotTable2.RowFields.Add(pivotTable2.Fields["Name"]);
                
                //Add a rowfield
                var rowField = pivotTable2.RowFields.Add(pivotTable2.Fields["OrderDate"]);
                //This is a date field so we want to group by Years and quaters. This will create one additional field for years.
                rowField.AddDateGrouping(eDateGroupBy.Years | eDateGroupBy.Quarters);
                //Get the Quaters field and change the texts
                var quaterField = pivotTable2.Fields.GetDateGroupField(eDateGroupBy.Quarters);
                quaterField.Items[0].Text = "<"; //Values below min date, but we use auto so its not used
                quaterField.Items[1].Text = "Q1";
                quaterField.Items[2].Text = "Q2";
                quaterField.Items[3].Text = "Q3";
                quaterField.Items[4].Text = "Q4";
                quaterField.Items[5].Text = ">"; //Values above max date, but we use auto so its not used
                
                //Add a pagefield
                var pageField = pivotTable2.PageFields.Add(pivotTable2.Fields["Title"]);
                
                //Add the data fields and format them
                dataField = pivotTable2.DataFields.Add(pivotTable2.Fields["SubTotal"]);
                dataField.Format = "#,##0";
                dataField = pivotTable2.DataFields.Add(pivotTable2.Fields["Tax"]);
                dataField.Format = "#,##0";
                dataField = pivotTable2.DataFields.Add(pivotTable2.Fields["Freight"]);
                dataField.Format = "#,##0";
                
                //We want the datafields to appear in columns
                pivotTable2.DataOnRows = false;

                pck.Save();
            }
            return file;
        }

        private static List<SalesDTO> GetRandomData()
        {
            List<SalesDTO> ret = new List<SalesDTO>();
            var firstNames = new string[] {"John", "Gunnar", "Karl", "Alice"};
            var lastNames = new string[] {"Smith", "Johansson", "Lindeman"};
            Random r = new Random();
            for (int i = 0; i < 500; i++)
            {
                ret.Add(
                    new SalesDTO()
                    {
                        FirstName = firstNames[r.Next(4)],
                        LastName = lastNames[r.Next(3)],
                        OrderDate = new DateTime(2002, 1, 1).AddDays(r.Next(1000)),
                        Title="Sales Representative",
                        SubTotal = r.Next(100, 10000),
                        Tax = 0,
                        Freight = 0
                    });
            }
            return ret;
        }

        private static List<SalesDTO> GetDataFromSQL(string sqlServerName)
        {
            string connectionStr = string.Format(@"server={0};database=AdventureWorks;Integrated Security=true;", sqlServerName);
            var ret = new List<SalesDTO>();
            // lets connect to the AdventureWorks sample database for some data
            using (SqlConnection sqlConn = new SqlConnection(connectionStr))
            {
                sqlConn.Open();
                using (SqlCommand sqlCmd = new SqlCommand("select h.Title, FirstName, MiddleName, LastName, SubTotal, OrderDate, TaxAmt, Freight, TotalDue  from Sales.SalesOrderHeader s inner join HumanResources.Employee h on s.SalesPersonID = h.EmployeeID inner join Person.Contact c on c.ContactID = h.ContactID order by LastName, FirstName, MiddleName;", sqlConn))
                {
                    using (SqlDataReader sqlReader = sqlCmd.ExecuteReader())
                    {
                        //Get the data and fill rows 5 onwards
                        while (sqlReader.Read())
                        {
                            ret.Add(new SalesDTO
                            {
                                Title = sqlReader["Title"].ToString(),
                                FirstName=sqlReader["FirstName"].ToString(),
                                MiddleName=sqlReader["MiddleName"].ToString(),
                                LastName=sqlReader["LastName"].ToString(),
                                OrderDate = (DateTime)sqlReader["OrderDate"],
                                SubTotal = (decimal)sqlReader["SubTotal"],
                                Tax=(decimal)sqlReader["TaxAmt"],
                                Freight=(decimal)sqlReader["Freight"]
                            });
                        }
                    }
                }
            }
            return ret;
        }
    }
}