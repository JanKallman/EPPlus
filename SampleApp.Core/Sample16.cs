/* 
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * EPPlus provides server-side generation of Excel 2007 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
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
 * Jan Källman                      Added       		        2018-03-20
 *******************************************************************************/
 using OfficeOpenXml;
using OfficeOpenXml.Sparkline;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Text;

namespace EPPlusSamples
{
    public static class Sample16
    {
        public static void RunSample16()
        {
            using (var package = new ExcelPackage())
            {
                //Sample fx data
                var txt = "Date;AUD;CAD;CHF;DKK;EUR;GBP;HKD;JPY;MYR;NOK;NZD;RUB;SEK;THB;TRY;USD\r\n" +
               "2016-03-01;6,17350;6,42084;8,64785;1,25668;9,37376;12,01683;1,11067;0,07599;2,06900;0,99522;5,69227;0,11665;1,00000;0,24233;2,93017;8,63185\r\n" +
               "2016-03-02;6,27223;6,42345;8,63480;1,25404;9,35350;12,14970;1,11099;0,07582;2,07401;0,99311;5,73277;0,11757;1,00000;0,24306;2,94083;8,63825\r\n" +
               "2016-03-07;6,33778;6,38403;8,50245;1,24980;9,32373;12,05756;1,09314;0,07478;2,07171;0,99751;5,77539;0,11842;1,00000;0,23973;2,91088;8,48885\r\n" +
               "2016-03-08;6,30268;6,31774;8,54066;1,25471;9,36254;12,03361;1,09046;0,07531;2,05625;0,99225;5,72501;0,11619;1,00000;0,23948;2,91067;8,47020\r\n" +
               "2016-03-09;6,32630;6,33698;8,46118;1,24399;9,28125;11,98879;1,08544;0,07467;2,04128;0,98960;5,71601;0,11863;1,00000;0,23893;2,91349;8,42945\r\n" +
               "2016-03-10;6,24241;6,28817;8,48684;1,25260;9,34350;11,99193;1,07956;0,07392;2,04500;0,98267;5,58145;0,11769;1,00000;0,23780;2,89150;8,38245\r\n" +
               "2016-03-11;6,30180;6,30152;8,48295;1,24848;9,31230;12,01194;1,07545;0,07352;2,04112;0,98934;5,62335;0,11914;1,00000;0,23809;2,90310;8,34510\r\n" +
               "2016-03-15;6,19790;6,21615;8,42931;1,23754;9,22896;11,76418;1,07026;0,07359;2,00929;0,97129;5,49278;0,11694;1,00000;0,23642;2,86487;8,30540\r\n" +
               "2016-03-16;6,18508;6,22493;8,41792;1,23543;9,21149;11,72470;1,07152;0,07318;2,01179;0,96907;5,49138;0,11836;1,00000;0,23724;2,84767;8,31775\r\n" +
               "2016-03-17;6,25214;6,30642;8,45981;1,24327;9,26623;11,86396;1,05571;0,07356;2,01706;0,98159;5,59544;0,12024;1,00000;0,23543;2,87595;8,18825\r\n" +
               "2016-03-18;6,25359;6,32400;8,47826;1,24381;9,26976;11,91322;1,05881;0,07370;2,02554;0,98439;5,59067;0,12063;1,00000;0,23538;2,86880;8,20950";

                // Add a new worksheet to the empty workbook and load the fx rates from the text
                var ws = package.Workbook.Worksheets.Add("SEKRates");
                
                //Load the sample data with a Swedish culture setting
                ws.Cells["A1"].LoadFromText(txt, new ExcelTextFormat() { Delimiter = ';', Culture = CultureInfo.GetCultureInfo("sv-SE") }, TableStyles.Light10, true);
                ws.Cells["A2:A12"].Style.Numberformat.Format = "yyyy-mm-dd";

                // Add a column sparkline for  all currencies
                ws.Cells["A15"].Value = "Column";
                var sparklineCol = ws.SparklineGroups.Add(eSparklineType.Column, ws.Cells["B15:Q15"], ws.Cells["B2:Q12"]);
                sparklineCol.High = true;
                sparklineCol.ColorHigh.SetColor(Color.Red); 

                // Add a line sparkline for  all currencies
                ws.Cells["A16"].Value = "Line";
                var sparklineLine = ws.SparklineGroups.Add(eSparklineType.Line, ws.Cells["B16:Q16"], ws.Cells["B2:Q12"]);
                sparklineLine.DateAxisRange = ws.Cells["A2:A12"];

                // Add some more random values and add a stacked sparkline.
                ws.Cells["A17"].Value = "Stacked";
                ws.Cells["B17:Q17"].LoadFromArrays(new List<object[]> { new object[] { 2, -1, 3, -4, 8, 5, -12, 18, 99, 1, -4, 12, -8, 9, 0, -8 } });
                var sparklineStacked = ws.SparklineGroups.Add(eSparklineType.Stacked, ws.Cells["R17"], ws.Cells["B17:Q17"]);
                sparklineStacked.High = true;
                sparklineStacked.ColorHigh.SetColor(Color.Red);
                sparklineStacked.Low = true;
                sparklineStacked.ColorLow.SetColor(Color.Green);
                sparklineStacked.Negative = true;
                sparklineStacked.ColorNegative.SetColor(Color.Blue);

                ws.Cells["A15:A17"].Style.Font.Bold = true;
                ws.Cells.AutoFitColumns();
                ws.Row(15).Height = 40;
                ws.Row(16).Height = 40;
                ws.Row(17).Height = 40;

                package.SaveAs(Utils.GetFileInfo("Sample16.xlsx"));
            }
        }
}
}
