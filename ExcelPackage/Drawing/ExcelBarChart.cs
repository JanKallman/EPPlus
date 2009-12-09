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
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing
{
    public class ExcelBarChart : ExcelChart
    {
        internal ExcelBarChart(ExcelDrawings drawings, XmlNode node) :
            base(drawings, node)
        {
            SetChartNodeText();
        }
        internal ExcelBarChart(ExcelDrawings drawings, XmlNode node, eChartType type) :
            base(drawings, node, type)
        {
            SetChartNodeText();

            SetTypeProperties(drawings, type);
        }
        private void SetChartNodeText()
        {
            string text = GetChartNodeText();
            _directionPath = string.Format(_directionPath, text);
            _shapePath = string.Format(_shapePath,text);
        }

        string _directionPath = "c:chartSpace/c:chart/c:plotArea/{0}/c:barDir/@val";
        public eDirection Direction
        {
            get
            {
                return GetDirectionEnum(_chartXmlHelper.GetXmlNode(_directionPath));
            }
            set
            {
                _chartXmlHelper.SetXmlNode(_directionPath, GetDirectionText(value));
            }
        }
        string _shapePath = "c:chartSpace/c:chart/c:plotArea/{0}/c:shape/@val";
        public eShape Shape
        {
            get
            {
                return GetShapeEnum(_chartXmlHelper.GetXmlNode(_shapePath));
            }
            set
            {
                _chartXmlHelper.SetXmlNode(_shapePath, GetShapeText(value));
            }
        }
        #region "Direction Enum Traslation"
        private string GetDirectionText(eDirection direction)
        {
            switch (direction)
            {
                case eDirection.Bar:
                    return "bar";
                default:
                    return "col";
            }
        }

        private eDirection GetDirectionEnum(string direction)
        {
            switch (direction)
            {
                case "bar":
                    return eDirection.Bar;
                default:
                    return eDirection.Column;
            }
        }
        #endregion
        #region "Shape Enum Translation"
        private string GetShapeText(eShape Shape)
        {
            switch (Shape)
            {
                case eShape.Box:
                    return "box";
                case eShape.Cone:
                    return "cone";
                case eShape.ConeToMax:
                    return "coneToMax";
                case eShape.Cylinder:
                    return "cylinder";
                case eShape.Pyramid:
                    return "pyramid";
                case eShape.PyramidToMax:
                    return "pyramidToMax";
                default:
                    return "box";
            }
        }

        private eShape GetShapeEnum(string text)
        {
            switch (text)
            {
                case "box":
                    return eShape.Box;
                case "cone":
                    return eShape.Cone;
                case "coneToMax":
                    return eShape.ConeToMax;
                case "cylinder":
                    return eShape.Cylinder;
                case "pyramid":
                    return eShape.Pyramid;
                case "pyramidToMax":
                    return eShape.PyramidToMax;
                default:
                    return eShape.Box;
            }
        }
        #endregion
        private void SetTypeProperties(ExcelDrawings drawings, eChartType type)
        {
            /******* Bar direction *******/
            if (type == eChartType.xlBarClustered ||
                type == eChartType.xlBarStacked ||
                type == eChartType.xlBarStacked100 ||
                type == eChartType.xl3DBarClustered ||
                type == eChartType.xl3DBarStacked ||
                type == eChartType.xl3DBarStacked100 ||
                type == eChartType.xlConeBarClustered ||
                type == eChartType.xlConeBarStacked ||
                type == eChartType.xlConeBarStacked100 ||
                type == eChartType.xlCylinderBarClustered ||
                type == eChartType.xlCylinderBarStacked ||
                type == eChartType.xlCylinderBarStacked100 ||
                type == eChartType.xlPyramidBarClustered ||
                type == eChartType.xlPyramidBarStacked ||
                type == eChartType.xlPyramidBarStacked100)
            {
                Direction = eDirection.Bar;
            }
            else if (
                type == eChartType.xlColumnClustered ||
                type == eChartType.xlColumnStacked ||
                type == eChartType.xlColumnStacked100 ||
                type == eChartType.xl3DColumn ||
                type == eChartType.xl3DColumnClustered ||
                type == eChartType.xl3DColumnStacked ||
                type == eChartType.xl3DColumnStacked100 ||
                type == eChartType.xlConeCol ||
                type == eChartType.xlConeColClustered ||
                type == eChartType.xlConeColStacked ||
                type == eChartType.xlConeColStacked100 ||
                type == eChartType.xlCylinderCol ||
                type == eChartType.xlCylinderColClustered ||
                type == eChartType.xlCylinderColStacked ||
                type == eChartType.xlCylinderColStacked100 ||
                type == eChartType.xlPyramidCol ||
                type == eChartType.xlPyramidColClustered ||
                type == eChartType.xlPyramidColStacked ||
                type == eChartType.xlPyramidColStacked100)
            {
                Direction = eDirection.Column;
            }

            /****** Shape ******/
            if (type == eChartType.xlColumnClustered ||
                type == eChartType.xlColumnStacked ||
                type == eChartType.xlColumnStacked100 ||
                type == eChartType.xl3DColumn ||
                type == eChartType.xl3DColumnClustered ||
                type == eChartType.xl3DColumnStacked ||
                type == eChartType.xl3DColumnStacked100)
            {
                Shape = eShape.Box;
            }
            else if (
                type == eChartType.xlCylinderBarClustered ||
                type == eChartType.xlCylinderBarStacked ||
                type == eChartType.xlCylinderBarStacked100 ||
                type == eChartType.xlCylinderCol ||
                type == eChartType.xlCylinderColClustered ||
                type == eChartType.xlCylinderColStacked ||
                type == eChartType.xlCylinderColStacked100)
            {
                Shape = eShape.Cylinder;
            }
            else if (
                type == eChartType.xlConeBarClustered ||
                type == eChartType.xlConeBarStacked ||
                type == eChartType.xlConeBarStacked100 ||
                type == eChartType.xlConeCol ||
                type == eChartType.xlConeColClustered ||
                type == eChartType.xlConeColStacked ||
                type == eChartType.xlConeColStacked100)
            {
                Shape = eShape.Cone;
            }
            else if (
                type == eChartType.xlPyramidBarClustered ||
                type == eChartType.xlPyramidBarStacked ||
                type == eChartType.xlPyramidBarStacked100 ||
                type == eChartType.xlPyramidCol ||
                type == eChartType.xlPyramidColClustered ||
                type == eChartType.xlPyramidColStacked ||
                type == eChartType.xlPyramidColStacked100)
            {
                Shape = eShape.Pyramid;
            }
        }
    }
}
