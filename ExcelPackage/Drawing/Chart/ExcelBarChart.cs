/* 
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * EPPlus provides server-side generation of Excel 2007 spreadsheets.
 *
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

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Bar chart
    /// </summary>
    public sealed class ExcelBarChart : ExcelChart
    {
        #region "Constructors"
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
        #endregion
        #region "Private functions"
        string _chartTopPath="c:chartSpace/c:chart/c:plotArea/{0}";
        private void SetChartNodeText()
        {
            string text = GetChartNodeText();
            _chartTopPath=string.Format(_chartTopPath, text);
            _directionPath = string.Format(_directionPath, _chartTopPath);
            _shapePath = string.Format(_shapePath, _chartTopPath);
        }
        private void SetTypeProperties(ExcelDrawings drawings, eChartType type)
        {
            /******* Bar direction *******/
            if (type == eChartType.BarClustered ||
                type == eChartType.BarStacked ||
                type == eChartType.BarStacked100 ||
                type == eChartType.BarClustered3D ||
                type == eChartType.BarStacked3D ||
                type == eChartType.BarStacked1003D ||
                type == eChartType.ConeBarClustered ||
                type == eChartType.ConeBarStacked ||
                type == eChartType.ConeBarStacked100 ||
                type == eChartType.CylinderBarClustered ||
                type == eChartType.CylinderBarStacked ||
                type == eChartType.CylinderBarStacked100 ||
                type == eChartType.PyramidBarClustered ||
                type == eChartType.PyramidBarStacked ||
                type == eChartType.PyramidBarStacked100)
            {
                Direction = eDirection.Bar;
            }
            else if (
                type == eChartType.ColumnClustered ||
                type == eChartType.ColumnStacked ||
                type == eChartType.ColumnStacked100 ||
                type == eChartType.Column3D ||
                type == eChartType.ColumnClustered3D ||
                type == eChartType.ColumnStacked3D ||
                type == eChartType.ColumnStacked1003D ||
                type == eChartType.ConeCol ||
                type == eChartType.ConeColClustered ||
                type == eChartType.ConeColStacked ||
                type == eChartType.ConeColStacked100 ||
                type == eChartType.CylinderCol ||
                type == eChartType.CylinderColClustered ||
                type == eChartType.CylinderColStacked ||
                type == eChartType.CylinderColStacked100 ||
                type == eChartType.PyramidCol ||
                type == eChartType.PyramidColClustered ||
                type == eChartType.PyramidColStacked ||
                type == eChartType.PyramidColStacked100)
            {
                Direction = eDirection.Column;
            }

            /****** Shape ******/
            if (type == eChartType.ColumnClustered ||
                type == eChartType.ColumnStacked ||
                type == eChartType.ColumnStacked100 ||
                type == eChartType.Column3D ||
                type == eChartType.ColumnClustered3D ||
                type == eChartType.ColumnStacked3D ||
                type == eChartType.ColumnStacked1003D ||
                type == eChartType.BarClustered ||
                type == eChartType.BarStacked ||
                type == eChartType.BarStacked100 ||
                type == eChartType.BarClustered3D ||
                type == eChartType.BarStacked3D ||
                type == eChartType.BarStacked1003D)
            {
                Shape = eShape.Box;
            }
            else if (
                type == eChartType.CylinderBarClustered ||
                type == eChartType.CylinderBarStacked ||
                type == eChartType.CylinderBarStacked100 ||
                type == eChartType.CylinderCol ||
                type == eChartType.CylinderColClustered ||
                type == eChartType.CylinderColStacked ||
                type == eChartType.CylinderColStacked100)
            {
                Shape = eShape.Cylinder;
            }
            else if (
                type == eChartType.ConeBarClustered ||
                type == eChartType.ConeBarStacked ||
                type == eChartType.ConeBarStacked100 ||
                type == eChartType.ConeCol ||
                type == eChartType.ConeColClustered ||
                type == eChartType.ConeColStacked ||
                type == eChartType.ConeColStacked100)
            {
                Shape = eShape.Cone;
            }
            else if (
                type == eChartType.PyramidBarClustered ||
                type == eChartType.PyramidBarStacked ||
                type == eChartType.PyramidBarStacked100 ||
                type == eChartType.PyramidCol ||
                type == eChartType.PyramidColClustered ||
                type == eChartType.PyramidColStacked ||
                type == eChartType.PyramidColStacked100)
            {
                Shape = eShape.Pyramid;
            }
        }
        #endregion
        #region "Properties"
        string _directionPath = "{0}/c:barDir/@val";
        public eDirection Direction
        {
            get
            {
                return GetDirectionEnum(_chartXmlHelper.GetXmlNodeString(_directionPath));
            }
            internal set
            {
                _chartXmlHelper.SetXmlNodeString(_directionPath, GetDirectionText(value));
            }
        }
        string _shapePath = "{0}/c:shape/@val";
        public eShape Shape
        {
            get
            {
                return GetShapeEnum(_chartXmlHelper.GetXmlNodeString(_shapePath));
            }
            internal set
            {
                _chartXmlHelper.SetXmlNodeString(_shapePath, GetShapeText(value));
            }
        }
        ExcelChartDataLabel _DataLabel = null;
        public ExcelChartDataLabel DataLabel
        {
            get
            {
                if (_DataLabel == null)
                {
                    _DataLabel = new ExcelChartDataLabel(NameSpaceManager, _chartXmlHelper.TopNode.SelectSingleNode(_chartTopPath, NameSpaceManager));
                }
                return _DataLabel;
            }
        }
        #endregion
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
    }
}
