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
 *******************************************************************************
 * Jan Källman		Added		2009-10-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using OfficeOpenXml.Table.PivotTable;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Bar chart
    /// </summary>
    public sealed class ExcelBarChart : ExcelChart
    {
        #region "Constructors"
        //internal ExcelBarChart(ExcelDrawings drawings, XmlNode node) :
        //    base(drawings, node/*, 1*/)
        //{
        //    SetChartNodeText("");
        //}
        //internal ExcelBarChart(ExcelDrawings drawings, XmlNode node, eChartType type) :
        //    base(drawings, node, type)
        //{
        //    SetChartNodeText("");

        //    SetTypeProperties(drawings, type);
        //}
        internal ExcelBarChart(ExcelDrawings drawings, XmlNode node, eChartType type, ExcelChart topChart, ExcelPivotTable PivotTableSource) :
            base(drawings, node, type, topChart, PivotTableSource)
        {
            SetChartNodeText("");

            SetTypeProperties(drawings, type);
        }

        internal ExcelBarChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, Packaging.ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode) :
           base(drawings, node, uriChart, part, chartXml, chartNode)
        {
            SetChartNodeText(chartNode.Name);
        }

        internal ExcelBarChart(ExcelChart topChart, XmlNode chartNode) : 
            base(topChart, chartNode)
        {
            SetChartNodeText(chartNode.Name);
        }
        #endregion
        #region "Private functions"
        //string _chartTopPath="c:chartSpace/c:chart/c:plotArea/{0}";
        private void SetChartNodeText(string chartNodeText)
        {
            if(string.IsNullOrEmpty(chartNodeText))
            {
                chartNodeText = GetChartNodeText();
            }
            //_chartTopPath = string.Format(_chartTopPath, chartNodeText);
            //_directionPath = string.Format(_directionPath, _chartTopPath);
            //_shapePath = string.Format(_shapePath, _chartTopPath);
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
            if (/*type == eChartType.ColumnClustered ||
                type == eChartType.ColumnStacked ||
                type == eChartType.ColumnStacked100 ||*/
                type == eChartType.Column3D ||
                type == eChartType.ColumnClustered3D ||
                type == eChartType.ColumnStacked3D ||
                type == eChartType.ColumnStacked1003D ||
                /*type == eChartType.BarClustered ||
                type == eChartType.BarStacked ||
                type == eChartType.BarStacked100 ||*/
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
        string _directionPath = "c:barDir/@val";
        /// <summary>
        /// Direction, Bar or columns
        /// </summary>
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
        string _shapePath = "c:shape/@val";
        /// <summary>
        /// The shape of the bar/columns
        /// </summary>
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
        /// <summary>
        /// Access to datalabel properties
        /// </summary>
        public ExcelChartDataLabel DataLabel
        {
            get
            {
                if (_DataLabel == null)
                {
                    _DataLabel = new ExcelChartDataLabel(NameSpaceManager, ChartNode);
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
        internal override eChartType GetChartType(string name)
        {
            if (name == "barChart")
            {
                if (this.Direction == eDirection.Bar)
                {
                    if (Grouping == eGrouping.Stacked)
                    {
                        return eChartType.BarStacked;
                    }
                    else if (Grouping == eGrouping.PercentStacked)
                    {
                        return eChartType.BarStacked100;
                    }
                    else
                    {
                        return eChartType.BarClustered;
                    }
                }
                else
                {
                    if (Grouping == eGrouping.Stacked)
                    {
                        return eChartType.ColumnStacked;
                    }
                    else if (Grouping == eGrouping.PercentStacked)
                    {
                        return eChartType.ColumnStacked100;
                    }
                    else
                    {
                        return eChartType.ColumnClustered;
                    }
                }
            }
            if (name == "bar3DChart")
            {
                #region "Bar Shape"
                if (this.Shape==eShape.Box)
                {
                    if (this.Direction == eDirection.Bar)
                    {
                        if (Grouping == eGrouping.Stacked)
                        {
                            return eChartType.BarStacked3D;
                        }
                        else if (Grouping == eGrouping.PercentStacked)
                        {
                            return eChartType.BarStacked1003D;
                        }
                        else
                        {
                            return eChartType.BarClustered3D;
                        }
                    }
                    else
                    {
                        if (Grouping == eGrouping.Stacked)
                        {
                            return eChartType.ColumnStacked3D;
                        }
                        else if (Grouping == eGrouping.PercentStacked)
                        {
                            return eChartType.ColumnStacked1003D;
                        }
                        else
                        {
                            return eChartType.ColumnClustered3D;
                        }
                    }
                }
                #endregion
                #region "Cone Shape"
                if (this.Shape == eShape.Cone || this.Shape == eShape.ConeToMax)
                {
                    if (this.Direction == eDirection.Bar)
                    {
                        if (Grouping == eGrouping.Stacked)
                        {
                            return eChartType.ConeBarStacked;
                        }
                        else if (Grouping == eGrouping.PercentStacked)
                        {
                            return eChartType.ConeBarStacked100;
                        }
                        else if (Grouping == eGrouping.Clustered)
                        {
                            return eChartType.ConeBarClustered;
                        }
                    }
                    else
                    {
                        if (Grouping == eGrouping.Stacked)
                        {
                            return eChartType.ConeColStacked;
                        }
                        else if (Grouping == eGrouping.PercentStacked)
                        {
                            return eChartType.ConeColStacked100;
                        }
                        else if (Grouping == eGrouping.Clustered)
                        {
                            return eChartType.ConeColClustered;
                        }
                        else
                        {
                            return eChartType.ConeCol;
                        }
                    }
                }
                #endregion
                #region "Cylinder Shape"
                if (this.Shape == eShape.Cylinder)
                {
                    if (this.Direction == eDirection.Bar)
                    {
                        if (Grouping == eGrouping.Stacked)
                        {
                            return eChartType.CylinderBarStacked;
                        }
                        else if (Grouping == eGrouping.PercentStacked)
                        {
                            return eChartType.CylinderBarStacked100;
                        }
                        else if (Grouping == eGrouping.Clustered)
                        {
                            return eChartType.CylinderBarClustered;
                        }
                    }
                    else
                    {
                        if (Grouping == eGrouping.Stacked)
                        {
                            return eChartType.CylinderColStacked;
                        }
                        else if (Grouping == eGrouping.PercentStacked)
                        {
                            return eChartType.CylinderColStacked100;
                        }
                        else if (Grouping == eGrouping.Clustered)
                        {
                            return eChartType.CylinderColClustered;
                        }
                        else
                        {
                            return eChartType.CylinderCol;
                        }
                    }
                }
                #endregion
                #region "Pyramide Shape"
                if (this.Shape == eShape.Pyramid || this.Shape == eShape.PyramidToMax)
                {
                    if (this.Direction == eDirection.Bar)
                    {
                        if (Grouping == eGrouping.Stacked)
                        {
                            return eChartType.PyramidBarStacked;
                        }
                        else if (Grouping == eGrouping.PercentStacked)
                        {
                            return eChartType.PyramidBarStacked100;
                        }
                        else if (Grouping == eGrouping.Clustered)
                        {
                            return eChartType.PyramidBarClustered;
                        }
                    }
                    else
                    {
                        if (Grouping == eGrouping.Stacked)
                        {
                            return eChartType.PyramidColStacked;
                        }
                        else if (Grouping == eGrouping.PercentStacked)
                        {
                            return eChartType.PyramidColStacked100;
                        }
                        else if (Grouping == eGrouping.Clustered)
                        {
                            return eChartType.PyramidColClustered;
                        }
                        else
                        {
                            return eChartType.PyramidCol;
                        }
                    }
                }
                #endregion
            }
            return base.GetChartType(name);
        }
    }
}
