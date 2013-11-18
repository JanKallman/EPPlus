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
    public sealed class ExcelSurfaceChart : ExcelChart
    {
        #region "Constructors"
        internal ExcelSurfaceChart(ExcelDrawings drawings, XmlNode node, eChartType type, ExcelChart topChart, ExcelPivotTable PivotTableSource) :
            base(drawings, node, type, topChart, PivotTableSource)
        {
            Init();
        }
        internal ExcelSurfaceChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, Packaging.ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode) :
           base(drawings, node, uriChart, part, chartXml, chartNode)
        {
            Init();
        }

        internal ExcelSurfaceChart(ExcelChart topChart, XmlNode chartNode) : 
            base(topChart, chartNode)
        {
            Init();
        }
        private void Init()
        {
 	        _floor=new ExcelChartSurface(NameSpaceManager, _chartXmlHelper.TopNode.SelectSingleNode("c:floor", NameSpaceManager));
            _backWall = new ExcelChartSurface(NameSpaceManager, _chartXmlHelper.TopNode.SelectSingleNode("c:sideWall", NameSpaceManager));
            _sideWall = new ExcelChartSurface(NameSpaceManager, _chartXmlHelper.TopNode.SelectSingleNode("c:backWall", NameSpaceManager));
            SetTypeProperties();
        }
        #endregion


        /// <summary>
        /// 3D-settings
        /// </summary>
        public ExcelView3D View3D
        {
            get
            {
                if (IsType3D())
                {
                    return new ExcelView3D(NameSpaceManager, ChartXml.SelectSingleNode("//c:view3D", NameSpaceManager));
                }
                else
                {
                    throw (new Exception("Charttype does not support 3D"));
                }

            }
        }
        ExcelChartSurface _floor;
        public ExcelChartSurface Floor
        {
            get
            {
                return _floor;
            }
        }
        ExcelChartSurface _sideWall;
        public ExcelChartSurface SideWall
        {
            get
            {
                return _sideWall;
            }
        }
        ExcelChartSurface _backWall;
        public ExcelChartSurface BackWall
        {
            get
            {
                return _backWall;
            }
        }
        const string WIREFRAME_PATH = "c:wireframe/@val";
        public bool Wireframe
        {
            get
            {
                return _chartXmlHelper.GetXmlNodeBool(WIREFRAME_PATH);
            }
            set
            {
                _chartXmlHelper.SetXmlNodeBool(WIREFRAME_PATH, value);
            }
        }        
        internal void SetTypeProperties()
        {
               if(ChartType==eChartType.SurfaceWireframe || ChartType==eChartType.SurfaceTopViewWireframe)
               {
                   Wireframe=true;
               }
               else 
               {
                   Wireframe=false;
               }

                if(ChartType==eChartType.SurfaceTopView || ChartType==eChartType.SurfaceTopViewWireframe)
                {
                   View3D.RotY = 0;
                   View3D.RotX = 90;
                }
                else
                {
                   View3D.RotY = 20;
                   View3D.RotX = 15;
                }
                View3D.RightAngleAxes = false;
                View3D.Perspective = 0;
                Axis[1].CrossBetween = eCrossBetween.MidCat;
        }
        internal override eChartType GetChartType(string name)
        {
            if(Wireframe)
            {
                if (name == "surfaceChart")
                {
                    return eChartType.SurfaceTopViewWireframe;
                }
                else
                {
                    return eChartType.SurfaceWireframe;
                }
            }
            else
            {
                if (name == "surfaceChart")
                {
                    return eChartType.SurfaceTopView;
                }
                else
                {
                    return eChartType.Surface;
                }
            }
        }
    }
}
