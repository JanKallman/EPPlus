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
 * Eyal Seagull        Added       		  2012-04-03
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using System.Globalization;
namespace OfficeOpenXml.ConditionalFormatting
{
    /// <summary>
    /// Databar
    /// </summary>
    public class ExcelConditionalFormattingDataBar
      : ExcelConditionalFormattingRule,
        IExcelConditionalFormattingDataBarGroup
    {
        /****************************************************************************************/

        #region Private Properties

        #endregion Private Properties

        /****************************************************************************************/

        #region Constructors
        /// <summary>
        /// 
        /// </summary>
        /// <param name="type"></param>
        /// <param name="priority"></param>
        /// <param name="address"></param>
        /// <param name="worksheet"></param>
        /// <param name="itemElementNode"></param>
        /// <param name="namespaceManager"></param>
        internal ExcelConditionalFormattingDataBar(
          eExcelConditionalFormattingRuleType type,
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet,
          XmlNode itemElementNode,
          XmlNamespaceManager namespaceManager)
            : base(
              type,
              address,
              priority,
              worksheet,
              itemElementNode,
              (namespaceManager == null) ? worksheet.NameSpaceManager : namespaceManager)
        {
            SchemaNodeOrder = new string[] { "cfvo", "color" };
            //Create the <dataBar> node inside the <cfRule> node
            if (itemElementNode!=null && itemElementNode.HasChildNodes)
            {
                bool high=false;
                foreach (XmlNode node in itemElementNode.SelectNodes("d:dataBar/d:cfvo", NameSpaceManager))
                {
                    if (high == false)
                    {
                        LowValue = new ExcelConditionalFormattingIconDataBarValue(
                                type,
                                address,
                                worksheet,
                                node,
                                namespaceManager);
                        high = true;
                    }
                    else
                    {
                        HighValue = new ExcelConditionalFormattingIconDataBarValue(
                                type,
                                address,
                                worksheet,
                                node,
                                namespaceManager);
                    }
                }
            }
            else
            {
                var iconSetNode = CreateComplexNode(
                  Node,
                  ExcelConditionalFormattingConstants.Paths.DataBar);

                var lowNode = iconSetNode.OwnerDocument.CreateElement(ExcelConditionalFormattingConstants.Paths.Cfvo, ExcelPackage.schemaMain);
                iconSetNode.AppendChild(lowNode);
                LowValue = new ExcelConditionalFormattingIconDataBarValue(eExcelConditionalFormattingValueObjectType.Min,
                        0,
                        "",
                        eExcelConditionalFormattingRuleType.DataBar,
                        address,
                        priority,
                        worksheet,
                        lowNode,
                        namespaceManager);

                var highNode = iconSetNode.OwnerDocument.CreateElement(ExcelConditionalFormattingConstants.Paths.Cfvo, ExcelPackage.schemaMain);
                iconSetNode.AppendChild(highNode);
                HighValue = new ExcelConditionalFormattingIconDataBarValue(eExcelConditionalFormattingValueObjectType.Max,
                        0,
                        "",
                        eExcelConditionalFormattingRuleType.DataBar,
                        address,
                        priority,
                        worksheet,
                        highNode,
                        namespaceManager);
            }
            Type = type;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="priority"></param>
        /// <param name="address"></param>
        /// <param name="worksheet"></param>
        /// <param name="itemElementNode"></param>
        internal ExcelConditionalFormattingDataBar(
          eExcelConditionalFormattingRuleType type,
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet,
          XmlNode itemElementNode)
            : this(
              type,
              address,
              priority,
              worksheet,
              itemElementNode,
              null)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="priority"></param>
        /// <param name="address"></param>
        /// <param name="worksheet"></param>
        internal ExcelConditionalFormattingDataBar(
          eExcelConditionalFormattingRuleType type,
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet)
            : this(
              type,
              address,
              priority,
              worksheet,
              null,
              null)
        {
        }
        #endregion Constructors
        private const string _showValuePath="d:dataBar/@showValue";
        public bool ShowValue
        {
            get
            {
                return GetXmlNodeBool(_showValuePath, true);
            }
            set
            {
                SetXmlNodeBool(_showValuePath, value);
            }
        }


        public ExcelConditionalFormattingIconDataBarValue LowValue
        {
            get;
            internal set;
        }

        public ExcelConditionalFormattingIconDataBarValue HighValue
        {
            get;
            internal set;
        }


        private const string _colorPath = "d:dataBar/d:color/@rgb";
        public Color Color
        {
            get
            {
                var rgb=GetXmlNodeString(_colorPath);
                if(!string.IsNullOrEmpty(rgb))
                {
                    return Color.FromArgb(int.Parse(rgb, NumberStyles.HexNumber));
                }
                return Color.White;
            }
            set
            {
                SetXmlNodeString(_colorPath, value.ToArgb().ToString("X"));
            }
        }
    }
}