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
 * Jan Källman		Added		26-MAR-2012
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.VBA
{
    public enum eModuleType
    {
        /// <summary>
        /// A Workbook or Worksheet objects
        /// </summary>
        Document=0,
        /// <summary>
        /// A Module
        /// </summary>
        Module=1,
        /// <summary>
        /// A class
        /// </summary>
        Class=2,
        /// <summary>
        /// Designer, typically a user form
        /// </summary>
        Designer=3
    }
    internal delegate void ModuleNameChange(string value);

    public class ExcelVBAModule
    {
        string _name = "";
        ModuleNameChange _nameChangeCallback = null;
        public ExcelVBAModule()
        {

        }
        internal ExcelVBAModule(ModuleNameChange nameChangeCallback)
        {
            _nameChangeCallback = nameChangeCallback;
        }
        public string Name 
        {   
            get
            {
                return _name;
            }
            set
            {
                _name = value;
                streamName = value;
                if (_nameChangeCallback != null)
                {
                    _nameChangeCallback(value);
                }
            }
        }
        public string Description { get; set; }
        public string Code { get; set; }
        public int HelpContext { get; set; }

        internal string streamName { get; set; }
        internal ushort Cookie { get; set; }
        internal uint ModuleOffset { get; set; }
        internal string ClassID { get; set; }
        public eModuleType Type { get; internal set; }
        public bool ReadOnly { get; set; }
        public bool Private { get; set; }
        public override string ToString()
        {
            return Name;
        }
    }
}
