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
 * Mats Alm   		                Added       		        2013-03-01 (Prior file history on https://github.com/swmal/ExcelFormulaParser)
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
    public class FormulaDependencies
    {
        public FormulaDependencies()
            : this(new FormulaDependencyFactory())
        {

        }

        public FormulaDependencies(FormulaDependencyFactory formulaDependencyFactory)
        {
            _formulaDependencyFactory = formulaDependencyFactory;
        }

        private readonly FormulaDependencyFactory _formulaDependencyFactory;
        private readonly Dictionary<string, FormulaDependency> _dependencies = new Dictionary<string, FormulaDependency>();

        public IEnumerable<KeyValuePair<string, FormulaDependency>> Dependencies { get { return _dependencies; } }

        public void AddFormulaScope(ParsingScope parsingScope)
        {
            //var dependency = _formulaDependencyFactory.Create(parsingScope);
            //var address = parsingScope.Address.ToString();
            //if (!_dependencies.ContainsKey(address))
            //{
            //    _dependencies.Add(address, dependency);
            //}
            //if (parsingScope.Parent != null)
            //{
            //    var parentAddress = parsingScope.Parent.Address.ToString();
            //    if (_dependencies.ContainsKey(parentAddress))
            //    {
            //        var parent = _dependencies[parentAddress];
            //        parent.AddReferenceTo(parsingScope.Address);
            //        dependency.AddReferenceFrom(parent.Address);
            //    }
            //}
        }

        public void Clear()
        {
            _dependencies.Clear();
        }
    }
}
