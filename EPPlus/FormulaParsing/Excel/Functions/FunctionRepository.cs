/* Copyright (C) 2011  Jan Källman
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
 * Mats Alm   		                Added		                2013-12-03
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    /// <summary>
    /// This class provides methods for accessing/modifying VBA Functions.
    /// </summary>
    public class FunctionRepository : IFunctionNameProvider
    {
        private Dictionary<string, ExcelFunction> _functions = new Dictionary<string, ExcelFunction>(StringComparer.InvariantCulture);

        private FunctionRepository()
        {

        }

        public static FunctionRepository Create()
        {
            var repo = new FunctionRepository();
            repo.LoadModule(new BuiltInFunctions());
            return repo;
        }

        /// <summary>
        /// Loads a module of <see cref="ExcelFunction"/>s to the function repository.
        /// </summary>
        /// <param name="module">A <see cref="IFunctionModule"/> that can be used for adding functions</param>
        public virtual void LoadModule(IFunctionModule module)
        {
            foreach (var key in module.Functions.Keys)
            {
                var lowerKey = key.ToLower(CultureInfo.InvariantCulture);
                _functions[lowerKey] = module.Functions[key];
            }
        }

        public virtual ExcelFunction GetFunction(string name)
        {
            if(!_functions.ContainsKey(name.ToLower(CultureInfo.InvariantCulture)))
            {
                //throw new InvalidOperationException("Non supported function: " + name);
                //throw new ExcelErrorValueException("Non supported function: " + name, ExcelErrorValue.Create(eErrorType.Name));
                return null;
            }
            return _functions[name.ToLower(CultureInfo.InvariantCulture)];
        }

        /// <summary>
        /// Removes all functions from the repository
        /// </summary>
        public virtual void Clear()
        {
            _functions.Clear();
        }

        /// <summary>
        /// Returns true if the the supplied <paramref name="name"/> exists in the repository.
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public bool IsFunctionName(string name)
        {
            return _functions.ContainsKey(name.ToLower(CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Returns the names of all implemented functions.
        /// </summary>
        public IEnumerable<string> FunctionNames
        {
            get { return _functions.Keys; }
        }

        /// <summary>
        /// Adds or replaces a function.
        /// </summary>
        /// <param name="functionName"> Case-insensitive name of the function that should be added or replaced.</param>
        /// <param name="functionImpl">An implementation of an <see cref="ExcelFunction"/>.</param>
        public void AddOrReplaceFunction(string functionName, ExcelFunction functionImpl)
        {
            Require.That(functionName).Named("functionName").IsNotNullOrEmpty();
            Require.That(functionImpl).Named("functionImpl").IsNotNull();
            var fName = functionName.ToLower(CultureInfo.InvariantCulture);
            if (_functions.ContainsKey(fName))
            {
                _functions.Remove(fName);
            }
            _functions[fName] = functionImpl;
        }
    }
}
