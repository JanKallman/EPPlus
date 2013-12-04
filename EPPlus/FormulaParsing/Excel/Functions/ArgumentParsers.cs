using System;
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
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    public class ArgumentParsers
    {
        private static object _syncRoot = new object();
        private readonly Dictionary<DataType, ArgumentParser> _parsers = new Dictionary<DataType, ArgumentParser>();
        private readonly ArgumentParserFactory _parserFactory;

        public ArgumentParsers()
            : this(new ArgumentParserFactory())
        {

        }

        public ArgumentParsers(ArgumentParserFactory factory)
        {
            Require.That(factory).Named("argumentParserfactory").IsNotNull();
            _parserFactory = factory;
        }

        public ArgumentParser GetParser(DataType dataType)
        {
            if (!_parsers.ContainsKey(dataType))
            {
                lock (_syncRoot)
                {
                    if (!_parsers.ContainsKey(dataType))
                    {
                        _parsers.Add(dataType, _parserFactory.CreateArgumentParser(dataType));
                    }
                }
            }
            return _parsers[dataType];
        }
    }
}
