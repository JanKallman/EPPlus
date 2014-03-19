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
    public class FormulaDependency
    {
        public FormulaDependency(ParsingScope scope)
	    {
            ScopeId = scope.ScopeId;
            Address = scope.Address;
	    }
        public Guid ScopeId { get; private set; }

        public RangeAddress Address { get; private set; }

        private List<RangeAddress> _referencedBy = new List<RangeAddress>();

        private List<RangeAddress> _references = new List<RangeAddress>();

        public virtual void AddReferenceFrom(RangeAddress rangeAddress)
        {
            if (Address.CollidesWith(rangeAddress) || _references.Exists(x => x.CollidesWith(rangeAddress)))
            {
                throw new CircularReferenceException("Circular reference detected at " + rangeAddress.ToString());
            }
            _referencedBy.Add(rangeAddress);
        }

        public virtual void AddReferenceTo(RangeAddress rangeAddress)
        {
            if (Address.CollidesWith(rangeAddress) || _referencedBy.Exists(x => x.CollidesWith(rangeAddress)))
            {
                throw new CircularReferenceException("Circular reference detected at " + rangeAddress.ToString());
            }
            _references.Add(rangeAddress);
        }
    }
}
