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
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public abstract class Expression
    {
        protected string ExpressionString { get; private set; }
        private readonly List<Expression> _children = new List<Expression>();
        public IEnumerable<Expression> Children { get { return _children; } }
        public Expression Next { get; set; }
        public Expression Prev { get; set; }
        public IOperator Operator { get; set; }
        public abstract bool IsGroupedExpression { get; }

        public Expression()
        {

        }

        public Expression(string expression)
        {
            ExpressionString = expression;
            Operator = null;
        }

        public virtual bool ParentIsLookupFunction
        {
            get;
            set;
        }

        public virtual bool HasChildren
        {
            get { return _children.Any(); }
        }

        public virtual Expression  PrepareForNextChild()
        {
            return this;
        }

        public virtual Expression AddChild(Expression child)
        {
            if (_children.Any())
            {
                var last = _children.Last();
                child.Prev = last;
                last.Next = child;
            }
            _children.Add(child);
            return child;
        }

        public virtual Expression MergeWithNext()
        {
            var expression = this;
            if (Next != null && Operator != null)
            {
                var result = Operator.Apply(Compile(), Next.Compile());
                expression = ExpressionConverter.Instance.FromCompileResult(result);
                if (Next != null)   
                {
                    expression.Operator = Next.Operator;
                }
                else
                {
                    expression.Operator = null;
                }
                expression.Next = Next.Next;
                if (expression.Next != null) expression.Next.Prev = expression;
                expression.Prev = Prev;
            }
            if (Prev != null)
            {
                Prev.Next = expression;
            }
            return expression;
        }

        public abstract CompileResult Compile();

    }
}
