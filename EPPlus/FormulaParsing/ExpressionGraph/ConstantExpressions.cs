/* Copyright (C) 2011  Jan Källman
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
 * Author Change                      Date
 *******************************************************************************
 * Mats Alm Added		                2016-12-27
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public static class ConstantExpressions
    {
        public static Expression Percent
        {
            get { return new ConstantExpression("Percent", () => new CompileResult(0.01, DataType.Decimal)); }
        }
    }

    public class ConstantExpression : AtomicExpression
    {
        private readonly Func<CompileResult> _factoryMethod;

        public ConstantExpression(string title, Func<CompileResult> factoryMethod)
            : base(title)
        {
            _factoryMethod = factoryMethod;
        }

        public override CompileResult Compile()
        {
            return _factoryMethod();
        }
    }
}
