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
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace OfficeOpenXml.FormulaParsing
{
    /// <summary>
    /// Represents a parsing of a single input or workbook addrses.
    /// </summary>
    public class ParsingScope : IDisposable
    {
        private readonly ParsingScopes _parsingScopes;

        public ParsingScope(ParsingScopes parsingScopes, RangeAddress address)
            : this(parsingScopes, null, address)
        {
        }

        public ParsingScope(ParsingScopes parsingScopes, ParsingScope parent, RangeAddress address)
        {
            _parsingScopes = parsingScopes;
            Parent = parent;
            Address = address;
            ScopeId = Guid.NewGuid();
        }

        /// <summary>
        /// Id of the scope.
        /// </summary>
        public Guid ScopeId { get; private set; }

        /// <summary>
        /// The calling scope.
        /// </summary>
        public ParsingScope Parent { get; private set; }

        /// <summary>
        /// The address of the cell currently beeing parsed.
        /// </summary>
        public RangeAddress Address { get; private set; }

        /// <summary>
        /// True if the current scope is a Subtotal function beeing executed.
        /// </summary>
        public bool IsSubtotal { get; set; }

        public void Dispose()
        {
            _parsingScopes.KillScope(this);
        }
    }
}
