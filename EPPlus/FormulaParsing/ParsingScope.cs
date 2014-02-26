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
