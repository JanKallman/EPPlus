using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace OfficeOpenXml.FormulaParsing
{
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

        public Guid ScopeId { get; private set; }

        public ParsingScope Parent { get; private set; }

        public RangeAddress Address { get; private set; }

        public bool IsSubtotal { get; set; }

        public void Dispose()
        {
            _parsingScopes.KillScope(this);
        }
    }
}
