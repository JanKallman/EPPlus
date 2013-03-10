using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace OfficeOpenXml.FormulaParsing
{
    public class ParsingScopes
    {
        private readonly IParsingLifetimeEventHandler _lifetimeEventHandler;

        public ParsingScopes(IParsingLifetimeEventHandler lifetimeEventHandler)
        {
            _lifetimeEventHandler = lifetimeEventHandler;
        }
        private Stack<ParsingScope> _scopes = new Stack<ParsingScope>();

        public virtual ParsingScope NewScope(RangeAddress address)
        {
            ParsingScope scope;
            if (_scopes.Count() > 0)
            {
                scope = new ParsingScope(this, _scopes.Peek(), address);
            }
            else
            {
                scope = new ParsingScope(this, address);
            }
            _scopes.Push(scope);
            return scope;
        }

        public virtual ParsingScope Current
        {
            get { return _scopes.Count() > 0 ? _scopes.Peek() : null; }
        }

        public virtual void KillScope(ParsingScope parsingScope)
        {
            _scopes.Pop();
            if (_scopes.Count() == 0)
            {
                _lifetimeEventHandler.ParsingCompleted();
            }
        }
    }
}
