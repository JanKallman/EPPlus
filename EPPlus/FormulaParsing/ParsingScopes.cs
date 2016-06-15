using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace OfficeOpenXml.FormulaParsing
{
    /// <summary>
    /// This class implements a stack on which instances of <see cref="ParsingScope"/>
    /// are put. Each ParsingScope represents the parsing of an address in the workbook.
    /// </summary>
    public class ParsingScopes
    {
        private readonly IParsingLifetimeEventHandler _lifetimeEventHandler;

        public ParsingScopes(IParsingLifetimeEventHandler lifetimeEventHandler)
        {
            _lifetimeEventHandler = lifetimeEventHandler;
        }
        private Stack<ParsingScope> _scopes = new Stack<ParsingScope>();

        /// <summary>
        /// Creates a new <see cref="ParsingScope"/> and puts it on top of the stack.
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
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


        /// <summary>
        /// The current parsing scope.
        /// </summary>
        public virtual ParsingScope Current
        {
            get { return _scopes.Count() > 0 ? _scopes.Peek() : null; }
        }

        /// <summary>
        /// Removes the current scope, setting the calling scope to current.
        /// </summary>
        /// <param name="parsingScope"></param>
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
