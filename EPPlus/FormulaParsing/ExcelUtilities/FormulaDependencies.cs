using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
    public class FormulaDependencies
    {
        public FormulaDependencies()
            : this(new FormulaDependencyFactory())
        {

        }

        public FormulaDependencies(FormulaDependencyFactory formulaDependencyFactory)
        {
            _formulaDependencyFactory = formulaDependencyFactory;
        }

        private readonly FormulaDependencyFactory _formulaDependencyFactory;
        private readonly Dictionary<string, FormulaDependency> _dependencies = new Dictionary<string, FormulaDependency>();

        public IEnumerable<KeyValuePair<string, FormulaDependency>> Dependencies { get { return _dependencies; } }

        public void AddFormulaScope(ParsingScope parsingScope)
        {
            //var dependency = _formulaDependencyFactory.Create(parsingScope);
            //var address = parsingScope.Address.ToString();
            //if (!_dependencies.ContainsKey(address))
            //{
            //    _dependencies.Add(address, dependency);
            //}
            //if (parsingScope.Parent != null)
            //{
            //    var parentAddress = parsingScope.Parent.Address.ToString();
            //    if (_dependencies.ContainsKey(parentAddress))
            //    {
            //        var parent = _dependencies[parentAddress];
            //        parent.AddReferenceTo(parsingScope.Address);
            //        dependency.AddReferenceFrom(parent.Address);
            //    }
            //}
        }

        public void Clear()
        {
            _dependencies.Clear();
        }
    }
}
