using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    /// <summary>
    /// This class provides methods for accessing/modifying VBA Functions.
    /// </summary>
    public class FunctionRepository : IFunctionNameProvider
    {
        private Dictionary<string, ExcelFunction> _functions = new Dictionary<string, ExcelFunction>();

        private FunctionRepository()
        {

        }

        public static FunctionRepository Create()
        {
            var repo = new FunctionRepository();
            repo.LoadModule(new BuiltInFunctions());
            return repo;
        }

        /// <summary>
        /// Loads a module of <see cref="ExcelFunction"/>s to the function repository.
        /// </summary>
        /// <param name="module">A <see cref="IFunctionModule"/> that can be used for adding functions</param>
        public virtual void LoadModule(IFunctionModule module)
        {
            foreach (var key in module.Functions.Keys)
            {
                var lowerKey = key.ToLower();
                _functions[lowerKey] = module.Functions[key];
            }
        }

        public virtual ExcelFunction GetFunction(string name)
        {
            if(!_functions.ContainsKey(name.ToLower()))
            {
                throw new InvalidOperationException("Non supported function: " + name);
            }
            return _functions[name.ToLower()];
        }

        /// <summary>
        /// Removes all functions from the repository
        /// </summary>
        public virtual void Clear()
        {
            _functions.Clear();
        }

        /// <summary>
        /// Returns true if the the supplied <paramref name="name"/> exists in the repository.
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public bool IsFunctionName(string name)
        {
            return _functions.ContainsKey(name.ToLower());
        }
    }
}
