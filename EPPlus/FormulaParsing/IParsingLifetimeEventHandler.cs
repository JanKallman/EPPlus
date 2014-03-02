using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing
{
    public interface IParsingLifetimeEventHandler
    {
        void ParsingCompleted();
    }
}
