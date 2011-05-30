using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml
{
    internal class XmlHelperInstance : XmlHelper
    {
        internal XmlHelperInstance(XmlNamespaceManager namespaceManager)
            : base(namespaceManager)
        {}

        internal XmlHelperInstance(XmlNamespaceManager namespaceManager, XmlNode topNode)
            : base(namespaceManager, topNode)
        {}

    }

    internal static class XmlHelperFactory
    {
        internal static XmlHelper Create(XmlNamespaceManager namespaceManager)
        {
            return new XmlHelperInstance(namespaceManager);
        }

        internal static XmlHelper Create(XmlNamespaceManager namespaceManager, XmlNode topNode)
        {
            return new XmlHelperInstance(namespaceManager, topNode);
        }
    }
}
