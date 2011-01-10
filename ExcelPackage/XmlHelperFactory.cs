using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml
{
    internal class XmlHelperInstance : XmlHelper
    {
        public XmlHelperInstance(XmlNamespaceManager namespaceManager)
            : base(namespaceManager)
        {}

        public XmlHelperInstance(XmlNamespaceManager namespaceManager, XmlNode topNode)
            : base(namespaceManager, topNode)
        {}

    }

    public static class XmlHelperFactory
    {
        public static XmlHelper Create(XmlNamespaceManager namespaceManager)
        {
            return new XmlHelperInstance(namespaceManager);
        }

        public static XmlHelper Create(XmlNamespaceManager namespaceManager, XmlNode topNode)
        {
            return new XmlHelperInstance(namespaceManager, topNode);
        }
    }
}
