using System.Xml;
using DocumentFormat.OpenXml;

namespace XmlMappingsRepro
{
    /// <summary>
    ///     Hack to set default namespace instead of x:
    ///     https://github.com/OfficeDev/Open-XML-SDK/issues/90
    /// </summary>
    internal class OpenXmlElementNamespaceInjector<TElement> where TElement : OpenXmlPartRootElement
    {
        private const string Xmlns = "xmlns";
        private readonly XmlDocument _document;

        public OpenXmlElementNamespaceInjector(TElement rootElement)
        {
            _document = new XmlDocument();
            _document.LoadXml(rootElement.OuterXml);
        }

        private static string GetNamespacePrefixDefinition(string prefix)
        {
            return string.IsNullOrWhiteSpace(prefix) ? Xmlns : $"{Xmlns}:{prefix}";
        }

        public OpenXmlElementNamespaceInjector<TElement> CopyToDefault(string prefix)
        {
            var uri = _document?.DocumentElement?.GetAttribute(GetNamespacePrefixDefinition(prefix));
            _document?.DocumentElement?.SetAttribute(GetNamespacePrefixDefinition(string.Empty), uri);
            return this;
        }

        public TElement GetInstance()
        {
            var type = typeof(TElement);
            var ctor = type.GetConstructor(new[] { typeof(string) });
            return (TElement)ctor?.Invoke(new object[] { _document.OuterXml });
        }
    }
}