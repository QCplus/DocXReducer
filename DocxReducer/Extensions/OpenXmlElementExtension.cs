using DocumentFormat.OpenXml;

namespace DocxReducer.Extensions
{
    internal static class OpenXmlElementExtension
    {
        public static void CloneChildrenTo(this OpenXmlElement source, OpenXmlElement target)
        {
            foreach (var c in source.Elements())
                target.Append(c.CloneNode(true));
        }
    }
}
