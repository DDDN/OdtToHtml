using System.Collections.Generic;
using System.Xml.Linq;

namespace DDDN.Office.Odf
{
    public interface IOdfStyle
    {
        string Name { get; }
        string Type { get; }
        string NamespaceName { get; }
        string ParentStyleName { get; }
        string Family { get; }
        List<OdfStyleAttr> Attrs { get; }
        Dictionary<string, List<OdfStylePropAttr>> PropAttrs { get; }
        void AddPropertyAttributes(XElement element);
    }
}
