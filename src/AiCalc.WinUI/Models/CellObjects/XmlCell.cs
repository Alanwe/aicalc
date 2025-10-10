using System.Collections.Generic;
using System.Xml;

namespace AiCalc.Models.CellObjects;

public class XmlCell : CellObjectBase
{
    public override CellObjectType ObjectType => CellObjectType.Xml;
    
    public string XmlText { get; set; }
    
    public override string? DisplayValue => $"<> XML ({XmlText?.Length ?? 0} chars)";
    
    public XmlCell(string xmlText) : base(xmlText)
    {
        XmlText = xmlText ?? "<root/>";
    }
    
    public override bool IsValid()
    {
        try
        {
            var doc = new XmlDocument();
            doc.LoadXml(XmlText);
            return true;
        }
        catch
        {
            return false;
        }
    }
    
    public override IEnumerable<string> GetAvailableOperations()
    {
        yield return "XML_VALIDATE";
        yield return "XML_FORMAT";
        yield return "XML_TO_JSON";
        yield return "XML_XPATH";
        yield return "XML_TO_TABLE";
        yield return "XML_TRANSFORM";
    }
    
    public override ICellObject Clone() => new XmlCell(XmlText);
}
