using System.Collections.Generic;

namespace AiCalc.Models.CellObjects;

public class CodeCell : CellObjectBase
{
    private readonly CellObjectType _objectType;
    
    public override CellObjectType ObjectType => _objectType;
    
    public string Code { get; set; }
    public string Language { get; }
    
    public override string? DisplayValue => $"💻 {Language} ({Code?.Length ?? 0} chars)";
    
    public CodeCell(CellObjectType objectType, string language, string code) : base(code)
    {
        _objectType = objectType;
        Language = language;
        Code = code ?? string.Empty;
    }
    
    public override bool IsValid() => Code != null;
    
    public override IEnumerable<string> GetAvailableOperations()
    {
        yield return "CODE_EXECUTE";
        yield return "CODE_FORMAT";
        yield return "CODE_LINT";
        yield return "CODE_ANALYZE";
        yield return "CODE_MINIFY";
        yield return "CODE_BEAUTIFY";
        
        if (Language == "Python")
        {
            yield return "PYTHON_EXECUTE";
            yield return "PYTHON_DEPLOY";
        }
    }
    
    public override ICellObject Clone() => new CodeCell(_objectType, Language, Code);
    
    public static CodeCell CreatePython(string code) => new(CellObjectType.CodePython, "Python", code);
    public static CodeCell CreateCSharp(string code) => new(CellObjectType.CodeCSharp, "C#", code);
    public static CodeCell CreateJavaScript(string code) => new(CellObjectType.CodeJavaScript, "JavaScript", code);
    public static CodeCell CreateTypeScript(string code) => new(CellObjectType.CodeTypeScript, "TypeScript", code);
    public static CodeCell CreateCss(string code) => new(CellObjectType.CodeCss, "CSS", code);
    public static CodeCell CreateHtml(string code) => new(CellObjectType.CodeHtml, "HTML", code);
    public static CodeCell CreateSql(string code) => new(CellObjectType.CodeSql, "SQL", code);
}
