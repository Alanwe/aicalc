namespace AiCalc.Models;

public enum CellObjectType
{
    Empty,
    Number,
    Text,
    Boolean,
    DateTime,
    Image,
    Audio,
    Video,
    Directory,
    File,
    Table,
    Script,
    Json,
    Xml,
    Markdown,
    Link,
    Error,
    
    // PDF Types
    Pdf,
    PdfPage,
    
    // Code Types
    CodePython,
    CodeCSharp,
    CodeJavaScript,
    CodeTypeScript,
    CodeCss,
    CodeHtml,
    CodeSql,
    CodeJson,
    
    // Chart Types
    Chart,
    ChartImage,
    
    // Data Types
    Pivot,
    DataSet,
    
    // Rich Content
    RichText,
    Html
}
