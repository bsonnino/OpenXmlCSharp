using System.IO.Packaging;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

//using var package = Package.Open(@"D:\Backup\R_drive\Documentos\Artigos\Artigos\CSharp\OpenXMLDelphi\exemplo.docx", FileMode.Open, FileAccess.Read);
//package.GetParts().Select(p => p.ContentType).OrderBy(p => p).ToList().ForEach(p => Console.WriteLine(p));
//package.GetRelationships().ToList().ForEach(r => Console.WriteLine($"{r.Id} - {r.SourceUri} - {r.TargetUri} - {r.TargetMode} - {r.RelationshipType}")); 
//package.GetParts().Where(p => p.ContentType.Contains("main+xml")).ToList().ForEach(p => Console.WriteLine(p.Uri));
//package.GetParts().Where(p => p.ContentType.Contains("core-properties")).ToList().ForEach(p => WriteXmlToConsole(package.GetPart(p.Uri).GetStream()));
//package.GetParts().Where(p => p.ContentType.Contains("extended-properties")).ToList().ForEach(p => WriteXmlToConsole(package.GetPart(p.Uri).GetStream()));
//package.GetPart(new Uri("/word/document.xml", UriKind.Relative)).GetStream().CopyTo(Console.OpenStandardOutput());

//CreateDoc("hello.docx", "Hello World!");

CreateDocWithAllFonts("allfonts.docx");

// Write formatted XML to console
void WriteXmlToConsole(Stream stream)
{
    var doc = new XmlDocument();
    doc.Load(stream);
    var settings = new XmlWriterSettings
    {
        Indent = true,
        IndentChars = "  ",
        NewLineChars = "\r\n",
        NewLineHandling = NewLineHandling.Replace
    };
    using var writer = XmlWriter.Create(Console.OpenStandardOutput(), settings);
    doc.Save(writer);
}

void CreateDoc(string filepath, string message)
{
    using (WordprocessingDocument doc = WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
    {
        MainDocumentPart mainPart = doc.AddMainDocumentPart();

        mainPart.Document = new Document();
        Body body = mainPart.Document.AppendChild(new Body());
        Paragraph para = body.AppendChild(new Paragraph());
        Run run = para.AppendChild(new Run());
        run.AppendChild(new Text(message));
        para.AppendChild(new Run());
    }
}

void CreateDocWithAllFonts(string filepath)
{
    using (WordprocessingDocument doc = WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
    {
        MainDocumentPart mainPart = doc.AddMainDocumentPart();

        mainPart.Document = new Document();
        Body body = mainPart.Document.AppendChild(new Body());

        // Get all fonts available in the system
        var fonts = System.Drawing.FontFamily.Families.Select(f => f.Name).ToList();

        foreach (var font in fonts)
        {
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text(font));
            run.RunProperties = new RunProperties(new RunFonts() { Ascii = font });
            para.AppendChild(new Run());
        }
    }
}




