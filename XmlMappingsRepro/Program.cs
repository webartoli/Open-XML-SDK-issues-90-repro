using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XmlMappingsRepro.ColumnTypes;

namespace XmlMappingsRepro
{
    class Program
    {
        static void Main(string[] args)
        {
            GenerateFile("buggy", x => x);
            GenerateFile("expected", x => new OpenXmlElementNamespaceInjector<MapInfo>(x).CopyToDefault("x").GetInstance());
        }

        private static void GenerateFile(string filename, Func<MapInfo, MapInfo> applyHack)
        {
            var template = new FileInfo(AssetFile("brands-file.xlsx"));
            var fileInfo = template.CopyTo($"{filename}.xlsx", true);

            using (var document = SpreadsheetDocument.Open(fileInfo.FullName, true))
            {
                var workbook = document.WorkbookPart;
                var worksheet = workbook.WorksheetParts.First();

                var xsd = File.ReadAllText(AssetFile("brands-definition.xsd"));
                new XmlExcelMappingsWriter(Metadata(), xsd, applyHack).Execute(document, worksheet);
            }
        }

        private static string AssetFile(string filename)
        {
            return Path.Combine("..", "..","..", "assets", filename);
        }

        private static TableMetadata Metadata()
        {
            return new TableMetadata("Brands")
                .Column(new ColumnMetadata
                {
                    Name = "Brand Id",
                    Type = new IntegerColumnType()
                })
                .Column(new ColumnMetadata
                {
                    Name = "Brand",
                    Type = new StringColumnType()
                });
        }
    }
}
