using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XmlMappingsRepro.ColumnTypes;

namespace XmlMappingsRepro
{
    public class XmlExcelMappingsWriter
    {
        private const uint ConnectionId = 1;
        private const string ConnectionName = "ConnectionName";

        private const uint MapId = 1;

        private const string RootElementName = "Set";
        private const string RowElementName = "Row";
        private const string SchemaId = "Schema1";

        private readonly TableMetadata _metadata;
        private readonly string _xmlSchema;
        private readonly Func<MapInfo, MapInfo> _applyHack;


        public XmlExcelMappingsWriter(TableMetadata metadata, string xmlSchema, Func<MapInfo, MapInfo> applyHack)
        {
            _metadata = metadata;
            _xmlSchema = xmlSchema;
            _applyHack = applyHack;
        }

        public void Execute(SpreadsheetDocument document, WorksheetPart worksheet)
        {
            SetTableProperties(worksheet);
            SetConnection(document);
            SetCustomXmlMappings(document);
        }

        private void SetConnection(SpreadsheetDocument document)
        {
            var connectionPart = document.WorkbookPart.AddNewPart<ConnectionsPart>(NewId());
            var connections = new Connections();

            var connection = new Connection
            {
                Id = ConnectionId,
                Name = ConnectionName,
                Type = 4U,
                RefreshedVersion = 0,
                Background = true
            };

            connection.AppendChild(new WebQueryProperties
            {
                XmlSource = true,
                SourceData = true,
                ParsePreTag = true,
                Consecutive = true,
                Url = "C:\\file.xml",
                HtmlTables = true
            });
            connections.AppendChild(connection);
            connectionPart.Connections = connections;
        }

        private static string NewId()
        {
            return "ID" + Guid.NewGuid().ToString("N");
        }

        private void SetTableProperties(WorksheetPart worksheet)
        {
            var tableDefs = worksheet.TableDefinitionParts.First(x => x.Table.Name == _metadata.Name);

            tableDefs.Table.TableType = TableValues.Xml;
            tableDefs.Table.ConnectionId = ConnectionId;
            tableDefs.Table.TotalsRowShown = false;


            foreach (var tableColumn in tableDefs.Table.TableColumns.OfType<TableColumn>())
            {
                var columnMetadata = _metadata.Headers.First(x => x.Index + 1 == tableColumn.Id.Value);

                var escapedColumnName = Escape(tableColumn);

                tableColumn.UniqueName = escapedColumnName;
                tableColumn.AppendChild(new XmlColumnProperties
                {
                    MapId = MapId,
                    XPath = $"/{RootElementName}/{RowElementName}/{escapedColumnName}",
                    XmlDataType = XmlDataTypeFor(columnMetadata)
                });
            }
        }

        private static XmlDataValues XmlDataTypeFor(ColumnMetadata columnMetadata)
        {
            switch (columnMetadata.Type)
            {
                case IntegerColumnType _:
                    return XmlDataValues.Integer;
                case StringColumnType _:
                    return XmlDataValues.String;
            }
            return XmlDataValues.String;
        }

        private static string Escape(TableColumn tableColumn)
        {
            return tableColumn.Name.Value.Replace(" ", "_x005F_x0020_");
        }

        private void SetCustomXmlMappings(SpreadsheetDocument document)
        {
            var customXmlMappingsPart = document.WorkbookPart.AddNewPart<CustomXmlMappingsPart>(NewId());
            var mapInfo = new MapInfo { SelectionNamespaces = "" };

            var schema = new Schema { Id = SchemaId };

            var xmlSchemaElement = OpenXmlUnknownElement.CreateOpenXmlUnknownElement(_xmlSchema);

            schema.AppendChild(xmlSchemaElement);

            var map = new Map
            {
                ID = MapId,
                Name = _metadata.Name,
                RootElement = RootElementName,
                SchemaId = SchemaId,
                ShowImportExportErrors = false,
                AutoFit = true,
                AppendData = false,
                PreserveAutoFilterState = true,
                PreserveFormat = true
            };

            var dataBinding = new DataBinding
            {
                FileBinding = true,
                ConnectionId = ConnectionId,
                DataBindingLoadMode = 1U
            };

            map.AppendChild(dataBinding);
            mapInfo.AppendChild(schema);
            mapInfo.AppendChild(map);

            customXmlMappingsPart.MapInfo = _applyHack(mapInfo);
        }
    }
}