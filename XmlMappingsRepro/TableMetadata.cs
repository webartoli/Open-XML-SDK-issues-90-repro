using System.Collections.Generic;
using XmlMappingsRepro.ColumnTypes;

namespace XmlMappingsRepro
{
    public class TableMetadata
    {
        public string Name { get; }

        private readonly List<ColumnMetadata> _headers;

        public TableMetadata(string name)
        {
            Name = name;
            _headers = new List<ColumnMetadata>();
        }

        public IEnumerable<ColumnMetadata> Headers => _headers;

        public TableMetadata Column(string name, IColumnType type)
        {
            return Column(new ColumnMetadata
            {
                Name = name,
                Type = type
            });
        }


        public TableMetadata Column(ColumnMetadata column)
        {
            column.Index = _headers.Count;
            _headers.Add(column);
            return this;
        }
    }
}