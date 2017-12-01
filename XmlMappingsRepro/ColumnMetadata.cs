using XmlMappingsRepro.ColumnTypes;

namespace XmlMappingsRepro
{
    public class ColumnMetadata
    {
        public string Name { get; set; }
        public IColumnType Type { get; set; }
        public int Index { get; set; }

        public override string ToString()
        {
            return Name;
        }
    }
}