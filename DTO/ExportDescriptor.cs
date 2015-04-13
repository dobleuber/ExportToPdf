using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Carvajal.Cosmos.Domain.DTO
{
    public class ExportDescriptor
    {
        public string FileName { get; set; }

        public string ExportObject { get; set; }

        public List<Field> FieldList { get; private set; }

        public ExportDescriptor()
        {
            FieldList = new List<Field>();
        }

        public class Field
        {
            public string Label { get; set; }
            public string Property { get; set; }
            public string Format { get; set; }
            public ExportDescriptor DetailDescriptor { get; set; }
        }

        public void AddField(string label, string property, ExportDescriptor detailDescriptor)
        {
            FieldList.Add(new Field
            {
                Label = label,
                Property = property,
                DetailDescriptor = detailDescriptor
            });
        }

        public void AddField(string label, string property, string format = null)
        {
            FieldList.Add(new Field
            {
                Label = label,
                Property = property,
                Format = format
            });
        }
    }
}
