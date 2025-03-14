using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataHandler
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelColumnAttribute : Attribute
    {
        public string ColumnName { get; }

        public ExcelColumnAttribute(string columnName)
        {
            ColumnName = columnName;
        }
    }
}
