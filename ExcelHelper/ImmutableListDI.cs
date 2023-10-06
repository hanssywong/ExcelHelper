using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHelper
{
    public static class ImmutableListDI
    {
        public static DataTable ToDataTable<T>(this ImmutableList<T> items, string tablename) where T : class
        {
            DataTable table = new(tablename);
            foreach(var item in items)
            {
                var properties = item.GetType().GetProperties();
                if (table.Columns.Count == 0)
                {
                    foreach (var property in properties)
                    {
                        table.Columns.Add(property.Name, Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType);
                    }
                }
                DataRow row = table.NewRow();
                foreach (var property in properties)
                {
                    row[property.Name] = property.GetValue(item) ?? DBNull.Value;
                }
                table.Rows.Add(row);
            }
            return table;
        }
    }
}
