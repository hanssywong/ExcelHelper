using Serilog;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHelper
{
    public class BaseExcel
    {
        public BaseExcel(string path, ILogger logger)
        {
            Logger = logger;
            Path = path;
            Wb.Worksheets.Clear();
        }
        ~BaseExcel()
        {
            Wb.Dispose();
        }
        public ILogger? Logger { get; protected set; }
        public string Path { get; protected set; }
        public Workbook Wb { get; protected set; } = new();
        public enum Region
        {
            MY,
            SG
        }
        public void AddSheet<T>(string sheetname, ImmutableList<T> data) where T : class
        {
            Worksheet? sheet1 = Wb.Worksheets[sheetname];
            if (sheet1==null)
            {
                sheet1 = Wb.Worksheets.Add(sheetname);
            }
            sheet1.InsertDataTable(data.ToDataTable(sheetname), true, 1, 1);
        }
        public void Publish2Excel()
        {
            Wb.SaveToFile(Path, ExcelVersion.Version2013);
        }
        public void ReadFromExcel()
        {
            Wb.LoadFromFile(Path);
        }
    }
}
