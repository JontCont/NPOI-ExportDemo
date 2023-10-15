// See https://aka.ms/new-console-template for more information
using System.ComponentModel;
using System.Reflection;
using App.Model.ExportDtos;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

internal class Program
{
    private static void Main(string[] args)
    {
        var data = new List<Demo>(){
            new Demo{ name = "test1", date = DateTime.Now },
            new Demo{ name = "test2", date = DateTime.Now },
        };

        ExportExcel(data);

        void ExportExcel<T>(List<T> data)
        {
            //建立excel檔案物件
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = (XSSFSheet)workbook.CreateSheet("Default");
            Type header = typeof(T);
            //製作header
            IRow headers = (XSSFRow)sheet.CreateRow(0);
            if (header.GetType().GetProperties().Any())
            {
                int index = 0;
                foreach (var propertyInfo in header.GetProperties())
                {
                    var colName = propertyInfo.GetCustomAttribute<DisplayNameAttribute>()?.DisplayName ?? propertyInfo.Name;
                    headers.CreateCell(index).SetCellValue(colName);
                    index++;
                }
            }

            if (data.Any())
            {
                foreach (var items in data.Select((row, index) => new { row, index }))
                {
                    var item = items.row;
                    IRow row = (XSSFRow)sheet.CreateRow(items.index + 1);
                    foreach (var prop in item.GetType().GetProperties().Select((data, index) => new { data, index }))
                    {
                        var values = prop.data.GetValue(item);
                        row.CreateCell(prop.index).SetCellValue(values.ToString());
                    }
                }
            }

            // 將工作簿保存到文件
            using (var fs = new FileStream("Sample.xlsx", FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }

            Console.WriteLine("Excel文件已生成。");
        }
    }
}