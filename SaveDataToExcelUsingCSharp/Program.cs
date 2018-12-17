using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SaveDataToExcelUsingCSharp
{
    public class Product
    {
        public Product(int Id, string Name)
        {
            this.Id = Id;
            this.Name = Name;
        }

        public int Id { get; set; }

        public string Name { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            var products = new Product[]
            {
                new Product(1, "Cellphone"), new Product(2, "Laptop"),
            };
            //
            var app = new Application();
            app.Workbooks.Add();
            Worksheet worksheet = app.ActiveWorkbook.ActiveSheet;
            int row = 1;
            //
            worksheet.Cells[row, 1] = nameof(Product.Id);
            worksheet.Cells[row, 2] = nameof(Product.Name);
            row++;
            //
            foreach (var product in products)
            {
                worksheet.Cells[row, 1] = product.Id;
                worksheet.Cells[row, 2] = product.Name;
                //
                row++;
            }
            var pathToSave = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/products.xlsx";
            app.ActiveWorkbook.SaveAs(pathToSave, XlFileFormat.xlOpenXMLWorkbook);
            app.ActiveWorkbook.Close(false);
            app.Quit();
        }
    }
}
