using Dapper;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using OfficeOpenXml;
using System.ComponentModel;
using System.Data;

namespace ExcelSheetPractice.Api.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelData
    {
        //public ExcelData(string Barcode, string Description, string Price)
        //{
        //    this.Barcode = Barcode;
        //    this.Description = Description;
        //    this.Price = Price;
        //}
        public string? ItemBarCode { get; set; }
        public string? Description { get; set; }
        public string? Price { get; set; }
    }
    public class DbProducts
    {

        public string? prf_barcode { get; set; }
        public string? prf_name_ar { get; set; }
        public bool? prf_status { get; set; }
        public string? prf_price { get; set; }
    }
    public class ExcelSheetController : ControllerBase
    {
        [HttpPost("SubmitExcelSheet")]
        public async Task<IActionResult>SubmitExcelSheet(/*string filePath*/)
        {
            ////Path = "C:\\Users\\Sheen\\Desktop\\Seoudi Market full 10-16-2024 (3).xlsx"

            //#region ReadDataFromExcelSheet

            //ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            //var dataList = new List<ExcelData>();
            //var fileInfo = new FileInfo(filePath);
            //using (var package = new ExcelPackage(fileInfo))
            //{
            //    var worksheet = package.Workbook.Worksheets.FirstOrDefault(s => s.Name == "Sheet1");
            //    if (worksheet == null)
            //    {
            //        throw new Exception("Worksheet 'Sheet1' not found.");
            //    }

            //    if (worksheet.Dimension == null)
            //    {
            //        throw new Exception("The worksheet has no data.");
            //    }

            //    int rowCount = worksheet.Dimension.Rows;
            //    const int batchSize = 1000; // Adjust batch size as needed

            //    for (int startRow = 2; startRow <= rowCount; startRow += batchSize)
            //    {
            //        int endRow = Math.Min(startRow + batchSize - 1, rowCount);

            //        for (int row = startRow; row <= endRow; row++)
            //        {
            //            //Console.WriteLine(worksheet.Cells[row, 1].Value.ToString() + " " + worksheet.Cells[row, 2].Value.ToString() + " " + worksheet.Cells[row, 3].Value.ToString());
            //            dataList.Add(new ExcelData
            //            {
            //                Barcode = worksheet.Cells[row, 1].Value.ToString(),
            //                Description=worksheet.Cells[row, 2].Value.ToString(),
            //                Price=worksheet.Cells[row, 3].Value.ToString()
            //            });
            //        }
            //    }
            //}
            //#endregion

            //return Ok();
            ////return dataList;
            ///
            // Define columns
            DataTable dataTable = new DataTable("Products");
            dataTable.Columns.Add("Barcode", typeof(string));
            dataTable.Columns.Add("Name", typeof(string));
            dataTable.Columns.Add("Price", typeof(decimal));

            // Add rows
            dataTable.Rows.Add("1234567890", "Product A", 9.99m);
            dataTable.Rows.Add("0987654321", "Product B", 19.99m);
            dataTable.Rows.Add("1122334455", "Product C", 29.99m);

            var response = JsonConvert.SerializeObject(dataTable);
            return Ok (dataTable.ToString);

        }

        [HttpPost("UpsertExcelData")]
        public async Task<IActionResult> UpsertExcelData([FromBody] List<ExcelData> ValidDataTable)
        {

            #region GetDataFromDb
            string connectionString = "Data Source=41.215.243.252;Persist Security Info=True;User ID=sa;Password=Shura@123;Encrypt=True;Trust Server Certificate=True";
            var DBProducts = new List<DbProducts>();

            using (IDbConnection dbConnection = new SqlConnection(connectionString))
            {
                dbConnection.Open();
                string sqlQuery = "select prf_barcode , prf_name_ar , prf_status , prf_price from [dbo].[product_for_files] where prf_market_id = 41 and prf_barcode Is Not Null";
                DBProducts = dbConnection.Query<DbProducts>(sqlQuery).ToList();
            }
            #endregion

            //var columnNames = ValidDataTable.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToList();
            //var dataList = new List<ExcelData>();
            //dataList = ValidDataTable.AsEnumerable().Skip(1).Select(row => {
            //    ExcelData obj = new ExcelData();
            //    foreach (var columnName in columnNames)
            //    {
            //        var property = typeof(ExcelData).GetProperty(columnName);
            //        if (property != null && row[columnName] != DBNull.Value)
            //        {
            //            property.SetValue(obj, Convert.ChangeType(row[columnName], property.PropertyType));
            //        }
            //    }

            //    return obj;
            //}).ToList();
            var dataList = ValidDataTable;


            #region CollectData
            var newData = dataList.Where(d => !DBProducts.Select(s => s.prf_barcode).ToArray().Contains(d.ItemBarCode)).Select(s => new DbProducts { prf_barcode=s.ItemBarCode, prf_status=false,prf_name_ar= s.Description, prf_price=s.Price }).ToList();
            var existingData = DBProducts.Where(f => dataList.Select(x => x.ItemBarCode).Contains(f.prf_barcode)).ToList();
            var outofStock = DBProducts.Where(f => !dataList.Select(x => x.ItemBarCode).Contains(f.prf_barcode)).Select(s => new DbProducts { prf_barcode = s.prf_barcode, prf_status = false, prf_name_ar = s.prf_name_ar, prf_price = s.prf_price }).ToList();
            var DbList = new List<DbProducts>(newData);
            DbList.AddRange(existingData);
            DbList.AddRange(outofStock);
            #endregion

            
            return Ok(DbList);
        }
    }
}
