using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System.Text.Json;
using GenerateSpreadsheets.Models;
using Newtonsoft.Json;

namespace GenerateSpreadsheets.Controllers
{
    [Route("[controller]")]
    public class Spreadsheets : Controller
    {
        private readonly ILogger<Spreadsheets> _logger;

        public Spreadsheets(ILogger<Spreadsheets> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost("/get-spreadsheet")]
        public ActionResult ConvertToSpreadsheet([FromBody] dynamic entity)
        {
            DataTable internalEntity = (DataTable)JsonConvert.DeserializeObject(entity, new JsonSerializerSettings());
            //var respData = jsonToCSV(entity, ",");
            using (var outFile = new StreamWriter(@$"c:\temp\myfile.csv"))
            {
                foreach (TimeReport str in entity)
                {
                    outFile.WriteLine($"{str.DepartmentName},{str.StartDate},{str.EndDate}");
                }
            }
            return Ok();
        }
        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View("Error!");
        }

        private static string jsonToCSV(string jsonContent, string delimiter)
        {
            StringWriter csvString = new StringWriter();
            using (var csv = new CsvHelper.CsvWriter(csvString, System.Globalization.CultureInfo.CurrentCulture, false))
            {
                using (var dt = jsonStringToTable(jsonContent))
                {
                    foreach (DataColumn column in dt.Columns)
                    {
                        csv.WriteField(column.ColumnName);
                    }
                    csv.NextRecord();

                    foreach (DataRow row in dt.Rows)
                    {
                        for (var i = 0; i < dt.Columns.Count; i++)
                        {
                            csv.WriteField(row[i]);
                        }
                        csv.NextRecord();
                    }
                }
            }
            return csvString.ToString();
        }

        private static DataTable jsonStringToTable(string jsonContent)
        {
            DataTable dt = Newtonsoft.Json.JsonConvert.DeserializeObject<DataTable>(jsonContent);
            return dt;
        }
    }
}