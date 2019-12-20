using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using EPPlusExport.Models;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;

namespace EPPlusExport.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ValuesController : ControllerBase
    {
        // GET api/values
        [HttpGet]
        public ActionResult<IEnumerable<string>> Get()
        {
            return new string[] { "value1", "value2" };
        }

        // GET api/values/5
        [HttpGet("{id}")]
        public ActionResult<string> Get(int id)
        {
            return "value";
        }

        // POST api/values
        [HttpPost]
        public void Post([FromBody] string value)
        {
        }

        // PUT api/values/5
        [HttpPut("{id}")]
        public void Put(int id, [FromBody] string value)
        {
        }

        // DELETE api/values/5
        [HttpDelete("{id}")]
        public void Delete(int id)
        {
        }

        [HttpGet]
        [Route("ExportCustomer")]
        public IActionResult ExportCustomer()
        {
            ExcelPackage package = new ExcelPackage();

            List<Customer> customerList = new List<Customer>
            {
                new Customer(1, "Alvaro Turizo", "avtu@sinco.com.co", "Ecuador"),
                new Customer(2, "Camilo Barreto", "cabe@sinco.com.co", "Argentina"),
                new Customer(3, "Luis Orjuela", "luor@sinco.com.co", "Mexico"),
                new Customer(4, "Jessica Carrillo", "jeca@sinco.com.co", "Peru"),
                new Customer(5, "Antonio Diaz", "andi@sinco.com.co", "Bolivia"),
                new Customer(6, "Nestor Castañeda", "neca@sinco.com.co", "Venezuela"),
                new Customer(7, "Arnod Chirivi", "arch@sinco.com.co", "Uruguay"),
                new Customer(8, "Felipe Diaz", "fedi@sinco.com.co", "Paraguay"),
            };

            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Customer");
            int totalRows = customerList.Count();

            worksheet.Cells[1, 1].Value = "Customer ID";
            worksheet.Cells[1, 2].Value = "Customer Name";
            worksheet.Cells[1, 3].Value = "Customer Email";
            worksheet.Cells[1, 4].Value = "customer Country";
            int i = 0;
            for (int row = 2; row <= totalRows + 1; row++)
            {
                worksheet.Cells[row, 1].Value = customerList[i].Id;
                worksheet.Cells[row, 2].Value = customerList[i].Name;
                worksheet.Cells[row, 3].Value = customerList[i].Email;
                worksheet.Cells[row, 4].Value = customerList[i].Country;
                i++;
            }

            return File(package.GetAsByteArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        }
    }
}
