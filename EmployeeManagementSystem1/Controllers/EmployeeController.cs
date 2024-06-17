using EmployeeManagementSystem.Models;
using EmployeeManagementSystem.Services;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace EmployeeManagementSystem.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class EmployeeController : ControllerBase
    {
        private readonly ICosmosDbService _cosmosDbService;

        public EmployeeController(ICosmosDbService cosmosDbService)
        {
            _cosmosDbService = cosmosDbService;
        }

        
        [HttpGet]
        public async Task<IEnumerable<EmployeeBasicDetails>> Get()
        {
            return await _cosmosDbService.GetEmployeesAsync("SELECT * FROM c");
        }

        
        [HttpGet("{id}")]
        public async Task<ActionResult<EmployeeBasicDetails>> Get(string id)
        {
            var employee = await _cosmosDbService.GetEmployeeAsync(id);
            if (employee == null)
            {
                return NotFound();
            }
            return employee;
        }

        
        [HttpPost]
        public async Task<ActionResult> Post([FromBody] EmployeeBasicDetails employee)
        {
            await _cosmosDbService.AddEmployeeAsync(employee);
            return CreatedAtAction(nameof(Get), new { id = employee.EmployeeID }, employee);
        }

        
        [HttpPut("{id}")]
        public async Task<ActionResult> Put(string id, [FromBody] EmployeeBasicDetails employee)
        {
            var existingEmployee = await _cosmosDbService.GetEmployeeAsync(id);
            if (existingEmployee == null)
            {
                return NotFound();
            }

            await _cosmosDbService.UpdateEmployeeAsync(id, employee);
            return NoContent();
        }

        
        [HttpDelete("{id}")]
        public async Task<ActionResult> Delete(string id)
        {
            var employee = await _cosmosDbService.GetEmployeeAsync(id);
            if (employee == null)
            {
                return NotFound();
            }

            await _cosmosDbService.DeleteEmployeeAsync(id);
            return NoContent();
        }

        
        [HttpPost("import")]
        public async Task<IActionResult> Import(IFormFile file)
        {
            if (file == null || file.Length <= 0)
            {
                return BadRequest("Invalid file.");
            }

            using (var stream = new MemoryStream())
            {
                await file.CopyToAsync(stream);
                using (var package = new ExcelPackage(stream))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                    {
                        var employee = new EmployeeBasicDetails
                        {
                            FirstName = worksheet.Cells[row, 2].Value?.ToString(),
                            LastName = worksheet.Cells[row, 3].Value?.ToString(),
                            Email = worksheet.Cells[row, 4].Value?.ToString(),
                            Mobile = worksheet.Cells[row, 5].Value?.ToString(),
                            ReportingManagerName = worksheet.Cells[row, 6].Value?.ToString(),
                            Address = new Address(),
                        };

                        await _cosmosDbService.AddEmployeeAsync(employee);
                    }
                }
            }

            return Ok();
        }

        
        [HttpGet("export")]
        public async Task<IActionResult> Export()
        {
            var employees = await _cosmosDbService.GetEmployeesAsync("SELECT * FROM c");
            var stream = new MemoryStream();

            using (var package = new ExcelPackage(stream))
            {
                var worksheet = package.Workbook.Worksheets.Add("Employees");
                worksheet.Cells[1, 1].Value = "Sr.No";
                worksheet.Cells[1, 2].Value = "First Name";
                worksheet.Cells[1, 3].Value = "Last Name";
                worksheet.Cells[1, 4].Value = "Email";
                worksheet.Cells[1, 5].Value = "Phone No";
                worksheet.Cells[1, 6].Value = "Reporting Manager Name";
                worksheet.Cells[1, 7].Value = "Date Of Birth";
                worksheet.Cells[1, 8].Value = "Date Of Joining";

                int row = 2;
                int serialNo = 1;
                foreach (var employee in employees)
                {
                    worksheet.Cells[row, 1].Value = serialNo++;
                    worksheet.Cells[row, 2].Value = employee.FirstName;
                    worksheet.Cells[row, 3].Value = employee.LastName;
                    worksheet.Cells[row, 4].Value = employee.Email;
                    worksheet.Cells[row, 5].Value = employee.Mobile;
                    worksheet.Cells[row, 6].Value = employee.ReportingManagerName;
                    worksheet.Cells[row, 7].Value = employee.Address?.Street; 
                    worksheet.Cells[row, 8].Value = employee.Address?.City; 
                    row++;
                }

                package.Save();
            }

            stream.Position = 0;
            string excelName = $"EmployeeList-{DateTime.Now:yyyyMMddHHmmssfff}.xlsx";
            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", excelName);
        }

        
        [HttpGet("paged")]
        public async Task<IActionResult> GetPaged([FromQuery] int pageNumber = 1, [FromQuery] int pageSize = 10)
        {
            var employees = await _cosmosDbService.GetEmployeesAsync("SELECT * FROM c");
            var pagedEmployees = employees.Skip((pageNumber - 1) * pageSize).Take(pageSize);

            var totalRecords = employees.Count();
            var totalPages = (int)Math.Ceiling(totalRecords / (double)pageSize);

            var paginationMetadata = new
            {
                totalRecords,
                totalPages,
                pageSize,
                currentPage = pageNumber
            };

            Response.Headers.Add("X-Pagination", Newtonsoft.Json.JsonConvert.SerializeObject(paginationMetadata));

            return Ok(pagedEmployees);
        }

        
        [HttpGet("search")]
        public async Task<IActionResult> Search([FromQuery] string term)
        {
            var query = $"SELECT * FROM c WHERE CONTAINS(c.firstName, '{term}') OR CONTAINS(c.lastName, '{term}')";
            var employees = await _cosmosDbService.GetEmployeesAsync(query);
            return Ok(employees);
        }
    }
}