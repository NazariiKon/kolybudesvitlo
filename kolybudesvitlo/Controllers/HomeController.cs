using Aspose.Cells;
using kolybudesvitlo.Helpers;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;


namespace kolybudesvitlo.Controllers
{
    public class HomeController : Controller
    {
        [HttpPost("create")]
        public IActionResult Create(string street)
        {
            if (string.IsNullOrWhiteSpace(street)) return BadRequest("Невірна адреса");
            var result = ExcelWorker.FindStreetInExcelTable(street);
            if (result == null)
            {
                return BadRequest("Інформацію про адресу не знайдено. Перевірьте корректність.");
            } 
            else
            {
                return Ok(result);
            }
        }
    }
}
