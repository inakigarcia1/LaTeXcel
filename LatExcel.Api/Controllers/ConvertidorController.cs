using System.Diagnostics;
using LatExcel.Aplicacion;
using LatExcel.Aplicacion.DTOs;
using Microsoft.AspNetCore.Mvc;

namespace LatExcel.Api.Controllers;
[ApiController]
[Route("plapi")]
public class ConvertidorController : ControllerBase
{
    [HttpPost("convertir")]
    public IActionResult Convertir([FromBody] Modelo modelo)
    {
        var bytes = GeneradorDeExcel.GenerarArchivoExcel(modelo);
        return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Modelo_Solver.xlsx");
    }
}
