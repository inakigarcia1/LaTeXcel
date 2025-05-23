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
        //System.IO.File.WriteAllBytes("modelo.xlsx", bytes);
        //string ruta = Path.Combine(Directory.GetCurrentDirectory(), "bin", "Debug", "net8.0");
        //Process.Start("explorer", ruta);
        return Ok(modelo);
    }
}
