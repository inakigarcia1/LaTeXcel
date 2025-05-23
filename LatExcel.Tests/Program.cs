using System.Diagnostics;
using System.Text.Json;
using System.Text.Json.Serialization;
using LatExcel.Aplicacion;
using LatExcel.Aplicacion.DTOs;

namespace LatExcel.Tests;

internal class Program
{
    static void Main(string[] args)
    {
        string json = @"{
  ""objetivo"": {
    ""funcion"": {
      ""X"": 5,
      ""Y"": -3,
      ""Z"": 2.5
    },
    ""tipoObjetivo"": ""Max""
  },
  ""restricciones"": [
    {
      ""ladoIzquierdo"": { ""X"": 1, ""Y"": 2 },
      ""tipoRestriccion"": ""MenorIgual"",
      ""ladoDerecho"": 10
    },
    {
      ""ladoIzquierdo"": { ""X"": 3, ""Z"": 1 },
      ""tipoRestriccion"": ""MayorIgual"",
      ""ladoDerecho"": 5
    },
    {
      ""ladoIzquierdo"": { ""Y"": -1, ""Z"": 4 },
      ""tipoRestriccion"": ""Igual"",
      ""ladoDerecho"": 7.5
    }
  ]
}";

        string json2 = @"{
  ""objetivo"": {
    ""funcion"": {
      ""A"": 10.5,
      ""B"": 0,
      ""C"": -2
    },
    ""tipoObjetivo"": ""Min""
  },
  ""restricciones"": [
    {
      ""ladoIzquierdo"": { ""A"": 1, ""B"": 1, ""C"": 1 },
      ""tipoRestriccion"": ""Igual"",
      ""ladoDerecho"": 100
    },
    {
      ""ladoIzquierdo"": { ""A"": 0.5, ""C"": -3 },
      ""tipoRestriccion"": ""MenorIgual"",
      ""ladoDerecho"": 20
    }
  ]
}";

        string json3 = @"{
  ""objetivo"": {
    ""funcion"": {
      ""W"": 1.1,
      ""X"": -4.4,
      ""Y"": 3.3,
      ""Z"": 2.2
    },
    ""tipoObjetivo"": ""Eq""
  },
  ""restricciones"": [
    {
      ""ladoIzquierdo"": { ""W"": 1, ""X"": 1 },
      ""tipoRestriccion"": ""MayorIgual"",
      ""ladoDerecho"": 8
    },
    {
      ""ladoIzquierdo"": { ""Y"": -2, ""Z"": 5 },
      ""tipoRestriccion"": ""MenorIgual"",
      ""ladoDerecho"": 15.5
    }
  ]
}";


        var opciones = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true,
            Converters = { new JsonStringEnumConverter() }
        };

        // Ejemplo con el primero
        var modelo = JsonSerializer.Deserialize<Modelo>(json2, opciones);

        if (modelo is null) return;

        var bytes = GeneradorDeExcel.GenerarArchivoExcel(modelo);
        File.WriteAllBytes("modelo1.xlsx", bytes);

        var ruta = Path.Combine(Directory.GetCurrentDirectory());
        Process.Start("explorer", ruta);
    }
}
