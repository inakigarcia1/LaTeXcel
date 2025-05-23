using LatExcel.Aplicacion.Modelo;
using System.Drawing;
using System.Runtime.InteropServices;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace LatExcel.Aplicacion;
public static class GeneradorDeExcel
{
    private static readonly List<ExcelRange> CeldasConOperacion = [];
    private static readonly Color ColorCeldasOpcionales = Color.LightGreen;
    private static readonly Color ColorCeldasObligatorias = Color.LightSalmon;
    public static byte[] GenerarArchivoExcel(DTOs.Modelo modelo)
    {
        ExcelPackage.License.SetNonCommercialPersonal("Inaki");

        using var paquete = new ExcelPackage();

        var hoja = paquete.Workbook.Worksheets.Add("Modelo");

        var todasLasVariables = ObtenerTodasLasVariables(modelo);

        AgregarEncabezadosRestricciones(hoja, todasLasVariables);
        var fila = AgregarRestricciones(modelo, hoja, todasLasVariables);
        fila += 2; // Espacio
        fila = CrearFuncionObjetivo(modelo, hoja, fila, todasLasVariables);
        fila += 1; // Espacio
        RellenarColumnasVariable(1, fila, hoja, todasLasVariables);
        fila++;
        var rangoVariables = RellenarColumnas(1, fila, hoja, todasLasVariables.Count);
        AgregarOperaciones(rangoVariables);
        return paquete.GetAsByteArray();
    }

    private static void AgregarOperaciones((string inicial, string final) rangoVariables)
    {
        foreach (var celda in CeldasConOperacion)
        {
            celda.Formula += $"{rangoVariables.inicial}:{rangoVariables.final})";
        }
    }

    private static void DarEstiloACelda(ExcelRange celda, Color color, bool negrita = false)
    {
        PintarCelda(celda, color);
        PonerBordesACelda(celda);
        if (negrita)
            PonerNegritaACelda(celda);
    }

    private static void PonerNegritaACelda(ExcelRange celda)
    {
        celda.Style.Font.Bold = true;
    }

    private static void PintarCelda(ExcelRange celda, Color color)
    {
        celda.Style.Fill.PatternType = ExcelFillStyle.Solid;
        celda.Style.Fill.BackgroundColor.SetColor(color);
    }

    private static int CrearFuncionObjetivo(DTOs.Modelo modelo, ExcelWorksheet hoja, int fila, List<string> todasLasVariables)
    {
        var objetivo = modelo.Objetivo.TipoObjetivo == TipoObjetivo.Max ? "Max" : "Min";
        var celda = hoja.Cells[fila, 1];
        celda.Value = $"Función Objetivo {objetivo} Z";

        fila++;

        var columna = RellenarColumnas(1, fila, hoja, todasLasVariables);
        DarEstiloACelda(hoja.Cells[fila, columna], ColorCeldasOpcionales);

        hoja.Cells[fila, columna].Value = "Valor Z";
        DarEstiloACelda(hoja.Cells[fila, columna], ColorCeldasOpcionales, true);

        fila++;

        columna = 1;
        var primeraCelda = hoja.Cells[fila, columna];

        foreach (var variable in todasLasVariables)
        {
            DarEstiloACelda(hoja.Cells[fila, columna], ColorCeldasObligatorias);
            hoja.Cells[fila, columna++].Value = modelo.Objetivo.Funcion.GetValueOrDefault(variable, 0);
        }

        var ultimaCelda = hoja.Cells[fila, columna - 1];

        // TODO: Agregar lo faltante despues de la coma de SUMPRODUCT
        hoja.Cells[fila, columna].Formula = $"SUMPRODUCT({primeraCelda.Address}:{ultimaCelda.Address}, ";
        DarEstiloACelda(hoja.Cells[fila, columna], ColorCeldasObligatorias);
        CeldasConOperacion.Add(hoja.Cells[fila, columna]);
        fila++;

        return fila;
    }

    private static int AgregarRestricciones(DTOs.Modelo modelo, ExcelWorksheet hoja, List<string> todasLasVariables)
    {
        var fila = 3;
        var numeroRestriccion = 1;

        foreach (var restriccion in modelo.Restricciones)
        {
            var columna = 1;
            DarEstiloACelda(hoja.Cells[fila, columna], ColorCeldasOpcionales, true);
            hoja.Cells[fila, columna++].Value = $"R{numeroRestriccion++}";

            var primeraCelda = hoja.Cells[fila, columna];

            foreach (var variable in todasLasVariables)
            {
                restriccion.LadoIzquierdo.TryGetValue(variable, out double coef);
                DarEstiloACelda(hoja.Cells[fila, columna], ColorCeldasObligatorias);
                hoja.Cells[fila, columna++].Value = coef;
            }

            var ultimaCelda = hoja.Cells[fila, columna - 1];

            // TODO: Agregar lo faltante despues de la coma de SUMPRODUCT
            DarEstiloACelda(hoja.Cells[fila, columna], ColorCeldasObligatorias);
            hoja.Cells[fila, columna++].Formula = $"SUMPRODUCT({primeraCelda.Address}:{ultimaCelda.Address}, ";
            CeldasConOperacion.Add(hoja.Cells[fila, columna - 1]);

            DarEstiloACelda(hoja.Cells[fila, columna], ColorCeldasOpcionales);
            hoja.Cells[fila, columna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            hoja.Cells[fila, columna].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            hoja.Cells[fila, columna++].Value = TraducirTipoRestriccion(restriccion.TipoRestriccion);
            DarEstiloACelda(hoja.Cells[fila, columna], ColorCeldasObligatorias);
            hoja.Cells[fila, columna].Value = restriccion.LadoDerecho;
            fila++;
        }

        return fila;
    }

    private static void AgregarEncabezadosRestricciones(ExcelWorksheet hoja, List<string> todasLasVariables)
    {
        hoja.Cells["A2"].Value = "Restricciones";
        hoja.Cells["A2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
        hoja.Cells["A2"].Style.Fill.BackgroundColor.SetColor(ColorCeldasOpcionales);
        PonerBordesACelda(hoja.Cells["A2"]);
        PonerNegritaACelda(hoja.Cells["A2"]);

        var columna = RellenarColumnas(2, 2, hoja, todasLasVariables);

        DarEstiloACelda(hoja.Cells[2, columna], ColorCeldasOpcionales, true);
        hoja.Cells[2, columna++].Value = "LI";

        DarEstiloACelda(hoja.Cells[2, columna], ColorCeldasOpcionales, true);
        hoja.Cells[2, columna++].Value = "Relación";

        DarEstiloACelda(hoja.Cells[2, columna], ColorCeldasOpcionales, true);
        hoja.Cells[2, columna].Value = "LD";
    }

    private static int RellenarColumnas(int columnaInicial, int fila, ExcelWorksheet hoja, List<string> variables)
    {
        var columna = columnaInicial;
        foreach (var variable in variables)
        {
            DarEstiloACelda(hoja.Cells[fila, columna], ColorCeldasOpcionales, true);
            hoja.Cells[fila, columna++].Value = "c" + variable;
        }
        return columna;
    }

    private static void RellenarColumnasVariable(int columnaInicial, int fila, ExcelWorksheet hoja, List<string> variables)
    {
        var columna = columnaInicial;
        foreach (var variable in variables)
        {
            DarEstiloACelda(hoja.Cells[fila, columna], ColorCeldasOpcionales, true);
            hoja.Cells[fila, columna++].Value = variable;
        }
    }

    private static (string inicial, string final) RellenarColumnas(int columnaInicial, int fila, ExcelWorksheet hoja, int cantidadColumnas)
    {
        var primeraCelda = hoja.Cells[fila, columnaInicial].AddressAbsolute;
        for (int i = 0; i < cantidadColumnas; i++)
        {
            DarEstiloACelda(hoja.Cells[fila, columnaInicial], ColorCeldasObligatorias);
            hoja.Cells[fila, columnaInicial++].Value = 0;
        }
        var ultimaCelda = hoja.Cells[fila, columnaInicial - 1].AddressAbsolute;
        return (primeraCelda, ultimaCelda);
    }

    private static void PonerBordesACelda(ExcelRange celda)
    {
        var border = celda.Style.Border;
        border.Top.Style = ExcelBorderStyle.Thin;
        border.Bottom.Style = ExcelBorderStyle.Thin;
        border.Left.Style = ExcelBorderStyle.Thin;
        border.Right.Style = ExcelBorderStyle.Thin;
    }

    private static List<string> ObtenerTodasLasVariables(DTOs.Modelo modelo)
    {
        var variables = new HashSet<string>();

        foreach (var variable in modelo.Restricciones.SelectMany(restriccion => restriccion.LadoIzquierdo.Keys))
            variables.Add(variable);

        foreach (var variable in modelo.Objetivo.Funcion.Keys)
            variables.Add(variable);

        return variables.OrderBy(v => v).ToList();
    }

    private static string TraducirTipoRestriccion(TipoRestriccion tipo)
    {
        return tipo switch
        {
            TipoRestriccion.MenorIgual => "<=",
            TipoRestriccion.MayorIgual => ">=",
            TipoRestriccion.Igual => "=",
            _ => "?"
        };
    }
}
