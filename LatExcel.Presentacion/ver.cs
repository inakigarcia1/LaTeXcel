using LatExcel.Aplicacion.Modelo;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace LatExcel.Aplicacion;
public static class ver
{
    public static byte[] GenerarArchivoExcel(DTOs.Modelo modelo)
    {
        ExcelPackage.License.SetNonCommercialPersonal("Inaki");

        using var paquete = new ExcelPackage();

        var hoja = paquete.Workbook.Worksheets.Add("Modelo");

        var todasLasVariables = ObtenerTodasLasVariables(modelo);

        // -------------------- Encabezado de restricciones --------------------
        hoja.Cells["A2"].Value = "Restricciones";
        //hoja.Cells["A1:C1"].Merge = true;
        //hoja.Cells["D1"].Value = "Relación";
        //hoja.Cells["E1"].Value = "Lado derecho";

        hoja.Cells["A2"].Value = "Nombre";
        int col = 2;
        foreach (var variable in todasLasVariables)
        {
            hoja.Cells[1, col].Value = "Coef. " + variable;
            hoja.Cells[2, col].Value = "c" + variable;
            col++;
        }
        hoja.Cells["D2"].Value = "Operación";
        hoja.Cells["E2"].Value = "Valor";

        // -------------------- Datos de restricciones --------------------
        int fila = 3;
        int rIndex = 1;

        foreach (var restriccion in modelo.Restricciones)
        {
            hoja.Cells[fila, 1].Value = "R" + rIndex++;

            for (int i = 0; i < todasLasVariables.Count; i++)
            {
                var variable = todasLasVariables[i];
                restriccion.LadoIzquierdo.TryGetValue(variable, out double coef);
                hoja.Cells[fila, i + 2].Value = coef;
            }

            hoja.Cells[fila, col].Value = TraducirTipoRestriccion(restriccion.TipoRestriccion);
            hoja.Cells[fila, col + 1].Value = restriccion.LadoDerecho;

            fila++;
        }

        fila++; // Espacio

        // -------------------- Función Objetivo --------------------
        hoja.Cells[fila, 1].Value = "Función Objetivo Z";
        hoja.Cells[fila, 1, fila, col + 1].Merge = true;
        hoja.Cells[fila, 1].Style.Font.Bold = true;
        hoja.Cells[fila, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
        hoja.Cells[fila, 1].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);

        fila++;
        for (int i = 0; i < todasLasVariables.Count; i++)
        {
            var variable = todasLasVariables[i];
            hoja.Cells[fila, i + 2].Value = modelo.Objetivo.Funcion.ContainsKey(variable)
                ? modelo.Objetivo.Funcion[variable]
                : 0;
        }

        hoja.Cells[fila, 1].Value = "Coeficientes";

        fila++;

        // -------------------- Variables a variar --------------------
        hoja.Cells[fila, 1].Value = "X";
        hoja.Cells[fila + 1, 1].Value = "Y";

        fila += 3;

        // -------------------- Ajuste de estilo --------------------
        hoja.Cells[hoja.Dimension.Address].AutoFitColumns();

        // Opcional: bordes
        using (var rango = hoja.Cells[1, 1, fila, col + 1])
        {
            rango.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            rango.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            rango.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            rango.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        }

        return paquete.GetAsByteArray();
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
