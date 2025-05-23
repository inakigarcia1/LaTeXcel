using System.Text;

namespace LatExcel.Aplicacion.Modelo;
public class Restriccion
{
    public Dictionary<string, double> LadoIzquierdo { get; set; } = new();
    public TipoRestriccion TipoRestriccion { get; set; }
    public double LadoDerecho { get; set; }

    public override string ToString()
    {
        var funcion = new StringBuilder();
        foreach (var variable in LadoIzquierdo)
        {
            funcion.Append($"{variable.Value} * {variable.Key} + ");
        }
        funcion.Remove(funcion.Length - 3, 3); // Eliminar el último " + "

        var restriccion = TipoRestriccion switch
        {
            TipoRestriccion.Igual => "=",
            TipoRestriccion.MayorIgual => ">=",
            TipoRestriccion.MenorIgual => "<=",
            _ => ""
        };

        funcion.Append($" {restriccion} {LadoDerecho}");
        return funcion.ToString();
    }
}
