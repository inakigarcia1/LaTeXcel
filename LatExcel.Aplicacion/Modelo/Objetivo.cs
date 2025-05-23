using System.Text;

namespace LatExcel.Aplicacion.Modelo;
public class Objetivo
{
    public Dictionary<string, double> Funcion { get; set; } = new();
    public TipoObjetivo TipoObjetivo { get; set; }

    public override string ToString()
    {
        return $"Z {TipoObjetivo} = {MostrarFuncion()}";
    }

    private string MostrarFuncion()
    {
        var funcion = new StringBuilder();
        foreach (var variable in Funcion)
        {
            funcion.Append($"{variable.Value} * {variable.Key} + ");
        }
        funcion.Remove(funcion.Length - 3, 3); // Eliminar el último " + "
        return funcion.ToString();
    }
}
