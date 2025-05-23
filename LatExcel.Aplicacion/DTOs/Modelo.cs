using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LatExcel.Aplicacion.Modelo;

namespace LatExcel.Aplicacion.DTOs;
public class Modelo
{
    public Objetivo Objetivo { get; set; }
    public List<Restriccion> Restricciones { get; set; }

    public override string ToString()
    {
        return $"{Objetivo}\n Sujeto a:\n {MostrarRestricciones()}";
    }

    private string MostrarRestricciones()
    {
        string restricciones = string.Empty;
        foreach(var restriccion in Restricciones)
        {
            restricciones += $"{restriccion}\n";
        }
        return restricciones;
    }
}
