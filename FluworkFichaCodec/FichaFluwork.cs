using System;
using System.Collections.Generic;

namespace FluworkFichaCodec
{
    public sealed class FichaFluwork
    {
        public string RutaArchivo { get; set; }
        public string Ficha { get; set; }
        public string FechaIngreso { get; set; }
        public string FechaEgreso { get; set; }
        public string Estado { get; set; }
        public string NombreCompleto { get; set; }
        public string Telefono { get; set; }
        public string Adjuntos { get; set; }
        public string Problema { get; set; }
        public string SolucionRaw { get; set; }
        public string Presupuesto { get; set; }
        public string Precio { get; set; }
        public string AtendidoPor { get; set; }
        public string Tecnico { get; set; }
        public string Modelo { get; set; }
        public string NumeroSerie { get; set; }
        public string Direccion { get; set; }
        public string Email { get; set; }
        public string LlamarElDia { get; set; }
        public string ControladoPor { get; set; }
        public string AvisadoElDia { get; set; }
        public string AvisadoPor { get; set; }
        public string Confirmacion { get; set; }
        public IReadOnlyList<TrabajoRealizado> Trabajos { get; set; }

        public FichaFluwork()
        {
            Trabajos = Array.Empty<TrabajoRealizado>();
        }
    }
}
