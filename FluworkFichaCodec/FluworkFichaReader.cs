using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace FluworkFichaCodec
{
    public sealed class FluworkFichaReader
    {
        public const int RecordLength = 3842;
        public const int TrabajosOffset = 2153;
        public const int TrabajosLength = 1024;

        private static readonly FieldLayout[] Fields =
        {
            new FieldLayout("Ficha", 10),
            new FieldLayout("FechaIngreso", 10),
            new FieldLayout("FechaEgreso", 10),
            new FieldLayout("Estado", 10),
            new FieldLayout("NombreCompleto", 50),
            new FieldLayout("Telefono", 15),
            new FieldLayout("Adjuntos", 1024),
            new FieldLayout("Problema", 1024),
            new FieldLayout("SolucionRaw", 1024),
            new FieldLayout("Presupuesto", 10),
            new FieldLayout("Precio", 10),
            new FieldLayout("AtendidoPor", 50),
            new FieldLayout("Tecnico", 50),
            new FieldLayout("Modelo", 50),
            new FieldLayout("NumeroSerie", 50),
            new FieldLayout("Direccion", 200),
            new FieldLayout("Email", 75),
            new FieldLayout("LlamarElDia", 30),
            new FieldLayout("ControladoPor", 50),
            new FieldLayout("AvisadoElDia", 30),
            new FieldLayout("AvisadoPor", 30),
            new FieldLayout("Confirmacion", 30)
        };

        public FichaFluwork Read(string path)
        {
            if (path == null)
            {
                throw new ArgumentNullException(nameof(path));
            }

            byte[] bytes = File.ReadAllBytes(path);
            FichaFluwork ficha = Parse(bytes);
            ficha.RutaArchivo = path;
            return ficha;
        }

        public FichaFluwork Parse(byte[] bytes)
        {
            if (bytes == null)
            {
                throw new ArgumentNullException(nameof(bytes));
            }

            if (bytes.Length < RecordLength)
            {
                throw new InvalidDataException("El archivo de ficha es mas corto que el formato esperado.");
            }

            int offset = 0;
            Dictionary<string, string> values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (FieldLayout field in Fields)
            {
                values[field.Name] = ReadFixedString(bytes, offset, field.Length);
                offset += field.Length;
            }

            FichaFluwork ficha = new FichaFluwork
            {
                Ficha = values["Ficha"],
                FechaIngreso = values["FechaIngreso"],
                FechaEgreso = values["FechaEgreso"],
                Estado = values["Estado"],
                NombreCompleto = values["NombreCompleto"],
                Telefono = values["Telefono"],
                Adjuntos = values["Adjuntos"],
                Problema = values["Problema"],
                SolucionRaw = values["SolucionRaw"],
                Presupuesto = values["Presupuesto"],
                Precio = values["Precio"],
                AtendidoPor = values["AtendidoPor"],
                Tecnico = values["Tecnico"],
                Modelo = values["Modelo"],
                NumeroSerie = values["NumeroSerie"],
                Direccion = values["Direccion"],
                Email = values["Email"],
                LlamarElDia = values["LlamarElDia"],
                ControladoPor = values["ControladoPor"],
                AvisadoElDia = values["AvisadoElDia"],
                AvisadoPor = values["AvisadoPor"],
                Confirmacion = values["Confirmacion"]
            };

            ficha.Trabajos = ParseTrabajos(ficha.SolucionRaw);
            return ficha;
        }

        public IReadOnlyList<TrabajoRealizado> ParseTrabajos(string solucionRaw)
        {
            string[] partes = (solucionRaw ?? string.Empty).Split('|');
            List<TrabajoRealizado> trabajos = new List<TrabajoRealizado>();

            for (int i = 0; i < 10; i++)
            {
                string trabajo = GetPart(partes, i * 2);
                string importe = GetPart(partes, (i * 2) + 1);

                if (trabajo.Length == 0 && importe.Length == 0)
                {
                    continue;
                }

                trabajos.Add(new TrabajoRealizado
                {
                    Indice = i,
                    Trabajo = trabajo,
                    Importe = importe
                });
            }

            return trabajos;
        }

        private static string GetPart(string[] partes, int index)
        {
            if (index < 0 || index >= partes.Length)
            {
                return string.Empty;
            }

            return partes[index].Trim();
        }

        private static string ReadFixedString(byte[] bytes, int offset, int length)
        {
            return DecodeLatin1(bytes, offset, length).TrimEnd('\0', ' ');
        }

        private static string DecodeLatin1(byte[] bytes, int offset, int length)
        {
            char[] chars = new char[length];
            for (int i = 0; i < length; i++)
            {
                chars[i] = (char)bytes[offset + i];
            }

            return new string(chars);
        }

        private sealed class FieldLayout
        {
            public FieldLayout(string name, int length)
            {
                Name = name;
                Length = length;
            }

            public string Name { get; private set; }
            public int Length { get; private set; }
        }
    }
}
