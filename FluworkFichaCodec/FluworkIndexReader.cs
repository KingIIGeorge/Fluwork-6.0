using System;
using System.Collections.Generic;
using System.IO;

namespace FluworkFichaCodec
{
    public sealed class FluworkIndexReader
    {
        public const int RecordLength = 225;

        private static readonly FieldLayout[] Fields =
        {
            new FieldLayout("Ficha", 10),
            new FieldLayout("NombreCompleto", 50),
            new FieldLayout("Telefono", 15),
            new FieldLayout("Modelo", 50),
            new FieldLayout("Fecha", 10),
            new FieldLayout("Estado", 10),
            new FieldLayout("Tecnico", 50),
            new FieldLayout("Confirmacion", 30)
        };

        public IReadOnlyDictionary<string, IndexFluworkRecord> ReadByFicha(string path)
        {
            if (path == null)
            {
                throw new ArgumentNullException(nameof(path));
            }

            byte[] bytes = File.ReadAllBytes(path);
            Dictionary<string, IndexFluworkRecord> records = new Dictionary<string, IndexFluworkRecord>(StringComparer.OrdinalIgnoreCase);

            int count = bytes.Length / RecordLength;
            for (int i = 0; i < count; i++)
            {
                IndexFluworkRecord record = Parse(bytes, i * RecordLength);
                if (!string.IsNullOrWhiteSpace(record.Ficha))
                {
                    records[record.Ficha] = record;
                }
            }

            return records;
        }

        public IndexFluworkRecord Parse(byte[] bytes, int offset)
        {
            if (bytes == null)
            {
                throw new ArgumentNullException(nameof(bytes));
            }

            if (offset < 0 || bytes.Length < offset + RecordLength)
            {
                throw new InvalidDataException("El buffer no contiene un registro index.dat completo.");
            }

            int current = offset;
            Dictionary<string, string> values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (FieldLayout field in Fields)
            {
                values[field.Name] = ReadFixedString(bytes, current, field.Length);
                current += field.Length;
            }

            return new IndexFluworkRecord
            {
                Ficha = values["Ficha"],
                NombreCompleto = values["NombreCompleto"],
                Telefono = values["Telefono"],
                Modelo = values["Modelo"],
                Fecha = values["Fecha"],
                Estado = values["Estado"],
                Tecnico = values["Tecnico"],
                Confirmacion = values["Confirmacion"]
            };
        }

        private static string ReadFixedString(byte[] bytes, int offset, int length)
        {
            char[] chars = new char[length];
            for (int i = 0; i < length; i++)
            {
                chars[i] = (char)bytes[offset + i];
            }

            return new string(chars).TrimEnd('\0', ' ');
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
