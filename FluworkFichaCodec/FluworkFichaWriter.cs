using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace FluworkFichaCodec
{
    public sealed class FluworkFichaWriter
    {
        public WriteTrabajosResult WriteTrabajos(string path, IEnumerable<TrabajoRealizado> trabajos, bool createBackup)
        {
            if (path == null)
            {
                throw new ArgumentNullException(nameof(path));
            }

            byte[] bytes = File.ReadAllBytes(path);
            if (bytes.Length < FluworkFichaReader.TrabajosOffset + FluworkFichaReader.TrabajosLength)
            {
                throw new InvalidDataException("El archivo de ficha es mas corto que el formato esperado.");
            }

            string block = BuildTrabajosBlock(trabajos);
            byte[] blockBytes = EncodeLatin1(block);

            string backupPath = string.Empty;
            if (createBackup)
            {
                backupPath = CreateBackup(path);
            }

            Array.Copy(blockBytes, 0, bytes, FluworkFichaReader.TrabajosOffset, FluworkFichaReader.TrabajosLength);
            File.WriteAllBytes(path, bytes);

            return new WriteTrabajosResult
            {
                Path = path,
                BackupPath = backupPath,
                BytesWritten = FluworkFichaReader.TrabajosLength
            };
        }

        public string BuildTrabajosBlock(IEnumerable<TrabajoRealizado> trabajos)
        {
            string[] values = new string[20];
            for (int i = 0; i < values.Length; i++)
            {
                values[i] = string.Empty;
            }

            if (trabajos != null)
            {
                foreach (TrabajoRealizado trabajo in trabajos)
                {
                    if (trabajo == null || trabajo.Indice < 0 || trabajo.Indice > 9)
                    {
                        continue;
                    }

                    int offset = trabajo.Indice * 2;
                    values[offset] = NormalizeFieldValue(trabajo.Trabajo);
                    values[offset + 1] = NormalizeFieldValue(trabajo.Importe);
                }
            }

            string block = string.Join("|", values);
            if (block.Length > FluworkFichaReader.TrabajosLength)
            {
                throw new InvalidDataException("El bloque de trabajos supera 1024 caracteres.");
            }

            return block.PadRight(FluworkFichaReader.TrabajosLength);
        }

        private static string NormalizeFieldValue(string value)
        {
            return (value ?? string.Empty).Replace("|", " ").Trim();
        }

        private static string CreateBackup(string path)
        {
            string directory = Path.GetDirectoryName(path) ?? string.Empty;
            string fileName = Path.GetFileName(path);
            string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string backupPath = Path.Combine(directory, fileName + ".bak_" + stamp);

            int suffix = 1;
            while (File.Exists(backupPath))
            {
                backupPath = Path.Combine(directory, fileName + ".bak_" + stamp + "_" + suffix.ToString());
                suffix++;
            }

            File.Copy(path, backupPath, false);
            return backupPath;
        }

        private static byte[] EncodeLatin1(string value)
        {
            byte[] bytes = new byte[value.Length];
            for (int i = 0; i < value.Length; i++)
            {
                char c = value[i];
                bytes[i] = c <= 255 ? (byte)c : (byte)'?';
            }

            return bytes;
        }
    }
}
