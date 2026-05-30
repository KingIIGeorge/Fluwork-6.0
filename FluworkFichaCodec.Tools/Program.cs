using FluworkFichaCodec;

if (args.Length == 0)
{
    Console.WriteLine("Uso:");
    Console.WriteLine("  dotnet run --project FluworkFichaCodec.Tools -- <carpeta-fichas> [cantidad]");
    return 2;
}

string root = args[0];
int limit = args.Length > 1 && int.TryParse(args[1], out int parsedLimit)
    ? parsedLimit
    : 20;

if (!Directory.Exists(root))
{
    Console.Error.WriteLine("ERROR: No existe la carpeta: " + root);
    return 1;
}

var reader = new FluworkFichaReader();
var files = Directory.EnumerateFiles(root)
    .Select(path => new FileInfo(path))
    .Where(file => IsNumericName(file.Name) && file.Length >= FluworkFichaReader.RecordLength)
    .OrderByDescending(file => int.Parse(file.Name))
    .Take(limit)
    .ToList();

Console.WriteLine("Carpeta: " + root);
Console.WriteLine("Cantidad a validar: " + files.Count);
Console.WriteLine();
Console.WriteLine("Ficha\tFechaIngreso\tEstado\tTecnico\tCliente\tTelefono\tModelo\tTrabajos");

foreach (var file in files)
{
    try
    {
        FichaFluwork ficha = reader.Read(file.FullName);
        Console.WriteLine(string.Join("\t", new[]
        {
            ficha.Ficha,
            ficha.FechaIngreso,
            ficha.Estado,
            ficha.Tecnico,
            ficha.NombreCompleto,
            ficha.Telefono,
            ficha.Modelo,
            ficha.Trabajos.Count.ToString()
        }));
    }
    catch (Exception ex)
    {
        Console.WriteLine(file.Name + "\tERROR\t" + ex.Message);
    }
}

return 0;

static bool IsNumericName(string value)
{
    if (string.IsNullOrWhiteSpace(value))
    {
        return false;
    }

    foreach (char c in value)
    {
        if (!char.IsDigit(c))
        {
            return false;
        }
    }

    return true;
}
