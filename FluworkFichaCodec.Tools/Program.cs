using FluworkFichaCodec;

if (args.Length == 0)
{
    Console.WriteLine("Uso:");
    Console.WriteLine("  dotnet run --project FluworkFichaCodec.Tools -- <carpeta-fichas> [cantidad]");
    Console.WriteLine("  dotnet run --project FluworkFichaCodec.Tools -- compare-index <carpeta-fichas> [cantidad]");
    return 2;
}

bool compareIndex = args[0].Equals("compare-index", StringComparison.OrdinalIgnoreCase);
string root = compareIndex && args.Length > 1 ? args[1] : args[0];
int limitArgIndex = compareIndex ? 2 : 1;
int limit = args.Length > limitArgIndex && int.TryParse(args[limitArgIndex], out int parsedLimit)
    ? parsedLimit
    : 20;

if (!Directory.Exists(root))
{
    Console.Error.WriteLine("ERROR: No existe la carpeta: " + root);
    return 1;
}

var reader = new FluworkFichaReader();
IReadOnlyDictionary<string, IndexFluworkRecord>? indexByFicha = null;

if (compareIndex)
{
    string indexPath = Path.Combine(root, "index.dat");
    if (!File.Exists(indexPath))
    {
        Console.Error.WriteLine("ERROR: No existe index.dat en: " + root);
        return 1;
    }

    indexByFicha = new FluworkIndexReader().ReadByFicha(indexPath);
}

var files = Directory.EnumerateFiles(root)
    .Select(path => new FileInfo(path))
    .Where(file => IsNumericName(file.Name) && file.Length >= FluworkFichaReader.RecordLength)
    .OrderByDescending(file => int.Parse(file.Name))
    .Take(limit)
    .ToList();

Console.WriteLine("Carpeta: " + root);
Console.WriteLine("Cantidad a validar: " + files.Count);
if (compareIndex)
{
    Console.WriteLine("Modo: comparacion contra index.dat");
}
Console.WriteLine();
Console.WriteLine(compareIndex
    ? "Ficha\tResultado\tDiferencias"
    : "Ficha\tFechaIngreso\tEstado\tTecnico\tCliente\tTelefono\tModelo\tTrabajos");

int okCount = 0;
int differenceCount = 0;
int errorCount = 0;

foreach (var file in files)
{
    try
    {
        FichaFluwork ficha = reader.Read(file.FullName);
        if (compareIndex && indexByFicha != null)
        {
            ComparisonResult comparison = CompareWithIndex(ficha, indexByFicha);
            if (comparison.IsOk)
            {
                okCount++;
            }
            else
            {
                differenceCount++;
            }

            Console.WriteLine(comparison.Line);
            continue;
        }

        okCount++;
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
        errorCount++;
        Console.WriteLine(file.Name + "\tERROR\t" + ex.Message);
    }
}

Console.WriteLine();
Console.WriteLine("Resumen: OK=" + okCount + " Diferencias=" + differenceCount + " Errores=" + errorCount);

return 0;

static ComparisonResult CompareWithIndex(FichaFluwork ficha, IReadOnlyDictionary<string, IndexFluworkRecord> indexByFicha)
{
    if (!indexByFicha.TryGetValue(ficha.Ficha, out IndexFluworkRecord? index))
    {
        return new ComparisonResult(false, ficha.Ficha + "\tFALTA_INDEX\tNo existe en index.dat");
    }

    var differences = new List<string>();

    AddDifference(differences, "cliente", ficha.NombreCompleto, index.NombreCompleto);
    AddDifference(differences, "telefono", ficha.Telefono, index.Telefono);
    AddDifference(differences, "modelo", ficha.Modelo, index.Modelo);
    AddDifference(differences, "fecha", ficha.FechaIngreso, index.Fecha);
    AddDifference(differences, "estado", ficha.Estado, index.Estado);
    AddDifference(differences, "tecnico", ficha.Tecnico, index.Tecnico);
    AddDifference(differences, "confirmacion", ficha.Confirmacion, index.Confirmacion);

    return differences.Count == 0
        ? new ComparisonResult(true, ficha.Ficha + "\tOK\t")
        : new ComparisonResult(false, ficha.Ficha + "\tDIFERENCIAS\t" + string.Join("; ", differences));
}

static void AddDifference(List<string> differences, string field, string fichaValue, string indexValue)
{
    string left = Normalize(fichaValue);
    string right = Normalize(indexValue);

    if (!left.Equals(right, StringComparison.OrdinalIgnoreCase))
    {
        differences.Add(field + " ficha='" + left + "' index='" + right + "'");
    }
}

static string Normalize(string? value)
{
    return (value ?? string.Empty).Trim();
}

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

sealed class ComparisonResult
{
    public ComparisonResult(bool isOk, string line)
    {
        IsOk = isOk;
        Line = line;
    }

    public bool IsOk { get; }
    public string Line { get; }
}
