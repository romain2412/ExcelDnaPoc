using System;
using System.IO;

namespace ExcelDnaPoc;

// Journal fichier minimaliste pour diagnostiquer le chargement de l'add-in
// (utile car les erreurs d'AutoOpen ne sont visibles que dans Excel).
// Fichier : %TEMP%\ExcelDnaPoc.log
public static class Log
{
    public static readonly string FilePath =
        Path.Combine(Path.GetTempPath(), "ExcelDnaPoc.log");

    public static void Info(string message) => Write("INFO ", message);

    public static void Error(string message, Exception? ex = null) =>
        Write("ERROR", ex is null ? message : $"{message}{Environment.NewLine}{ex}");

    private static void Write(string level, string message)
    {
        try
        {
            File.AppendAllText(FilePath,
                $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} [{level}] {message}{Environment.NewLine}");
        }
        catch
        {
            // Le journal ne doit jamais casser l'add-in.
        }
    }
}
