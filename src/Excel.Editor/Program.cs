using Serilog;
using Serilog.Sinks.SystemConsole.Themes;
using Serilog.Templates;

namespace Excel.Editor
{
    public static class Program
    {
        public static void Main(params string[] files)
        {
            if (files.Any() != true)
            {
                files = new[]
                {
                    "example.md"
                };
            }

            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .WriteTo.Console(outputTemplate: "[{Timestamp:yyy-MM-ddTHH:mm:ss} {Level:w4}] {Message:lj}{NewLine}{Exception}", theme: AnsiConsoleTheme.Code)
                .CreateLogger();

            var editor = new ExcelEditor();

            foreach (var f in files)
            {
                if (!File.Exists(f))
                {
                    Log.Error("Cannot find " + f);
                    continue;
                }

                Log.Information("Processing " + f);
                var template = new Template(f);
                editor.Apply(template);
            }
        }
    }
}