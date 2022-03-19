using Serilog;
using Serilog.Templates;

namespace Excel.Editor
{
    public static class Program
    {
        public static void Main(params string[] files)
        {
            if (files?.Any() != true)
            {
                files = new[]
                {
                    "example.md"
                };
            }

            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .WriteTo.Console(new ExpressionTemplate(
                    "[{@t:yyy-MM-ddTHH:mm:ss} {@l:w4}] {@m}\n{@x}"))
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