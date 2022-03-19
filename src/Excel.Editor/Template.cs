using System.Text.RegularExpressions;

namespace Excel.Editor;

public class Template
{
    private static Regex _sheetNameRegex = new Regex("^#+(.+)", RegexOptions.Compiled);

    public Template(string file)
    {
        var lines = File.ReadAllLines(file);
        var currentSheet = "";
        var currentCommands = new List<string>();
        foreach (var line in lines)
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            var l = line.Trim();
            if (l.StartsWith("file:"))
            {
                ExcelFile = l.Split("file:")[1].Trim();
                continue;
            }

            if (l.StartsWith("params:"))
            {
                var l1 = l.ToLowerInvariant();
                var ps = l1.Split("params:")[1];
                if(ps.Contains("use-title"))
                {
                    UseTitle = true;
                }

                continue;
            }

            if (l.StartsWith("output:"))
            {
                OutputFile = l.Split("output:")[1].Trim();
                continue;
            }

            if (l.StartsWith("fill:"))
            {
                var ps = l.Split("fill:")[1];
                var p = ps.Split(",").Select(e => e.Trim());
                BlankColumns = p.ToArray();
                continue;
            }

            if (l.StartsWith("#"))
            {
                if (!string.IsNullOrWhiteSpace(currentSheet))
                {
                    Commands[currentSheet] = currentCommands.ToArray();
                }

                var match = _sheetNameRegex.Match(l);
                currentSheet = match.Groups[1].Value.Trim();
                continue;
            }

            currentCommands.Add(l);
        }

        if (!string.IsNullOrWhiteSpace(currentSheet))
        {
            Commands[currentSheet] = currentCommands.ToArray();
        }
    }

    public string ExcelFile { get; set; }
    public string OutputFile { get; set; }
    public bool UseTitle { get; set; }
    public string[] BlankColumns { get; set; }
    public IDictionary<string, string[]> Commands { get; set; } = new Dictionary<string, string[]>();
}