using System.Text.RegularExpressions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Serilog;

namespace Excel.Editor;

public class ExcelEditor
{
    private static Regex cmdRegex = new Regex(@"(\w+-\w+):(.+)", RegexOptions.Compiled);
    private static Regex titleCmdRegex = new Regex(@"([^:]+):(.*)", RegexOptions.Compiled);
    private static Regex commentRegex = new Regex(@"^//(.*)", RegexOptions.Compiled);

    public void Apply(Template template)
    {
        using (var stream = new FileStream(template.ExcelFile, FileMode.Open))
        {
            stream.Position = 0;
            var xssWorkbook = new XSSFWorkbook(stream);
            var useTitle = template.UseTitle;
            foreach (var c in template.Commands)
            {
                var sheet = xssWorkbook.GetSheet(c.Key);
                Log.Information($"Opened sheet {c.Key}");
                var headerRow = sheet.GetRow(0);
                var cells = headerRow.Cells;
                var title = cells.Select(e => e.ToString()).ToList();

                var lastRow = sheet.LastRowNum;
                var currentRow = -1;
                var currentCol = -1;
                var currentCmd = "";
                foreach (var cmd in c.Value)
                {
                    if (commentRegex.IsMatch(cmd))
                    {
                        continue;
                    }
                    if (useTitle)
                    {
                        var match = titleCmdRegex.Match(cmd);
                        if (match.Success)
                        {
                            var cmd1 = match.Groups[1].Value.Trim();
                            var text = match.Groups[2].Value.Trim();
                            var col = title.IndexOf(cmd1);
                            if (template.BlankColumns.Contains(cmd1))
                            {
                                // fill
                                currentCmd = cmd1;
                                WriteToCell(sheet, currentRow, col, text);
                                continue;
                            }
                            else
                            {
                                // find
                                (currentRow, currentCol) = GetCellEquals(sheet, col, lastRow, text);
                                continue;
                            }
                        }
                        else if (!string.IsNullOrWhiteSpace(currentCmd))
                        {
                            var text = cmd.Trim();
                            var col = title.IndexOf(currentCmd);
                            AppendToCell(sheet, currentRow, col, text);
                            continue;
                        }
                    }
                    else
                    {
                        if (cmd.StartsWith("cell"))
                        {
                            if (cmd.StartsWith("cell equals"))
                            {
                                var cmd1 = cmd.Replace("cell equals:", "").Trim();
                                (currentRow, currentCol) = GetCellEquals(lastRow, sheet, cmd1);
                            }
                            else if (cmd.StartsWith("cell starts"))
                            {
                                var cmd1 = cmd.Replace("cell starts:", "").Trim();
                                (currentRow, currentCol) = GetCellStartsWith(lastRow, sheet, cmd1);
                            }

                            continue;
                        }
                    }

                    if (cmd.StartsWith("add-"))
                    {
                        var match = cmdRegex.Match(cmd);
                        var cmd1 = match.Groups[1].Value.Trim();
                        var text = match.Groups[2].Value.Trim();
                        currentCmd = cmd1;
                        switch (cmd1)
                        {
                            case "add-r":
                                {
                                    currentCol++;
                                    WriteToCell(sheet, currentRow, currentCol, text);
                                    break;
                                }
                            case "add-l":
                                {
                                    currentCol--;
                                    WriteToCell(sheet, currentRow, currentCol, text);
                                    break;
                                }
                            case "add-b":
                                {
                                    currentRow++;
                                    WriteToCell(sheet, currentRow, currentCol, text);
                                    break;
                                }
                            case "add-t":
                                {
                                    currentRow--;
                                    WriteToCell(sheet, currentRow, currentCol, text);
                                    break;
                                }
                        }

                        continue;
                    }

                    var cmd3 = cmd.Trim();
                    switch (currentCmd)
                    {
                        case "add-r":
                        case "add-l":
                        case "add-b":
                        case "add-t":
                            AppendToCell(sheet, currentRow, currentCol, cmd3);
                            break;
                    }
                }
            }

            try
            {
                Save(template, xssWorkbook);
            }
            catch (Exception e)
            {
                Log.Error(e.Message);
            }
        }
    }

    private static void Save(Template template, XSSFWorkbook xssWorkbook)
    {
        using (var memoryStream = new MemoryStream()) //creating memoryStream
        {
            xssWorkbook.Write(memoryStream);
            using (var file = new FileStream(template.OutputFile, FileMode.Create, FileAccess.Write))
            {
                memoryStream.WriteTo(file);
                memoryStream.Close();
            }
        }
    }

    private static void WriteToCell(ISheet sheet, int currentRow, int currentCol, string text)
    {
        var row = sheet.GetRow(currentRow) ?? sheet.CreateRow(currentRow);
        var cell = row.GetCell(currentCol) ?? row.CreateCell(currentCol);
        if (!string.IsNullOrWhiteSpace(text))
        {
            if (text.StartsWith("-"))
            {
                text = "'" + text;
            }
            cell.SetCellValue(text);
            Log.Information("added " + TrimText(text));
        }

        cell.CellStyle.VerticalAlignment = VerticalAlignment.Top;
    }

    private static void AppendToCell(ISheet sheet, int currentRow, int currentCol, string cmd3)
    {
        var row = sheet.GetRow(currentRow);
        var cell = row.GetCell(currentCol);
        var t = cell.ToString();
        if (!string.IsNullOrWhiteSpace(t))
        {
            t = t + Environment.NewLine + cmd3;
        }
        else
        {
            t = cmd3;
        }
        cell.SetCellValue(t);

        cell.CellStyle.WrapText = true;
        Log.Information("appended " + TrimText(cmd3) + $" at ({cell.Address})");
    }

    private static (int, int) GetMatchedCell(int lastRow, ISheet sheet, string cmd1, Func<string, bool> func)
    {
        for (var r = 0; r <= lastRow; r++)
        {
            var row = sheet.GetRow(r);
            var cellCount = row.LastCellNum;
            for (var col = 0; col < cellCount; col++)
            {
                var cell = row.GetCell(col);
                if (cell == null)
                {
                    continue;
                }

                if (func(cell.ToString()))
                {
                    Log.Information("found \"" + TrimText(cmd1) + "\"" + $" at ({cell.Address})");
                    return (r, col);
                }
            }
        }

        Log.Error("failed to find \"" + TrimText(cmd1) + "\"");
        return (-1, -1);
    }

    private static string TrimText(string t)
    {
        const int len = 50;
        if (t.Length < len)
        {
            return t;
        }

        return t.Substring(0, len - 3) + "...";
    }

    private static (int, int) GetCellEquals(ISheet sheet, int column, int lastRow, string cmd1)
    {
        for (var r = 0; r <= lastRow; r++)
        {
            var row = sheet.GetRow(r);

            var cell = row.GetCell(column);
            if (cell == null)
            {
                continue;
            }

            if (cell.ToString() == cmd1)
            {
                Log.Information("found \"" + TrimText(cmd1) + "\"" + $" at ({cell.Address})");
                return (r, column);
            }
        }

        Log.Error("failed to find \"" + TrimText(cmd1) + "\"");
        return (-1, -1);
    }

    private static (int, int) GetCellEquals(int lastRow, ISheet sheet, string cmd1)
    {
        return GetMatchedCell(lastRow, sheet, cmd1, (s) => s == cmd1);
    }

    private static (int, int) GetCellStartsWith(int lastRow, ISheet sheet, string cmd1)
    {
        return GetMatchedCell(lastRow, sheet, cmd1, (s) => s.StartsWith(cmd1));
    }
}