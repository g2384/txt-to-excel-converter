using System.Text.RegularExpressions;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using Serilog;

namespace Excel.Editor;

public class ExcelEditor
{
    private static readonly Regex cmdRegex = new Regex(@"(\w+-\w+):(.+)", RegexOptions.Compiled);
    private static readonly Regex titleCmdRegex = new Regex(@"([^:]+):(.*)", RegexOptions.Compiled);
    private static readonly Regex commentRegex = new Regex(@"^//(.*)", RegexOptions.Compiled);

    public void Apply(Template template)
    {
        var stream = File.Open(template.ExcelFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

        Log.Information($"Opening Excel file: {template.ExcelFile}");
        stream.Position = 0;
        var workbook = new XLWorkbook(stream);
        var useTitle = template.UseTitle;
        foreach (var c in template.Commands)
        {
            var sheet = workbook.Worksheet(c.Key);
            Log.Information($"Opened sheet: {c.Key}");
            var title = GetHeaders(sheet);

            var lastRow = sheet.LastRowUsed().LastCellUsed().Address.RowNumber;
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
                        int col = GetColumnIndex(title, cmd1);
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
                        var col = GetColumnIndex(title, currentCmd);
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
            Save(template, workbook);
        }
        catch (Exception e)
        {
            Log.Error(e.Message);
        }
    }

    private static int GetColumnIndex(List<string> title, string cmd1)
    {
        return title.IndexOf(cmd1) + 1;
    }

    private static List<string> GetHeaders(IXLWorksheet sheet)
    {
        var headerRow = sheet.Row(1);
        var cells = headerRow.Cells();
        var title = cells.Select(GetCellValue).ToList();
        return title;
    }

    private static void Save(Template template, XLWorkbook xssWorkbook)
    {
        Log.Information($"Saving to {template.OutputFile}");
        using (var memoryStream = new MemoryStream()) //creating memoryStream
        {
            xssWorkbook.SaveAs(template.OutputFile);
        }
    }

    private static void WriteToCell(IXLWorksheet sheet, int currentRow, int currentCol, string text)
    {
        var row = sheet.Row(currentRow);
        var cell = row.Cell(currentCol);
        if (!string.IsNullOrWhiteSpace(text))
        {
            if (text.StartsWith("-"))
            {
                text = "'" + text;
            }
            cell.SetValue(text);
            Log.Information("added " + TrimText(text));
        }

        cell.Style.Alignment.WrapText = true;
        cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
    }

    private static void AppendToCell(IXLWorksheet sheet, int currentRow, int currentCol, string cmd3)
    {
        var row = sheet.Row(currentRow);
        var cell = row.Cell(currentCol);
        var t = cell.GetValue<string>();
        var action = "added";
        if (!string.IsNullOrWhiteSpace(t))
        {
            t = t + Environment.NewLine + cmd3;
            action = "appended";
        }
        else
        {
            t = "'" + cmd3;
        }
        cell.SetValue(t);

        cell.Style.Alignment.WrapText = true;
        cell.WorksheetRow().AdjustToContents(cell.Address.ColumnNumber);
        Log.Information($"{action} {TrimText(cmd3)} at ({cell.Address})");
    }

    private static (int, int) GetMatchedCell(int lastRow, IXLWorksheet sheet, string cmd1, Func<string, bool> func)
    {
        for (var r = 0; r <= lastRow; r++)
        {
            var row = sheet.Row(r);
            var cellCount = row.LastCellUsed().Address.ColumnNumber;
            for (var col = 0; col < cellCount; col++)
            {
                var cell = row.Cell(col);
                if (cell == null)
                {
                    continue;
                }

                if (func(cell.GetValue<string>()!))
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

    private static (int, int) GetCellEquals(IXLWorksheet sheet, int column, int lastRow, string cmd1)
    {
        var xlColumn = sheet.Column(column);
        for (var r = 1; r <= lastRow; r++)
        {
            var cell = xlColumn.Cell(r);
            if (cell == null)
            {
                continue;
            }

            var v = GetCellValue(cell);
            if (StringEquals(v, cmd1))
            {
                Log.Information("found \"" + TrimText(cmd1) + "\"" + $" at ({cell.Address})");
                return (r, column);
            }
        }

        Log.Error("failed to find \"" + TrimText(cmd1) + "\"");
        return (-1, -1);
    }

    private static bool StringEquals(string v1, string v2)
    {
        if (ReferenceEquals(v1, v2))
        {
            return true;
        }

        if (v1 == null && v2 == null)
        {
            return true;
        }

        if (v1 == null || v2 == null)
        {
            return false;
        }

        v1 = v1.Trim();
        v2 = v2.Trim();
        return v1.Equals(v2, StringComparison.OrdinalIgnoreCase);
    }

    private static string GetCellValue(IXLCell cell)
    {
        return cell.GetValue<string>();
    }

    private static (int, int) GetCellEquals(int lastRow, IXLWorksheet sheet, string cmd1)
    {
        return GetMatchedCell(lastRow, sheet, cmd1, (s) => s == cmd1);
    }

    private static (int, int) GetCellStartsWith(int lastRow, IXLWorksheet sheet, string cmd1)
    {
        return GetMatchedCell(lastRow, sheet, cmd1, (s) => s.StartsWith(cmd1));
    }
}