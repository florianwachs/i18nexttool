using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Text.Json.Nodes;

namespace Fexdev.I18NextTool;

public class I18NextExcelExporter
{
    public List<LanguagePathValue> FlattenJson(string jsonRaw, string language)
    {
        var rootNode = JsonNode.Parse(jsonRaw);
        var rootObj = rootNode.AsObject();

        List<LanguagePathValue> values = new();
        foreach (var x in rootObj)
        {
            WalkObjects(values, x.Value, x.Key, language);
        }

        return values.OrderBy(k => k.Path).ToList();
    }

    private void WalkObjects(List<LanguagePathValue> values, JsonNode? node, string path, string language)
    {
        if (node is JsonArray)
        {
            throw new InvalidOperationException("arrays ar unsupported for i18next");
        }
        else if (node is JsonValue value)
        {
            values.Add(new(language, path, value.ToString()));
        }
        else if (node is JsonObject obj)
        {
            foreach (var item in obj)
            {
                WalkObjects(values, item.Value, path + "." + item.Key, language);
            }
        }
    }

    public IWorkbook ExportToExcel(List<LanguagePathValue> values)
    {
        //         Path                lang   value
        Dictionary<string, Dictionary<string, string>> map = new();
        HashSet<string> languages = new();

        foreach (var item in values)
        {
            if (!map.TryGetValue(item.Path, out var languageValues))
            {
                languageValues = new();
                map[item.Path] = languageValues;
            }

            languageValues.Add(item.Language, item.Value);
            languages.Add(item.Language);
        }

        var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Translations");

        int rowIndex = 0;
        int columnIndex = 0;
        var headerRow = sheet.CreateRow(rowIndex);
        headerRow.CreateCell(columnIndex++).SetCellValue("Key");
        foreach (var lang in languages)
        {
            headerRow.CreateCell(columnIndex++).SetCellValue(lang);
        }

        rowIndex = 1;
        columnIndex = 0;

        var orderedPaths = map.Keys.OrderBy(k => k);
        foreach (var path in orderedPaths)
        {
            columnIndex = 0;
            var row = sheet.CreateRow(rowIndex++);
            var languageValues = map[path];

            row.CreateCell(columnIndex++).SetCellValue(path);
            foreach (var lang in languages)
            {
                if (languageValues.TryGetValue(lang, out var value))
                {
                    row.CreateCell(columnIndex).SetCellValue(lang);
                }

                columnIndex++;
            }
        }

        for (int i = 0; i < languages.Count; i++)
        {
            sheet.AutoSizeColumn(i);
        }

        return workbook;
    }
}
