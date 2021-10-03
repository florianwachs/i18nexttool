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

    private IWorkbook ExportToExcel(List<LanguagePathValue> values)
    {
        var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Translations");
        var headerRow = sheet.CreateRow(0);
        headerRow.CreateCell(0).SetCellValue("Key");
        headerRow.CreateCell(1).SetCellValue("en_EN");

        int rowIndex = 1;
        int columnIndex = 0;
        foreach (var item in values)
        {
            columnIndex = 0;
            var row = sheet.CreateRow(rowIndex++);
            row.CreateCell(columnIndex++).SetCellValue(item.Path);
            row.CreateCell(columnIndex++).SetCellValue(item.Value);
        }

        sheet.AutoSizeColumn(0);
        sheet.AutoSizeColumn(1);

        return workbook;
    }
}
