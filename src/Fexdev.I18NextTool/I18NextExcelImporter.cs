using System.Text.Json.Nodes;

namespace Fexdev.I18NextTool
{
    public class I18NextExcelImporter
    {
        private JsonObject Build(IReadOnlyCollection<LanguagePathValue> items)
        {
            var rootObj = new JsonObject();

            foreach (var item in items)
            {
                var parsed = ParsePath(item.Path);
                var obj = GetObject(rootObj, parsed.segments);
                obj[parsed.propertyName] = item.Value;
            }

            return rootObj;
        }

        // TODO: Switch expression
        private (string[] segments, string propertyName) ParsePath(string path)
        {
            var parts = path.Split('.', StringSplitOptions.RemoveEmptyEntries);

            if (parts.Length == 0)
            {
                throw new InvalidOperationException("invalid path " + path);
            }

            if (parts.Length == 1)
            {
                return (Array.Empty<string>(), parts[0]);
            }

            return (parts[..^1], parts[^1]);
        }


        private JsonObject GetObject(JsonObject rootObj, string[] segments)
        {
            var obj = rootObj;
            foreach (var part in segments)
            {
                var pathObj = (JsonObject?)obj[part];

                if (pathObj is null)
                {
                    pathObj = new JsonObject();
                    obj[part] = pathObj;
                }

                obj = pathObj;
            }

            return obj;
        }

    }
}
