using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Text.Json;

using Microsoft.AspNetCore.Authorization;

namespace SurveyDashboard.Pages;

[Authorize]
public class IndexModel : PageModel
{
    public record SeriesData(string Name, IReadOnlyList<int> Data, string Color);
    public record QuestionChartData(
        string Question,
        IReadOnlyList<string> Labels,
        IReadOnlyList<int> Values,
        IReadOnlyList<string> Colors,
        string ChartType = "pie",
        IReadOnlyList<SeriesData>? Series = null,
        IReadOnlyList<string>? OtherResponses = null);

    private static readonly string[] ColorPalette = new[]
    {
        "#4e79a7", "#f28e2b", "#e15759", "#76b7b2", "#59a14f",
        "#edc949", "#af7aa1", "#ff9da7", "#9c755f", "#bab0ab",
        "#3b4b8c", "#c66a32", "#b03a48", "#4a9c9b", "#3f7f44"
    };

    public List<QuestionChartData> ChartQuestions { get; private set; } = new();
    public List<string> Segments { get; private set; } = new();
    public string? SelectedSegment { get; private set; }
    public string? ExcelPath { get; private set; }
    public string CurrentView { get; private set; } = "Chatbots"; // Default to Chatbots

    public void OnGet(string? segment, string? view)
    {
        var basePath = Directory.GetCurrentDirectory();
        var path = Path.GetFullPath(Path.Combine(basePath, "Data", "Enquête_Totaal_ChatGPT.xlsx"));

        ExcelPath = path;
        SelectedSegment = string.IsNullOrWhiteSpace(segment) ? null : segment.Trim();
        CurrentView = string.IsNullOrWhiteSpace(view) ? "Chatbots" : view.Trim();

        Segments = LoadSegments(path);
        ChartQuestions = LoadQuestions(path, SelectedSegment, CurrentView);
    }

    private List<QuestionChartData> LoadQuestions(string path, string? selectedSegment, string currentView)
    {
        var result = new List<QuestionChartData>();

        if (!System.IO.File.Exists(path))
        {
            return result;
        }

        using var workbook = new XLWorkbook(path);
        var worksheet = workbook.Worksheets.FirstOrDefault();
        if (worksheet is null)
        {
            return result;
        }

        var headerRow = worksheet.FirstRowUsed();
        if (headerRow is null)
        {
            return result;
        }

        // Read all columns starting from index 3 (Column C)
        var questionHeaders = headerRow.Cells().Skip(2).ToList(); 
        var optionsByColumn = LoadOptionsByColumn(workbook, headerRow);

        // Define specific ranges for Bar charts (Grids)
        // Bar 1: G, H, I (indices 7, 8, 9) -> Relative in list: 4, 5, 6
        // Bar 2: J, K, L, M (indices 10, 11, 12, 13) -> Relative in list: 7, 8, 9, 10
        var bar1Indices = new[] { 7, 8, 9 };
        var bar2Indices = new[] { 10, 11, 12, 13 };

        // Process columns in order
        var processedColumns = new HashSet<int>();

        foreach (var header in questionHeaders)
        {
            int colNum = header.Address.ColumnNumber;
            
            if (processedColumns.Contains(colNum))
            {
                continue;
            }

            // FILTER LOGIC:
            // "Github Copilot" view -> Columns AB (28) and further
            // "General" view -> Columns before AB (< 28)
            string headerText = header.GetString().Trim();
            
            bool isCopilotQuestion = colNum >= 28; // AB = 28
            bool isCopilotView = string.Equals(currentView, "Copilot", StringComparison.OrdinalIgnoreCase);

            if (isCopilotView && !isCopilotQuestion)
            {
                continue; // Skip non-copilot questions in Copilot view
            }
            if (!isCopilotView && isCopilotQuestion)
            {
                continue; // Skip copilot questions in General view
            }

            // Check if this column is part of Bar 1
            if (bar1Indices.Contains(colNum))
            {
                var barHeaders = questionHeaders.Where(h => bar1Indices.Contains(h.Address.ColumnNumber)).ToList();
                result.Add(BuildStackedBarQuestion(worksheet, barHeaders, selectedSegment, optionsByColumn));
                foreach (var h in barHeaders) processedColumns.Add(h.Address.ColumnNumber);
                continue;
            }

            // Check if this column is part of Bar 2
            if (bar2Indices.Contains(colNum))
            {
                var barHeaders = questionHeaders.Where(h => bar2Indices.Contains(h.Address.ColumnNumber)).ToList();
                result.Add(BuildStackedBarQuestion(worksheet, barHeaders, selectedSegment, optionsByColumn));
                foreach (var h in barHeaders) processedColumns.Add(h.Address.ColumnNumber);
                continue;
            }

            // Check for "Chatbots" questions (New requirement)
            // Header starts with: "Gebruik je wel eens verschillende chatbots?"
            if (headerText.StartsWith("Gebruik je wel eens verschillende chatbots?", StringComparison.OrdinalIgnoreCase))
            {
                // We assume these are contiguous or we just find all of them.
                // Let's find all headers matching this prefix to treat them as one group.
                // To avoid re-processing, we need a flag or check.
                // We can check if we already processed this group? 
                // Since we iterate, if we find one, we can find all related ones immediately.
                
                // But wait, the loop iterates ALL headers. We must mark them as processed or skip.
                // Let's rely on a list of processed columns?
                // Or just: if this is the FIRST one we encounter, process the whole group.
                
                var chatbotHeaders = questionHeaders
                    .Where(h => h.GetString().Trim().StartsWith("Gebruik je wel eens verschillende chatbots?", StringComparison.OrdinalIgnoreCase))
                    .ToList();

                // If the current header is the first in that list, process. Else continue.
                // If the current header is the first in that list, process. Else continue.
                if (chatbotHeaders.Count > 0 && !processedColumns.Contains(colNum))
                {
                     result.Add(BuildStackedBarQuestion(worksheet, chatbotHeaders, selectedSegment, optionsByColumn, supportMultiValue: true));
                     foreach (var h in chatbotHeaders) processedColumns.Add(h.Address.ColumnNumber);
                }
                // Always continue if it matches, to skip default Single handling
                continue;
            }
            
            // Check for bracketed Copilot questions (Col >= 28)
            // If header contains [ and ], we treat it as a grid question group.
            if (colNum >= 28 && headerText.Contains('[') && headerText.Contains(']'))
            {
                var stem = ExtractQuestionStem(headerText);
                if (!string.IsNullOrWhiteSpace(stem))
                {
                    // Find all columns with the same stem
                    var groupHeaders = questionHeaders
                        .Where(h => 
                        {
                            var s = ExtractQuestionStem(h.GetString().Trim());
                            return string.Equals(s, stem, StringComparison.OrdinalIgnoreCase);
                        })
                        .ToList();
                        
                    if (groupHeaders.Count > 0)
                    {
                        result.Add(BuildStackedBarQuestion(worksheet, groupHeaders, selectedSegment, optionsByColumn));
                        foreach (var h in groupHeaders) processedColumns.Add(h.Address.ColumnNumber);
                    }
                    continue;
                }
            }

            // Process as Single Column (Pie OR Text)
            var questionText = headerText; // Use existing variable
            if (string.IsNullOrWhiteSpace(questionText))
            {
                questionText = $"Vraag {colNum - 1}";
            }

            var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var otherResponses = new HashSet<string>(StringComparer.OrdinalIgnoreCase); // For text answers

            bool isMultiValue = colNum == 14 || headerText.Equals("Wanneer kies je voor gebruik van Github Copilot?", StringComparison.OrdinalIgnoreCase);
            bool hasOptions = optionsByColumn.ContainsKey(colNum);
            
            // If no options, we treat it as Open Question (Text only), UNLESS we want to support auto-discovery of pie slices.
            // Requirement: "Als er geen opties zijn dan is het een open vraag en moeten alle responses in een tekstvak afgebeeld worden."
            // So: No options -> Text Chart.

            foreach (var row in worksheet.RowsUsed().Skip(1))
            {
                var raw = row.Cell(colNum).GetValue<string>();
                if (string.IsNullOrWhiteSpace(raw)) continue;

                var segmentValue = row.Cell(2).GetValue<string>().Trim().ToUpperInvariant();
                if (!string.IsNullOrWhiteSpace(selectedSegment) &&
                    !segmentValue.Equals(selectedSegment, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var parts = isMultiValue
                   ? raw.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                   : new[] { raw.Trim() };

                foreach (var val in parts)
                {
                    var value = val.Trim();
                    if (string.IsNullOrWhiteSpace(value)) continue;

                    if (!hasOptions)
                    {
                        // Open question: just collect the text
                        otherResponses.Add(value);
                    }
                    else
                    {
                        // Pie chart logic with Options Validation
                        string key = value;
                        bool isOther = false;
                        string? otherText = null;

                        var validOptions = optionsByColumn[colNum];
                        var match = validOptions.FirstOrDefault(o => string.Equals(o, value, StringComparison.OrdinalIgnoreCase));
                        
                        if (match is not null)
                        {
                            key = match;
                        }
                        else
                        {
                            isOther = true;
                            otherText = value;
                            key = "Others";
                        }

                        if (isOther && !string.IsNullOrWhiteSpace(otherText))
                        {
                            otherResponses.Add(otherText);
                        }

                        if (!counts.TryAdd(key, 1))
                        {
                            counts[key] += 1;
                        }
                    }
                }
            }
            
            // Ensure all valid predefined options are present in the counts, even if 0
            if (hasOptions)
            {
                var validOptions = optionsByColumn[colNum];
                foreach (var opt in validOptions)
                {
                    if (!counts.ContainsKey(opt))
                    {
                        counts[opt] = 0;
                    }
                }
            }

            if (!hasOptions)
            {
                // Open Text Question
                result.Add(new QuestionChartData(
                    questionText,
                    Array.Empty<string>(), 
                    Array.Empty<int>(), 
                    Array.Empty<string>(), 
                    "text", 
                    null, 
                    otherResponses.ToList()));
            }
            else
            {
                // Pie Chart
                var allLabels = counts.Keys.ToHashSet(StringComparer.OrdinalIgnoreCase);
                if (allLabels.Any())
                {
                    var orderedLabels = GetOrderedLabels(colNum, allLabels, optionsByColumn);
                    var colorMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    for (int i = 0; i < orderedLabels.Count; i++)
                    {
                         colorMap[orderedLabels[i]] = ColorPalette[i % ColorPalette.Length];
                    }

                    var labels = orderedLabels; // Use ordered list directly
                    // Re-sort if needed or just use ordered from options
                    // Let's stick to frequency sort for non-ordered items, but options order is preferred.
                    
                    // Actually, strict option order is better if options exist.
                    // But we might have "Others" which is not in options list.
                    // Let's keep logic simple: 
                    // 1. Options list order
                    // 2. Counts descending for rest (like Others)
                    
                    var values = labels.Select(l => counts.GetValueOrDefault(l, 0)).ToList();
                    var colors = labels.Select(l => colorMap[l]).ToList();

                    result.Add(new QuestionChartData(questionText, labels, values, colors, "pie", null, otherResponses.ToList()));
                }
            }
        }

        return result;
    }

    private List<string> LoadSegments(string path)
    {
        var result = new List<string>();

        if (!System.IO.File.Exists(path))
        {
            return result;
        }

        using var workbook = new XLWorkbook(path);
        var worksheet = workbook.Worksheets.FirstOrDefault();
        if (worksheet is null)
        {
            return result;
        }

        foreach (var row in worksheet.RowsUsed().Skip(1))
        {
            var segment = row.Cell(2).GetValue<string>().Trim().ToUpperInvariant();
            if (string.IsNullOrWhiteSpace(segment))
            {
                continue;
            }

            if (!result.Contains(segment, StringComparer.OrdinalIgnoreCase))
            {
                result.Add(segment);
            }
        }

        result.Sort(StringComparer.OrdinalIgnoreCase);

        var priority = new[] { "5HSD1", "5HSD2", "5HSD3", "5HSD4" };
        var comparer = StringComparer.OrdinalIgnoreCase;
        var ordered = priority.Where(p => result.Contains(p, comparer)).ToList();
        ordered.AddRange(result.Where(r => !priority.Contains(r, comparer)).OrderBy(r => r, comparer));

        return ordered;
    }

    public string SerializeChartData()
    {
        var options = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            WriteIndented = false
        };

        return JsonSerializer.Serialize(ChartQuestions, options);
    }

    public bool IsSegmentSelected(string? segment)
    {
        if (string.IsNullOrWhiteSpace(segment) && string.IsNullOrWhiteSpace(SelectedSegment))
        {
            return true;
        }

        return string.Equals(SelectedSegment, segment, StringComparison.OrdinalIgnoreCase);
    }

    private static string? ExtractBracketLabel(string? header)
    {
        if (string.IsNullOrWhiteSpace(header))
        {
            return null;
        }

        var start = header.IndexOf('[');
        var end = header.IndexOf(']');
        if (start >= 0 && end > start)
        {
            var inner = header.Substring(start + 1, end - start - 1).Trim();
            return string.IsNullOrWhiteSpace(inner) ? null : inner;
        }

        return null;
    }

    private static string? ExtractQuestionStem(string? header)
    {
        if (string.IsNullOrWhiteSpace(header))
        {
            return null;
        }

        var stem = header;
        var bracketIndex = stem.IndexOf('[');
        if (bracketIndex >= 0)
        {
            stem = stem.Substring(0, bracketIndex);
        }

        stem = stem.Trim();
        return string.IsNullOrWhiteSpace(stem) ? null : stem;
    }

    private QuestionChartData BuildStackedBarQuestion(
        IXLWorksheet worksheet, 
        List<IXLCell> barHeaders, 
        string? selectedSegment, 
        Dictionary<int, List<string>> optionsByColumn,
        bool supportMultiValue = false)
    {
        // Y-as labels uit de headers: tekst tussen brackets [].
        var itemLabels = barHeaders
            .Select(h => ExtractBracketLabel(h.GetString()) ?? (string.IsNullOrWhiteSpace(h.GetString()) ? $"Kolom {h.Address.ColumnNumber}" : h.GetString().Trim()))
            .ToList();

        // Keuze-labels (stack keys) uit de waarden.
        var choiceLabels = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var otherResponses = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var header in barHeaders)
        {
            var col = header.Address.ColumnNumber;
            foreach (var row in worksheet.RowsUsed().Skip(1))
            {
                var segmentValue = row.Cell(2).GetValue<string>().Trim();
                if (!string.IsNullOrWhiteSpace(selectedSegment) &&
                    !segmentValue.Equals(selectedSegment, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var raw = row.Cell(col).GetValue<string>();
                if (string.IsNullOrWhiteSpace(raw))
                {
                    continue;
                }

                var rawValues = supportMultiValue 
                    ? raw.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                    : new[] { raw.Trim() };

                foreach (var val in rawValues)
                {
                    var value = val; // Trim already handled above if split, or direct trim.

                    // Determine if other using options or legacy
                    string key = value;
                    bool isOther = false;
                    string? otherText = null;

                    if (optionsByColumn.TryGetValue(col, out var validOptions) && validOptions.Count > 0)
                    {
                        var match = validOptions.FirstOrDefault(o => string.Equals(o, value, StringComparison.OrdinalIgnoreCase));
                        if (match is not null)
                        {
                            key = match;
                        }
                        else
                        {
                            isOther = true;
                            otherText = value;
                            key = "Others";
                        }
                    }
                    else
                    {
                        if (value.StartsWith("other", StringComparison.OrdinalIgnoreCase))
                        {
                            isOther = true;
                            var idx = value.IndexOf(':');
                            if (idx >= 0 && idx < value.Length - 1)
                            {
                                otherText = value[(idx + 1)..].Trim();
                            }
                            key = "Others";
                        }
                    }

                    if (isOther && !string.IsNullOrWhiteSpace(otherText))
                    {
                        otherResponses.Add(otherText);
                    }

                    choiceLabels.Add(key);
                }
            }
        }

        var orderedChoices = GetOrderedLabels(barHeaders.First().Address.ColumnNumber, choiceLabels, optionsByColumn);

        var finalSeries = new List<SeriesData>();
        for (int c = 0; c < orderedChoices.Count; c++)
        {
            var choice = orderedChoices[c];
            var dataPerItem = new List<int>();

            for (int i = 0; i < barHeaders.Count; i++)
            {
                var header = barHeaders[i];
                var col = header.Address.ColumnNumber;
                var count = worksheet.RowsUsed().Skip(1).Count(row =>
                {
                    var segmentValue = row.Cell(2).GetValue<string>().Trim();
                    if (!string.IsNullOrWhiteSpace(selectedSegment) &&
                        !segmentValue.Equals(selectedSegment, StringComparison.OrdinalIgnoreCase))
                    {
                        return false;
                    }

                    var raw = row.Cell(col).GetValue<string>();
                    if (string.IsNullOrWhiteSpace(raw))
                    {
                        return false;
                    }
                    
                    var rawValues = supportMultiValue 
                        ? raw.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                        : new[] { raw.Trim() };
                    
                    // Check if ANY of the values matches the current choice
                    bool matchFound = false;
                    foreach (var val in rawValues)
                    {
                        string key = val;
                        if (optionsByColumn.TryGetValue(col, out var validOptions) && validOptions.Count > 0)
                        {
                            var m = validOptions.FirstOrDefault(o => string.Equals(o, val, StringComparison.OrdinalIgnoreCase));
                            if (m is not null) key = m;
                            else key = "Others";
                        }
                        else
                        {
                            if (val.StartsWith("other", StringComparison.OrdinalIgnoreCase)) key = "Others";
                        }
                        
                        if (string.Equals(key, choice, StringComparison.OrdinalIgnoreCase))
                        {
                            matchFound = true;
                            break; 
                        }
                    }
                    return matchFound;
                });

                dataPerItem.Add(count);
            }

            var color = ColorPalette[c % ColorPalette.Length];
            finalSeries.Add(new SeriesData(choice, dataPerItem, color));
        }

        var barQuestionTitle = ExtractQuestionStem(barHeaders.First().GetString()) ?? "Gridvraag";

        return new QuestionChartData(
            barQuestionTitle,
            itemLabels,
            Array.Empty<int>(),
            Array.Empty<string>(),
            "bar-stacked",
            finalSeries,
            otherResponses.ToList());
    }

    private Dictionary<int, List<string>> LoadOptionsByColumn(IXLWorkbook workbook, IXLRow mainHeaderRow)
    {
        var result = new Dictionary<int, List<string>>();
        var sheet = workbook.Worksheets.FirstOrDefault(w => string.Equals(w.Name, "opties", StringComparison.OrdinalIgnoreCase));
        if (sheet is null)
        {
            return result;
        }

        var headerRow = sheet.FirstRowUsed();
        if (headerRow is null)
        {
            return result;
        }

        var mainHeaders = mainHeaderRow.CellsUsed()
            .ToDictionary(c => c.GetString().Trim(), c => c.Address.ColumnNumber, StringComparer.OrdinalIgnoreCase);

        foreach (var cell in headerRow.CellsUsed())
        {
            var header = cell.GetString().Trim();
            if (string.IsNullOrWhiteSpace(header))
            {
                continue;
            }

            int colNumber = -1;
            if (mainHeaders.TryGetValue(header, out var foundCol))
            {
                colNumber = foundCol;
            }
            else
            {
                colNumber = ColumnLetterToNumber(header);
            }

            if (colNumber <= 0)
            {
                continue;
            }

            var options = new List<string>();
            foreach (var optCell in sheet.Column(cell.Address.ColumnNumber).CellsUsed().Skip(1))
            {
                var val = optCell.GetString().Trim();
                if (string.IsNullOrWhiteSpace(val))
                {
                    continue;
                }
                options.Add(val);
            }

            if (options.Count > 0)
            {
                result[colNumber] = options;
            }
        }

        return result;
    }

    private List<string> GetOrderedLabels(int columnNumber, IEnumerable<string> labels, Dictionary<int, List<string>> optionsByColumn)
    {
        var labelSet = labels.ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (!optionsByColumn.TryGetValue(columnNumber, out var opts) || opts.Count == 0)
        {
            return labelSet.OrderBy(l => l, StringComparer.OrdinalIgnoreCase).ToList();
        }

        var ordered = new List<string>();
        foreach (var opt in opts)
        {
            if (labelSet.Remove(opt))
            {
                ordered.Add(opt);
            }
        }

        ordered.AddRange(labelSet.OrderBy(l => l, StringComparer.OrdinalIgnoreCase));
        return ordered;
    }

    private int ColumnLetterToNumber(string col)
    {
        if (string.IsNullOrWhiteSpace(col))
        {
            return -1;
        }

        col = col.Trim();
        int result = 0;
        foreach (var ch in col.ToUpperInvariant())
        {
            if (ch < 'A' || ch > 'Z')
            {
                return -1;
            }
            result = result * 26 + (ch - 'A' + 1);
        }
        return result;
    }
}