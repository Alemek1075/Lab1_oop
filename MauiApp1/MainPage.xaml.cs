using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Serialization;

namespace MauiApp1;

public partial class MainPage : ContentPage
{
    private int _rowCount = 5;
    private int _colCount = 5;
    private Dictionary<string, string> _expressions = new();
    private List<List<Entry>> _cells = new();
    private bool _showValues = false;
    private bool _isDirty = false;

    public MainPage()
    {
        InitializeComponent();
        RefreshGrid();
    }

    private void RefreshGrid()
    {
        MainGrid.BatchBegin();

        try
        {
            MainGrid.Children.Clear();
            MainGrid.RowDefinitions.Clear();
            MainGrid.ColumnDefinitions.Clear();
            _cells.Clear();

            MainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            MainGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = 50 });

            var corner = new Label
            {
                Text = "#",
                HorizontalTextAlignment = TextAlignment.Center,
                VerticalTextAlignment = TextAlignment.Center,
                BackgroundColor = Colors.LightGray
            };
            MainGrid.Add(corner, 0, 0);

            for (int c = 0; c < _colCount; c++)
            {
                MainGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = 150 });
                var label = new Label
                {
                    Text = GetColumnName(c),
                    HorizontalTextAlignment = TextAlignment.Center,
                    VerticalTextAlignment = TextAlignment.Center,
                    FontAttributes = FontAttributes.Bold,
                    BackgroundColor = Colors.LightGray
                };
                MainGrid.Add(label, c + 1, 0);
            }

            for (int r = 0; r < _rowCount; r++)
            {
                MainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                var label = new Label
                {
                    Text = (r + 1).ToString(),
                    HorizontalTextAlignment = TextAlignment.Center,
                    VerticalTextAlignment = TextAlignment.Center,
                    BackgroundColor = Colors.LightGray
                };
                MainGrid.Add(label, 0, r + 1);

                var rowCells = new List<Entry>();
                for (int c = 0; c < _colCount; c++)
                {
                    CreateAndAddCell(r, c, rowCells);
                }
                _cells.Add(rowCells);
            }

            RefreshAllCellsVisuals();
        }
        finally
        {
            MainGrid.BatchCommit();
        }
    }

    private void CreateAndAddCell(int r, int c, List<Entry> rowCells)
    {
        string cellName = GetColumnName(c) + (r + 1);
        var entry = new Entry { BackgroundColor = Colors.White };

        entry.Focused += (s, e) =>
        {
            if (_showValues) return;

            if (_expressions.ContainsKey(cellName))
            {
                entry.Text = _expressions[cellName];
            }
            else
            {
                entry.Text = "";
            }
        };

        entry.Unfocused += (s, e) =>
        {
            string oldValue = _expressions.ContainsKey(cellName) ? _expressions[cellName] : "";

            if (_showValues)
            {
                if (entry.Text != oldValue && !string.IsNullOrEmpty(entry.Text) && IsCalculationResult(cellName, entry.Text))
                {
                    return;
                }
            }

            if (oldValue != entry.Text)
            {
                _expressions[cellName] = entry.Text;
                _isDirty = true;
                UpdateStatus("Є незбережені зміни");
                RefreshAllCellsVisuals();
            }
        };

        MainGrid.Add(entry, c + 1, r + 1);

        if (rowCells != null)
        {
            rowCells.Add(entry);
        }
    }

    private bool IsCalculationResult(string cellName, string text)
    {
        try
        {
            var visited = new List<string>();
            object val = ResolveCell(cellName, visited);

            string calculatedStr;
            if (val is decimal d)
            {
                string strVal = d.ToString(CultureInfo.InvariantCulture);
                calculatedStr = (strVal.Length > 10) ? d.ToString("E4", CultureInfo.InvariantCulture) : strVal;
            }
            else
            {
                calculatedStr = val.ToString();
            }

            return text == calculatedStr;
        }
        catch
        {
            return false;
        }
    }

    private void RefreshAllCellsVisuals()
    {
        for (int r = 0; r < _rowCount; r++)
        {
            for (int c = 0; c < _colCount; c++)
            {
                string cellName = GetColumnName(c) + (r + 1);
                string expr = _expressions.ContainsKey(cellName) ? _expressions[cellName] : "";

                if (_showValues)
                {
                    if (!string.IsNullOrEmpty(expr) && expr.StartsWith("="))
                    {
                        try
                        {
                            var visited = new List<string>();
                            object val = ResolveCell(cellName, visited);

                            if (val is decimal d)
                            {
                                string strVal = d.ToString(CultureInfo.InvariantCulture);
                                if (strVal.Length > 10)
                                {
                                    _cells[r][c].Text = d.ToString("E4", CultureInfo.InvariantCulture);
                                }
                                else
                                {
                                    _cells[r][c].Text = strVal;
                                }
                            }
                            else
                            {
                                _cells[r][c].Text = val.ToString();
                            }

                            _cells[r][c].TextColor = Colors.Black;
                        }
                        catch (InvalidOperationException)
                        {
                            _cells[r][c].Text = "Цикл";
                            _cells[r][c].TextColor = Colors.Red;
                        }
                        catch (ArgumentException)
                        {
                            _cells[r][c].Text = "Помилка";
                            _cells[r][c].TextColor = Colors.Red;
                        }
                        catch (Exception)
                        {
                            _cells[r][c].Text = "Помилка";
                            _cells[r][c].TextColor = Colors.Red;
                        }
                    }
                    else
                    {
                        _cells[r][c].Text = expr;
                        _cells[r][c].TextColor = Colors.Black;
                    }
                }
                else
                {
                    _cells[r][c].TextColor = Colors.Black;
                    _cells[r][c].Text = expr;
                }
            }
        }
    }

    private void ToggleMode_Clicked(object sender, EventArgs e)
    {
        _showValues = !_showValues;
        ModeButton.Text = _showValues ? "Режим: Значення" : "Режим: Вирази";
        RefreshAllCellsVisuals();
    }

    private string GetColumnName(int index)
    {
        const string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        string res = "";
        while (index >= 0)
        {
            res = alphabet[index % 26] + res;
            index = (index / 26) - 1;
        }
        return res;
    }

    private int GetColumnIndex(string columnName)
    {
        int index = 0;
        foreach (char c in columnName)
        {
            index = index * 26 + (c - 'A' + 1);
        }
        return index - 1;
    }

    private void AddRow_Clicked(object sender, EventArgs e)
    {
        MainGrid.BatchBegin();

        _rowCount++;
        _isDirty = true;
        UpdateStatus("Рядок додано (не збережено)");
        int newRowIdx = _rowCount;

        MainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var label = new Label
        {
            Text = _rowCount.ToString(),
            HorizontalTextAlignment = TextAlignment.Center,
            VerticalTextAlignment = TextAlignment.Center,
            BackgroundColor = Colors.LightGray
        };
        MainGrid.Add(label, 0, newRowIdx);

        var newRowCells = new List<Entry>();
        for (int c = 0; c < _colCount; c++)
        {
            CreateAndAddCell(_rowCount - 1, c, newRowCells);
        }
        _cells.Add(newRowCells);

        MainGrid.BatchCommit();
    }

    private void AddCol_Clicked(object sender, EventArgs e)
    {
        MainGrid.BatchBegin();

        _colCount++;
        _isDirty = true;
        UpdateStatus("Стовпчик додано (не збережено)");
        int newColIdx = _colCount;

        MainGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = 150 });

        var label = new Label
        {
            Text = GetColumnName(_colCount - 1),
            HorizontalTextAlignment = TextAlignment.Center,
            VerticalTextAlignment = TextAlignment.Center,
            FontAttributes = FontAttributes.Bold,
            BackgroundColor = Colors.LightGray
        };
        MainGrid.Add(label, newColIdx, 0);

        for (int r = 0; r < _rowCount; r++)
        {
            CreateAndAddCell(r, _colCount - 1, null);
            var entry = (Entry)MainGrid.Children.Last();
            _cells[r].Add(entry);
        }

        MainGrid.BatchCommit();
    }

    private void DelRow_Clicked(object sender, EventArgs e)
    {
        if (_rowCount > 0)
        {
            _isDirty = true;
            int rowToDelete = _rowCount;
            for (int c = 0; c < _colCount; c++)
            {
                string colName = GetColumnName(c);
                string cellKey = colName + rowToDelete;
                if (_expressions.ContainsKey(cellKey))
                {
                    _expressions.Remove(cellKey);
                }
            }
            _rowCount--;
            RefreshGrid();
            UpdateStatus("Рядок видалено (не збережено)");
        }
    }

    private void DelCol_Clicked(object sender, EventArgs e)
    {
        if (_colCount > 0)
        {
            _isDirty = true;
            string colToDelete = GetColumnName(_colCount - 1);
            for (int r = 0; r < _rowCount; r++)
            {
                string cellKey = colToDelete + (r + 1);
                if (_expressions.ContainsKey(cellKey))
                {
                    _expressions.Remove(cellKey);
                }
            }
            _colCount--;
            RefreshGrid();
            UpdateStatus("Стовпчик видалено (не збережено)");
        }
    }

    private async void Help_Clicked(object sender, EventArgs e)
    {
#pragma warning disable CS0618
        await DisplayAlert("Довідка",
            "Лабораторна робота №1. Варіант 24.\n\n" +
            "Підтримувані операції:\n" +
            "+, -, *, / : Арифметичні операції\n" +
            "mod, div : Остача та цілочисельне ділення\n" +
            "inc, dec : Інкремент (+1) та декремент (-1)\n" +
            "=, <, >, <=, >=, <> : Операції порівняння (1 - істина, 0 - хибність)\n\n" +
            "Приклад: inc(A1) + 10 * (B2 mod 3)\n" +
            "Десятковий роздільник - ТІЛЬКИ крапка (.)\n" +
            "Кома (,) та текст у формулах - помилка.",
            "OK");
#pragma warning restore CS0618
    }

    private async void About_Clicked(object sender, EventArgs e)
    {
#pragma warning disable CS0618
        await DisplayAlert("Про програму",
            "Лабораторна робота №1\n" +
            "Студента: Базилевича Олексія\n" +
            "Варіант: 24\n" +
            "Пункти (1, 2, 5, 8, 9;)\n\n" +
            "Програма призначена для обробки електронних таблиць з підтримкою математичних виразів.",
            "OK");
#pragma warning restore CS0618
    }

    private async void Save_Clicked(object sender, EventArgs e)
    {
        await SaveDataAsync();
    }

    private async Task<bool> SaveDataAsync()
    {
#if WINDOWS
        var savePicker = new Windows.Storage.Pickers.FileSavePicker();
        savePicker.SuggestedStartLocation = Windows.Storage.Pickers.PickerLocationId.DocumentsLibrary;
        savePicker.FileTypeChoices.Add("Excel-like XML", new List<string>() { ".xlcx" });
        savePicker.FileTypeChoices.Add("CSV", new List<string>() { ".csv" });
        savePicker.SuggestedFileName = "MyTable";

        var window = App.Current.Windows.FirstOrDefault();
        if (window != null)
        {
            var hwnd = ((MauiWinUIWindow)window.Handler.PlatformView).WindowHandle;
            WinRT.Interop.InitializeWithWindow.Initialize(savePicker, hwnd);
        }

        var file = await savePicker.PickSaveFileAsync();
        if (file != null)
        {
            string content;
            if (file.FileType == ".csv")
            {
                content = GetCsvContent();
            }
            else
            {
                content = GetXmlContent();
            }

            await Windows.Storage.FileIO.WriteTextAsync(file, content);
            _isDirty = false;
            UpdateStatus($"Збережено: {file.Name}");
            return true;
        }
        return false;
#else
        string content = GetXmlContent();
        string path = Path.Combine(FileSystem.AppDataDirectory, "table.xlcx");
        await File.WriteAllTextAsync(path, content);
        _isDirty = false;
        UpdateStatus($"Збережено (локально): {path}");
        return true;
#endif
    }

    private void Exit_Clicked(object sender, EventArgs e)
    {
        if (_isDirty)
        {
            ExitDialog.IsVisible = true;
        }
        else
        {
            Application.Current.Quit();
        }
    }

    private async void DialogSaveExit_Clicked(object sender, EventArgs e)
    {
        if (await SaveDataAsync())
        {
            Application.Current.Quit();
        }
        ExitDialog.IsVisible = false;
    }

    private void DialogNoSave_Clicked(object sender, EventArgs e)
    {
        Application.Current.Quit();
    }

    private void DialogCancel_Clicked(object sender, EventArgs e)
    {
        ExitDialog.IsVisible = false;
    }

    private string GetXmlContent()
    {
        var serializableData = new SavedData
        {
            Rows = _rowCount,
            Cols = _colCount,
            Cells = _expressions.Select(x => new CellRecord { Name = x.Key, Value = x.Value }).ToList()
        };

        var serializer = new XmlSerializer(typeof(SavedData));
        using (var writer = new StringWriter())
        {
            serializer.Serialize(writer, serializableData);
            return writer.ToString();
        }
    }

    private string GetCsvContent()
    {
        var sb = new StringBuilder();
        for (int r = 0; r < _rowCount; r++)
        {
            var rowValues = new List<string>();
            for (int c = 0; c < _colCount; c++)
            {
                string cellKey = GetColumnName(c) + (r + 1);
                string value = _expressions.ContainsKey(cellKey) ? _expressions[cellKey] : "";

                if (value.Contains(",") || value.Contains("\"") || value.Contains("\n"))
                {
                    value = $"\"{value.Replace("\"", "\"\"")}\"";
                }
                rowValues.Add(value);
            }
            sb.AppendLine(string.Join(",", rowValues));
        }
        return sb.ToString();
    }

    private async void Load_Clicked(object sender, EventArgs e)
    {
        if (_isDirty)
        {
            bool answer = await DisplayAlert("Увага", "Є незбережені зміни. Продовжити завантаження?", "Так", "Ні");
            if (!answer) return;
        }

        try
        {
            var result = await FilePicker.Default.PickAsync(new PickOptions
            {
                PickerTitle = "Оберіть файл",
                FileTypes = new FilePickerFileType(new Dictionary<DevicePlatform, IEnumerable<string>>
                {
                    { DevicePlatform.WinUI, new[] { ".xlcx", ".csv" } },
                    { DevicePlatform.Android, new[] { "application/xml", "text/csv" } },
                    { DevicePlatform.MacCatalyst, new[] { "public.xml", "public.comma-separated-values-text" } },
                })
            });

            if (result != null)
            {
                using var stream = await result.OpenReadAsync();
                using var reader = new StreamReader(stream);
                string content = await reader.ReadToEndAsync();

                if (result.FileName.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                {
                    ParseCsvContent(content);
                }
                else
                {
                    ParseXmlContent(content);
                }

                RefreshGrid();
                _isDirty = false;
                UpdateStatus("Завантажено успішно");
            }
        }
        catch (Exception ex)
        {
            UpdateStatus("Помилка завантаження");
        }
    }

    private void ParseXmlContent(string xmlContent)
    {
        var serializer = new XmlSerializer(typeof(SavedData));
        using (var reader = new StringReader(xmlContent))
        {
            var data = (SavedData)serializer.Deserialize(reader);
            if (data != null)
            {
                _rowCount = data.Rows;
                _colCount = data.Cols;
                _expressions = data.Cells.ToDictionary(k => k.Name, v => v.Value);
            }
        }
    }

    private void ParseCsvContent(string csvContent)
    {
        var lines = csvContent.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
        var data = new Dictionary<string, string>();
        int r = 0;
        int maxC = 0;

        foreach (var line in lines)
        {
            if (string.IsNullOrEmpty(line)) continue;

            var values = SplitCsvLine(line);
            if (values.Count > maxC) maxC = values.Count;

            for (int c = 0; c < values.Count; c++)
            {
                if (!string.IsNullOrEmpty(values[c]))
                {
                    string key = GetColumnName(c) + (r + 1);
                    data[key] = values[c];
                }
            }
            r++;
        }

        _rowCount = r;
        _colCount = maxC;
        _expressions = data;
    }

    private List<string> SplitCsvLine(string line)
    {
        var result = new List<string>();
        bool inQuotes = false;
        var current = new StringBuilder();

        for (int i = 0; i < line.Length; i++)
        {
            char c = line[i];

            if (c == '\"')
            {
                if (inQuotes && i + 1 < line.Length && line[i + 1] == '\"')
                {
                    current.Append('\"');
                    i++;
                }
                else
                {
                    inQuotes = !inQuotes;
                }
            }
            else if (c == ',' && !inQuotes)
            {
                result.Add(current.ToString());
                current.Clear();
            }
            else
            {
                current.Append(c);
            }
        }
        result.Add(current.ToString());
        return result;
    }

    private void Calculate_Clicked(object sender, EventArgs e)
    {
        RefreshAllCellsVisuals();
        UpdateStatus("Оновлено");
    }

    private void UpdateStatus(string message)
    {
        string dirtyMark = _isDirty ? "*" : "";
        StatusLabel.Text = $"{message} {dirtyMark}";
    }

    private object ResolveCell(string cellName, List<string> visited)
    {
        cellName = cellName.ToUpper();

        if (visited.Contains(cellName))
        {
            throw new InvalidOperationException("Cycle detected");
        }
        visited.Add(cellName);

        var match = Regex.Match(cellName, @"^([A-Z]+)(\d+)$");
        if (!match.Success) throw new ArgumentException("Invalid cell reference");

        string colPart = match.Groups[1].Value;
        int rowPart = int.Parse(match.Groups[2].Value);

        int colIndex = GetColumnIndex(colPart);
        int rowIndex = rowPart - 1;

        if (rowIndex < 0 || rowIndex >= _rowCount || colIndex < 0 || colIndex >= _colCount)
        {
            throw new ArgumentException("Cell out of bounds");
        }

        if (!_expressions.ContainsKey(cellName)) return 0m;

        string expr = _expressions[cellName];

        if (!expr.StartsWith("="))
        {
            if (decimal.TryParse(expr, NumberStyles.Float, CultureInfo.InvariantCulture, out decimal val)) return val;
            return expr;
        }

        if (Regex.IsMatch(expr, @"^=[A-Za-z]+\d+$"))
        {
            string target = expr.Substring(1).ToUpper();
            return ResolveCell(target, new List<string>(visited));
        }

        return Calculator.Evaluate(expr, (name) =>
        {
            var val = ResolveCell(name, new List<string>(visited));
            return val;
        });
    }

    public class SavedData
    {
        public int Rows { get; set; }
        public int Cols { get; set; }
        public List<CellRecord> Cells { get; set; } = new();
    }

    public class CellRecord
    {
        public string Name { get; set; }
        public string Value { get; set; }
    }
}