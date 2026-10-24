using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using ClosedXML.Excel;
using Microsoft.Win32;

namespace WpfApp1
{
    public enum ResultKind { Probability, Count, Rate, Time }

    // Строка исходных данных (как ввёл пользователь)
    public record InputRow(string Name, string Symbol, object Value, string Unit);

    // Строка результата. Для Time значение хранится в ЧАСАХ.
    public record ResultRow(string Name, string Symbol, double Value, ResultKind Kind, string Unit = "", bool ShowPercent = false);

    public partial class MainWindow : Window
    {
        private bool _initialized;

        // Снимок последнего расчёта — именно он выгружается в Excel
        private string _lastType = "";
        private List<InputRow>? _lastInputs;
        private List<ResultRow>? _lastResults;
        private double _lastLambda, _lastMu, _lastRho;

        public MainWindow()
        {
            InitializeComponent();
            _initialized = true;
            UpdateVisibility();
        }

        // ---------- Видимость полей ввода ----------
        private void InputModeChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_initialized) return;
            UpdateVisibility();
        }

        private void UpdateVisibility()
        {
            int type = CmbSystemType.SelectedIndex;
            GridN.Visibility = NeedsN(type) ? Visibility.Visible : Visibility.Collapsed;
            GridM.Visibility = NeedsM(type) ? Visibility.Visible : Visibility.Collapsed;

            bool lambdaIsRate = CmbLambdaInputType.SelectedIndex == 0;
            GridLambda.Visibility = lambdaIsRate ? Visibility.Visible : Visibility.Collapsed;
            GridTpr.Visibility = lambdaIsRate ? Visibility.Collapsed : Visibility.Visible;

            bool muIsRate = CmbMuInputType.SelectedIndex == 0;
            GridMu.Visibility = muIsRate ? Visibility.Visible : Visibility.Collapsed;
            GridTobs.Visibility = muIsRate ? Visibility.Collapsed : Visibility.Visible;
        }

        private static bool NeedsN(int type) => type == 1 || type == 3 || type == 5;
        private static bool NeedsM(int type) => type == 4 || type == 5;

        // ---------- Чтение и проверка ввода ----------
        private static double ReadDouble(TextBox box, string what)
        {
            string s = box.Text.Trim().Replace(',', '.');
            if (!double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out double v)
                || v <= 0 || double.IsInfinity(v))
                throw new Exception($"Введите {what} — положительное число.");
            return v;
        }

        private static int ReadInt(TextBox box, string what)
        {
            if (!int.TryParse(box.Text.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int v) || v < 1)
                throw new Exception($"Введите {what} — целое число не меньше 1.");
            return v;
        }

        private static string ItemText(ComboBox cmb) =>
            (cmb.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "";

        // Интенсивность: в час / в мин / в сек  ->  в час
        private static double RateToPerHour(double v, int unit) => unit switch
        {
            0 => v,
            1 => v * 60,
            _ => v * 3600
        };

        // Время: мин / час / сек  ->  часы
        private static double TimeToHours(double v, int unit) => unit switch
        {
            0 => v / 60,
            1 => v,
            _ => v / 3600
        };

        // ---------- Кнопка «Рассчитать» ----------
        private void BtnCalculate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                int type = CmbSystemType.SelectedIndex;
                string typeName = ItemText(CmbSystemType);
                var inputs = new List<InputRow> { new("Тип СМО", "", typeName, "") };

                int n = 1, m = 0;
                if (NeedsN(type))
                {
                    n = ReadInt(TxtN, "число каналов n");
                    inputs.Add(new("Число каналов", "n", n, "шт."));
                }
                if (NeedsM(type))
                {
                    m = ReadInt(TxtM, "максимальную длину очереди m");
                    inputs.Add(new("Максимальная длина очереди", "m", m, "мест"));
                }

                // λ — из интенсивности или из времени между заявками
                double lambda;
                if (CmbLambdaInputType.SelectedIndex == 0)
                {
                    double v = ReadDouble(TxtLambda, "интенсивность потока заявок λ");
                    lambda = RateToPerHour(v, CmbLambdaUnit.SelectedIndex);
                    inputs.Add(new("Интенсивность потока заявок", "λ", v, ItemText(CmbLambdaUnit)));
                }
                else
                {
                    double t = ReadDouble(TxtTpr, "среднее время между заявками t_пр");
                    lambda = 1.0 / TimeToHours(t, CmbTprUnit.SelectedIndex);
                    inputs.Add(new("Среднее время между заявками", "t_пр", t, ItemText(CmbTprUnit)));
                }

                // μ — из интенсивности или из времени обслуживания
                double mu;
                if (CmbMuInputType.SelectedIndex == 0)
                {
                    double v = ReadDouble(TxtMu, "интенсивность обслуживания μ");
                    mu = RateToPerHour(v, CmbMuUnit.SelectedIndex);
                    inputs.Add(new("Интенсивность обслуживания", "μ", v, ItemText(CmbMuUnit)));
                }
                else
                {
                    double t = ReadDouble(TxtTobs, "время обслуживания t_обс");
                    mu = 1.0 / TimeToHours(t, CmbTobsUnit.SelectedIndex);
                    inputs.Add(new("Среднее время обслуживания", "t_обс", t, ItemText(CmbTobsUnit)));
                }

                double rho = lambda / mu;

                List<ResultRow> results = type switch
                {
                    0 => CalculateSingleFail(lambda, mu),
                    1 => CalculateMultiFail(lambda, mu, n),
                    2 => CalculateSingleInfiniteQueue(lambda, rho),
                    3 => CalculateMultiInfiniteQueue(lambda, rho, n),
                    4 => CalculateSingleLimitedQueue(lambda, rho, m),
                    _ => CalculateMultiLimitedQueue(lambda, mu, rho, n, m)
                };

                // Сохраняем снимок для выгрузки
                _lastType = typeName;
                _lastInputs = inputs;
                _lastResults = results;
                _lastLambda = lambda;
                _lastMu = mu;
                _lastRho = rho;

                var sb = new StringBuilder();
                sb.AppendLine($"Тип СМО: {typeName}");
                sb.AppendLine("Приведённые параметры:");
                sb.AppendLine($"λ = {lambda:F4} заявок/час; μ = {mu:F4} обсл./час; ρ = {rho:F4}");
                sb.AppendLine(new string('-', 60));
                foreach (var r in results)
                    sb.AppendLine($"{r.Name} ({r.Symbol}): {FormatValue(r)}");

                TxtResults.Text = sb.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private static string FormatValue(ResultRow r) => r.Kind switch
        {
            ResultKind.Probability => r.ShowPercent ? $"{r.Value:F4} ({r.Value * 100:F2}%)" : $"{r.Value:F4}",
            ResultKind.Rate => $"{r.Value:F2} {r.Unit}",
            ResultKind.Time => FormatTime(r.Value),
            _ => string.IsNullOrEmpty(r.Unit) ? $"{r.Value:F4}" : $"{r.Value:F4} {r.Unit}"
        };

        private static string FormatTime(double hours)
        {
            if (hours < 0.016)
                return $"{hours * 3600:F2} сек.";
            return $"{hours:F4} ч. ({hours * 60:F2} мин.)";
        }

        // ---------- Выгрузка в настоящий .xlsx ----------
        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            if (_lastResults == null || _lastInputs == null)
            {
                MessageBox.Show("Сначала выполните расчёт, чтобы было что выгружать.", "Внимание",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var sfd = new SaveFileDialog
            {
                Filter = "Книга Excel (*.xlsx)|*.xlsx",
                DefaultExt = ".xlsx",
                FileName = "Результаты_СМО.xlsx"
            };
            if (sfd.ShowDialog() != true) return;

            try
            {
                ExportToExcel(sfd.FileName);

                var answer = MessageBox.Show("Файл сохранён.\nОткрыть его в Excel?", "Готово",
                    MessageBoxButton.YesNo, MessageBoxImage.Information);
                if (answer == MessageBoxResult.Yes)
                    Process.Start(new ProcessStartInfo(sfd.FileName) { UseShellExecute = true });
            }
            catch (IOException)
            {
                MessageBox.Show("Не удалось записать файл. Возможно, он уже открыт в Excel — закройте его и повторите.",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при сохранении файла: " + ex.Message, "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportToExcel(string path)
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Результаты СМО");

            ws.Column(1).Width = 42;
            ws.Column(2).Width = 14;
            ws.Column(3).Width = 18;
            ws.Column(4).Width = 16;

            int row = 1;
            ws.Cell(row, 1).Value = "Расчёт характеристик СМО";
            ws.Cell(row, 1).Style.Font.Bold = true;
            ws.Cell(row, 1).Style.Font.FontSize = 14;
            row += 2;

            // 1. Исходные данные — как ввёл пользователь
            WriteHeader(ws, row++, "Исходные данные");
            int start = row - 1;
            foreach (var inp in _lastInputs!)
            {
                ws.Cell(row, 1).Value = inp.Name;
                ws.Cell(row, 2).Value = inp.Symbol;
                SetValue(ws.Cell(row, 3), inp.Value);
                ws.Cell(row, 4).Value = inp.Unit;
                ws.Cell(row, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                row++;
            }
            Border(ws, start, row - 1);
            row++;

            // 2. Приведённые параметры (внутренние единицы расчёта)
            WriteHeader(ws, row++, "Приведённые параметры");
            start = row - 1;
            AddNumber(ws, ref row, "Интенсивность потока заявок", "λ", _lastLambda, "заявок/час", "0.0000");
            AddNumber(ws, ref row, "Интенсивность обслуживания", "μ", _lastMu, "обсл./час", "0.0000");
            AddNumber(ws, ref row, "Приведённая интенсивность (λ/μ)", "ρ", _lastRho, "", "0.0000");
            Border(ws, start, row - 1);
            row++;

            // 3. Результаты
            WriteHeader(ws, row++, "Результаты расчёта");
            start = row - 1;
            foreach (var r in _lastResults!)
            {
                switch (r.Kind)
                {
                    case ResultKind.Probability:
                        AddNumber(ws, ref row, r.Name, r.Symbol, r.Value, "", "0.0000");
                        break;
                    case ResultKind.Rate:
                        AddNumber(ws, ref row, r.Name, r.Symbol, r.Value, r.Unit, "0.0000");
                        break;
                    case ResultKind.Time:
                        AddNumber(ws, ref row, r.Name, r.Symbol, r.Value * 60, "мин", "0.0000");
                        break;
                    default:
                        AddNumber(ws, ref row, r.Name, r.Symbol, r.Value, r.Unit, "0.0000");
                        break;
                }
            }
            Border(ws, start, row - 1);

            wb.SaveAs(path);
        }

        private static void WriteHeader(IXLWorksheet ws, int row, string firstTitle)
        {
            string[] titles = { firstTitle, "Обозначение", "Значение", "Ед. изм." };
            for (int i = 0; i < titles.Length; i++)
            {
                var c = ws.Cell(row, i + 1);
                c.Value = titles[i];
                c.Style.Font.Bold = true;
                c.Style.Fill.BackgroundColor = XLColor.LightGray;
            }
        }

        private static void AddNumber(IXLWorksheet ws, ref int row, string name, string symbol,
                                      double value, string unit, string numberFormat)
        {
            ws.Cell(row, 1).Value = name;
            ws.Cell(row, 2).Value = symbol;
            ws.Cell(row, 3).Value = value;
            ws.Cell(row, 3).Style.NumberFormat.Format = numberFormat;
            ws.Cell(row, 4).Value = unit;
            row++;
        }

        private static void SetValue(IXLCell cell, object value)
        {
            switch (value)
            {
                case double d: cell.Value = d; break;
                case int i: cell.Value = i; break;
                default: cell.Value = value?.ToString() ?? ""; break;
            }
        }

        private static void Border(IXLWorksheet ws, int firstRow, int lastRow)
        {
            var range = ws.Range(firstRow, 1, lastRow, 4);
            range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        }

        // ---------- Математика ----------
        private static double Factorial(int n)
        {
            double res = 1;
            for (int i = 2; i <= n; i++) res *= i;
            return res;
        }

        private static ResultRow P0Row(double p0) =>
            new("Вероятность простоя", "P0", p0, ResultKind.Probability);
        private static ResultRow PotkRow(double potk) =>
            new("Вероятность отказа", "Pотк", potk, ResultKind.Probability, "", true);
        private static ResultRow QRow(double q) =>
            new("Относительная пропускная способность", "Q", q, ResultKind.Probability);
        private static ResultRow ARow(double a) =>
            new("Абсолютная пропускная способность", "A", a, ResultKind.Rate, "заявок/час");
        private static ResultRow KRow(double k) =>
            new("Среднее число занятых каналов", "k", k, ResultKind.Count, "");
        private static ResultRow LochRow(double l) =>
            new("Среднее число заявок в очереди", "Lоч", l, ResultKind.Count, "");
        private static ResultRow TochRow(double t) =>
            new("Среднее время ожидания в очереди", "Tоч", t, ResultKind.Time);
        private static ResultRow LsistRow(double l) =>
            new("Среднее число заявок в системе", "Lсист", l, ResultKind.Count, "");
        private static ResultRow TsistRow(double t) =>
            new("Среднее время в системе", "Tсист", t, ResultKind.Time);

        private static List<ResultRow> CalculateSingleFail(double lambda, double mu)
        {
            double p0 = mu / (lambda + mu);
            double potk = lambda / (lambda + mu);
            double q = 1 - potk;
            double a = lambda * q;

            return new() { P0Row(p0), PotkRow(potk), QRow(q), ARow(a) };
        }

        private static List<ResultRow> CalculateMultiFail(double lambda, double mu, int n)
        {
            double rho = lambda / mu;
            double sum = 0;
            for (int k = 0; k <= n; k++) sum += Math.Pow(rho, k) / Factorial(k);
            double p0 = 1 / sum;
            double potk = (Math.Pow(rho, n) / Factorial(n)) * p0;
            double q = 1 - potk;
            double a = lambda * q;
            double kAvg = a / mu;

            return new() { P0Row(p0), PotkRow(potk), QRow(q), ARow(a), KRow(kAvg) };
        }

        private static List<ResultRow> CalculateSingleInfiniteQueue(double lambda, double rho)
        {
            if (rho >= 1)
                throw new Exception("Условие ρ < 1 не выполняется! Очередь будет расти до бесконечности.");

            double p0 = 1 - rho;
            double loch = (rho * rho) / (1 - rho);
            double toch = loch / lambda;
            double lsist = rho / (1 - rho);
            double tsist = lsist / lambda;

            return new() { P0Row(p0), LochRow(loch), TochRow(toch), LsistRow(lsist), TsistRow(tsist) };
        }

        private static List<ResultRow> CalculateMultiInfiniteQueue(double lambda, double rho, int n)
        {
            if (rho / n >= 1)
                throw new Exception("Условие ρ/n < 1 не выполняется! Очередь будет расти до бесконечности.");

            double sum = 0;
            for (int k = 0; k < n; k++) sum += Math.Pow(rho, k) / Factorial(k);
            sum += (Math.Pow(rho, n) / Factorial(n)) * (n / (n - rho));
            double p0 = 1 / sum;

            double pn = (Math.Pow(rho, n) / Factorial(n)) * (n / (n - rho)) * p0;
            double loch = pn * (rho / n) / (1 - rho / n);
            double toch = loch / lambda;
            double lsist = loch + rho;
            double tsist = lsist / lambda;

            return new() { P0Row(p0), LochRow(loch), TochRow(toch), LsistRow(lsist), TsistRow(tsist) };
        }

        private static List<ResultRow> CalculateSingleLimitedQueue(double lambda, double rho, int m)
        {
            double p0;
            if (Math.Abs(rho - 1.0) < 0.00001)
                p0 = 1.0 / (m + 2);
            else
                p0 = (1 - rho) / (1 - Math.Pow(rho, m + 2));

            double potk = Math.Pow(rho, m + 1) * p0;
            double q = 1 - potk;
            double a = lambda * q;

            double loch = 0;
            for (int i = 1; i <= m; i++)
                loch += i * Math.Pow(rho, i + 1) * p0;

            double toch = loch / lambda;
            double kAvg = 1 - p0;
            double lsist = loch + kAvg;
            double tsist = lsist / lambda;

            return new()
            {
                P0Row(p0), PotkRow(potk), QRow(q), ARow(a),
                LochRow(loch), TochRow(toch), LsistRow(lsist), TsistRow(tsist)
            };
        }

        private static List<ResultRow> CalculateMultiLimitedQueue(double lambda, double mu, double rho, int n, int m)
        {
            double sum = 0;
            for (int k = 0; k <= n; k++) sum += Math.Pow(rho, k) / Factorial(k);
            for (int k = n + 1; k <= n + m; k++) sum += Math.Pow(rho, k) / (Factorial(n) * Math.Pow(n, k - n));
            double p0 = 1 / sum;

            double potk = Math.Pow(rho, n + m) / (Factorial(n) * Math.Pow(n, m)) * p0;
            double q = 1 - potk;
            double a = lambda * q;
            double kAvg = a / mu;

            double loch = 0;
            for (int k = n + 1; k <= n + m; k++)
            {
                double pk = Math.Pow(rho, k) / (Factorial(n) * Math.Pow(n, k - n)) * p0;
                loch += (k - n) * pk;
            }

            double toch = loch / lambda;
            double lsist = loch + kAvg;
            double tsist = lsist / lambda;

            return new()
            {
                P0Row(p0), PotkRow(potk), QRow(q), ARow(a), KRow(kAvg),
                LochRow(loch), TochRow(toch), LsistRow(lsist), TsistRow(tsist)
            };
        }
    }
}
