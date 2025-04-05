using AngleSharp;
using AngleSharp.Dom;
using ClosedXML.Excel;
using System.Diagnostics;

namespace VintedCompanion;

public partial class Form1 : Form
{
    // Lista transakcji wyciągniętych z HTMLa
    private readonly List<Transaction> transactions = [];

    public Form1()
    {
        InitializeComponent();
    }

    // Jeden przycisk wykonuje najpierw parsowanie HTML, a potem eksport do Excela
    private async void btnProcess_Click(object sender, EventArgs e)
    {
        await ProcessHtml();

        if (transactions.Count == 0)
        {
            MessageBox.Show("Nie znaleziono operacji do eksportu.\nWybierz poprawny plik *.html", "Błąd",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        ExportExcel();
    }

    private async Task ProcessHtml()
    {
        OpenFileDialog ofd = new()
        {
            Filter = "Pliki HTML (*.html)|*.html"
        };

        if (ofd.ShowDialog() != DialogResult.OK)
            return;

        string htmlContent = File.ReadAllText(ofd.FileName);

        var config = Configuration.Default;
        var context = BrowsingContext.New(config);
        var document = await context.OpenAsync(req => req.Content(htmlContent));

        transactions.Clear();

        // Szukamy elementów li z klasą "pile__element"
        var items = document.QuerySelectorAll("li.pile__element");

        foreach (var li in items)
        {
            var titleElement = li.QuerySelector(".web_ui__Cell__title");
            if (titleElement == null)
                continue;
            string operationTypeText = titleElement.TextContent.Trim();

            // Interesują nas tylko operacje: Sprzedane, Zakup oraz Zwrot kosztów
            if (operationTypeText != "Sprzedane" && operationTypeText != "Zakup" && operationTypeText != "Zwrot kosztów")
                continue;

            var descElement = li.QuerySelector(".web_ui__Cell__body");
            string description = descElement != null ? descElement.TextContent.Trim() : "Brak opisu";

            var amountElement = li.QuerySelector("h2.web_ui__Text__text.web_ui__Text__title.web_ui__Text__right.web_ui__Text__parent");
            if (amountElement == null)
                continue;
            string amountText = amountElement.TextContent.Trim();
            // Usuwamy zbędne znaki – pozostawiamy cyfry oraz przecinek/kropkę
            string cleanAmount = new([.. amountText.Where(c => char.IsDigit(c) || c == ',' || c == '.')]);
            cleanAmount = cleanAmount.Replace(',', '.');

            if (!decimal.TryParse(cleanAmount, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal amount))
            {
                amount = 0;
            }

            // Wyciągamy datę – zakładamy, że data znajduje się w elemencie suffix jako ostatni węzeł tekstowy
            string dateText = "";
            var suffixElement = li.QuerySelector(".web_ui__Cell__suffix");
            if (suffixElement != null && suffixElement.LastChild != null && suffixElement.LastChild.NodeType == NodeType.Text)
            {
                dateText = suffixElement.LastChild.TextContent.Trim();
            }

            TransactionType type;
            if (operationTypeText == "Sprzedane")
            {
                type = TransactionType.Sold;
                amount = Math.Abs(amount); // sprzedaż – dodatnia
            }
            else if (operationTypeText == "Zwrot kosztów")
            {
                type = TransactionType.Refund;
                amount = Math.Abs(amount); // zwroty – dodatnie
            }
            else if (operationTypeText == "Zakup")
            {
                type = TransactionType.Purchase;
                if (amount > 0)
                    amount = -amount; // zakupy – ujemne
            }
            else
            {
                continue;
            }

            transactions.Add(new Transaction
            {
                Description = description,
                Amount = amount,
                Type = type,
                Date = dateText
            });
        }

        statusLabel.Text = $"Status: Znaleziono {transactions.Count} operacji w pliku {ofd.FileName}";
    }

    private void ExportExcel()
    {
        statusLabel.Text = "Status: Eksport danych do pliku Excel...";

        // Rozdzielamy operacje na trzy grupy
        var purchaseTransactions = transactions.Where(t => t.Type == TransactionType.Purchase).ToList();
        var soldTransactions = transactions.Where(t => t.Type == TransactionType.Sold).ToList();
        var refundTransactions = transactions.Where(t => t.Type == TransactionType.Refund).ToList();

        // Obliczenia sum (dla zakupów wartość bezwzględna, bo zapisane jako ujemne)
        decimal purchaseSum = Math.Abs(purchaseTransactions.Sum(t => t.Amount));
        decimal soldSum = soldTransactions.Sum(t => t.Amount);
        decimal refundSum = refundTransactions.Sum(t => t.Amount);
        // Bilans = (sprzedane + zwroty) - zakupione
        decimal balance = soldSum + refundSum - purchaseSum;

        var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Raport transakcji");

        // Główny nagłówek – wiersz 1
        ws.Cell("A1").Value = "Raport operacji Vinted";
        ws.Range("A1:K1").Merge();
        ws.Cell("A1").Style.Font.FontSize = 26;
        ws.Cell("A1").Style.Font.Bold = true;
        ws.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ws.Row(1).Height = 30;

        // Pozostawiamy pusty wiersz 2

        // Sekcja zakupione (lewa tabela: kolumny A-C) – wiersz 3 i 4
        ws.Cell("A3").Value = "Zakupione";
        ws.Cell("A4").Value = "Opis";
        ws.Cell("B4").Value = "Kwota (zł)";
        ws.Cell("C4").Value = "Data";
        ws.Range("A3:C4").Style.Fill.BackgroundColor = XLColor.LightBlue;
        ws.Range("A3:C4").Style.Font.Bold = true;

        // Sekcja sprzedane (centralna tabela: kolumny E-G) – wiersz 3 i 4
        ws.Cell("E3").Value = "Sprzedane";
        ws.Cell("E4").Value = "Opis";
        ws.Cell("F4").Value = "Kwota (zł)";
        ws.Cell("G4").Value = "Data";
        ws.Range("E3:G4").Style.Fill.BackgroundColor = XLColor.LightGreen;
        ws.Range("E3:G4").Style.Font.Bold = true;

        // Sekcja zwrot kosztów (prawa tabela: kolumny I-K) – wiersz 3 i 4
        ws.Cell("I3").Value = "Zwrot kosztów";
        ws.Cell("I4").Value = "Opis";
        ws.Cell("J4").Value = "Kwota (zł)";
        ws.Cell("K4").Value = "Data";
        ws.Range("I3:K4").Style.Fill.BackgroundColor = XLColor.LightSalmon;
        ws.Range("I3:K4").Style.Font.Bold = true;

        int startDataRow = 5;
        int maxPurchase = purchaseTransactions.Count;
        int maxSold = soldTransactions.Count;
        int maxRefund = refundTransactions.Count;

        // Wypełnienie danych sekcji zakupionych
        for (int i = 0; i < maxPurchase; i++)
        {
            int currentRow = startDataRow + i;
            ws.Cell(currentRow, 1).Value = purchaseTransactions[i].Description;
            ws.Cell(currentRow, 2).Value = purchaseTransactions[i].Amount;
            ws.Cell(currentRow, 3).Value = purchaseTransactions[i].Date;
        }

        // Wypełnienie danych sekcji sprzedanych
        for (int i = 0; i < maxSold; i++)
        {
            int currentRow = startDataRow + i;
            ws.Cell(currentRow, 5).Value = soldTransactions[i].Description;
            ws.Cell(currentRow, 6).Value = soldTransactions[i].Amount;
            ws.Cell(currentRow, 7).Value = soldTransactions[i].Date;
        }

        // Wypełnienie danych sekcji zwrotów
        for (int i = 0; i < maxRefund; i++)
        {
            int currentRow = startDataRow + i;
            ws.Cell(currentRow, 9).Value = refundTransactions[i].Description;
            ws.Cell(currentRow, 10).Value = refundTransactions[i].Amount;
            ws.Cell(currentRow, 11).Value = refundTransactions[i].Date;
        }

        // Wiersze podsumowujące – w każdej sekcji wstawiamy pusty wiersz między danymi a podsumowaniem
        int sumPurchaseRow = startDataRow + maxPurchase + 1;
        int sumSoldRow = startDataRow + maxSold + 1;
        int sumRefundRow = startDataRow + maxRefund + 1;

        // Wyznaczamy ostatni wiersz podsumowania jako maksimum ze wszystkich sekcji
        int lastSummaryRow = Math.Max(Math.Max(sumPurchaseRow, sumSoldRow), sumRefundRow);

        // Wiersz z bilansem – umieszczony poniżej ostatniej sekcji (dodajemy np. 2 wiersze przerwy)
        int overallBalanceRow = lastSummaryRow + 2;

        // Podsumowania dla każdej sekcji pozostają bez zmian
        ws.Cell(sumPurchaseRow, 1).Value = $"Suma zakupionych ({purchaseTransactions.Count} operacji):";
        ws.Cell(sumPurchaseRow, 2).Value = purchaseSum;
        ws.Range($"A{sumPurchaseRow}:C{sumPurchaseRow}").Style.Fill.BackgroundColor = XLColor.LightGray;

        ws.Cell(sumSoldRow, 5).Value = $"Suma sprzedanych ({soldTransactions.Count} operacji):";
        ws.Cell(sumSoldRow, 6).Value = soldSum;
        ws.Range($"E{sumSoldRow}:G{sumSoldRow}").Style.Fill.BackgroundColor = XLColor.LightGray;

        ws.Cell(sumRefundRow, 9).Value = $"Suma zwrotów ({refundTransactions.Count} operacji):";
        ws.Cell(sumRefundRow, 10).Value = refundSum;
        ws.Range($"I{sumRefundRow}:K{sumRefundRow}").Style.Fill.BackgroundColor = XLColor.LightGray;

        // Wiersz z bilansem – umieszczony po prawej stronie, poniżej sekcji zwrotów
        ws.Cell(overallBalanceRow, 9).Value = "Bilans (Sprzedane + Zwroty - Zakupione):";
        ws.Cell(overallBalanceRow, 10).Value = balance;
        ws.Range($"I{overallBalanceRow}:K{overallBalanceRow}").Style.Fill.BackgroundColor = XLColor.Yellow;
        ws.Range($"I{overallBalanceRow}:K{overallBalanceRow}").Style.Font.Bold = true;
        ws.Range($"I{overallBalanceRow}:K{overallBalanceRow}").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        // Dodanie obramowania do zakresu danych – teraz obejmuje od wiersza 3 do overallBalanceRow
        //var dataRange = ws.Range($"A3:K{overallBalanceRow}");
        //dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
        //dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

        ws.Columns().AdjustToContents();

        SaveFileDialog sfd = new()
        {
            Filter = "Plik Excel (*.xlsx)|*.xlsx"
        };

        if (sfd.ShowDialog() == DialogResult.OK)
        {
            try
            {
                workbook.SaveAs(sfd.FileName);

                statusLabel.Text = $"Status: Zapisano plik Excel w {sfd.FileName}";

                // Wyświetlamy komunikat z przyciskiem "Tak" umożliwiającym otwarcie pliku
                DialogResult result = MessageBox.Show("Plik Excel został zapisany.\nCzy chcesz go otworzyć?", "Sukces", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (result == DialogResult.Yes)
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = sfd.FileName,
                        UseShellExecute = true
                    });
                }
            }
            catch (Exception)
            {
                statusLabel.Text = "Status: Ostatni eksport danych do pliku Excel zakończony niepowodzeniem.";

                MessageBox.Show("Nie udało się otworzyć/zapisać pliku.\nZadzwoń do Kacpra, i wyślij mu tekst z okienka, które pojawi się po tym jak wciśniesz OK.", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }
    }
}

public enum TransactionType
{
    Sold,      // operacje sprzedaży (bez zwrotów)
    Refund,    // operacje zwrotu kosztów
    Purchase   // operacje kupna
}

public class Transaction
{
    public string? Description { get; set; }
    // Kwota: przy sprzedaży i zwrotach – wartość dodatnia, przy zakupie – ujemna
    public decimal Amount { get; set; }
    public TransactionType Type { get; set; }
    public string? Date { get; set; }
}