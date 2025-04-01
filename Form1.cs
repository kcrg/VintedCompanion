using AngleSharp;
using ClosedXML.Excel;

namespace VintedCompanion;

public partial class Form1 : Form
{
    // Lista transakcji wyciągniętych z HTMLa
    private readonly List<Transaction> transactions = [];

    public Form1()
    {
        InitializeComponent();
    }

    private async void btnProcess_Click(object sender, EventArgs e)
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

        // Czyścimy listę przed rozpoczęciem parsowania
        transactions.Clear();

        // Szukamy wszystkich elementów li o klasie "pile__element"
        var items = document.QuerySelectorAll("li.pile__element");

        foreach (var li in items)
        {
            // Pobieramy tytuł operacji, np. "Sprzedane", "Zakup" lub "Zwrot kosztów"
            var titleElement = li.QuerySelector(".web_ui__Cell__title");
            if (titleElement == null)
                continue;
            string operationTypeText = titleElement.TextContent.Trim();

            // Interesują nas tylko operacje: Sprzedane, Zakup i Zwrot kosztów
            if (operationTypeText != "Sprzedane" && operationTypeText != "Zakup" && operationTypeText != "Zwrot kosztów")
                continue;

            // Pobieramy opis operacji
            var descElement = li.QuerySelector(".web_ui__Cell__body");
            string description = descElement != null ? descElement.TextContent.Trim() : "Brak opisu";

            // Pobieramy kwotę – zakładamy, że znajduje się w elemencie h2 z odpowiednimi klasami
            var amountElement = li.QuerySelector("h2.web_ui__Text__text.web_ui__Text__title.web_ui__Text__right.web_ui__Text__parent");
            if (amountElement == null)
                continue;
            string amountText = amountElement.TextContent.Trim();

            // Usuwamy zbędne znaki – pozostawiamy cyfry oraz przecinek lub kropkę
            string cleanAmount = new([.. amountText.Where(c => char.IsDigit(c) || c == ',' || c == '.')]);
            // Zamieniamy przecinek na kropkę dla poprawnego parsowania
            cleanAmount = cleanAmount.Replace(',', '.');

            if (!decimal.TryParse(cleanAmount, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal amount))
            {
                amount = 0;
            }

            TransactionType type;
            // Jeśli operacja to "Sprzedane" lub "Zwrot kosztów", traktujemy ją jako sprzedaż (przychód)
            if (operationTypeText == "Sprzedane" || operationTypeText == "Zwrot kosztów")
            {
                type = TransactionType.Sold;
                amount = Math.Abs(amount); // upewniamy się, że kwota jest dodatnia
            }
            else if (operationTypeText == "Zakup")
            {
                type = TransactionType.Purchase;
                // Jeśli kwota jest dodatnia, zmieniamy znak na ujemny
                if (amount > 0)
                    amount = -amount;
            }
            else
            {
                continue;
            }

            transactions.Add(new Transaction
            {
                Description = description,
                Amount = amount,
                Type = type
            });
        }

        statusLabel.Text = $"Znaleziono {transactions.Count} operacji w pliku {ofd.FileName}.";

        if (transactions.Count == 0)
        {
            MessageBox.Show($"Nie znaleziono żadnych operacji. Wybierz poprawny plik.", "Parsowanie zakończone", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        else
        {
            MessageBox.Show($"Znaleziono {transactions.Count} operacji.", "Parsowanie zakończone", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private void btnExportExcel_Click(object sender, EventArgs e)
    {
        if (transactions.Count == 0)
        {
            MessageBox.Show("Brak operacji do eksportu. Najpierw załaduj plik HTML.", "Błąd",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // Rozdzielamy operacje na zakupione i sprzedane
        var purchaseTransactions = transactions.Where(t => t.Type == TransactionType.Purchase).ToList();
        var soldTransactions = transactions.Where(t => t.Type == TransactionType.Sold).ToList();

        // Sumy operacji – przy zakupach kwoty są ujemne, dlatego używamy wartości bezwzględnej
        decimal purchaseSum = Math.Abs(purchaseTransactions.Sum(t => t.Amount));
        decimal soldSum = soldTransactions.Sum(t => t.Amount);
        decimal balance = soldSum - purchaseSum;

        // Utworzenie nowego workbooka i arkusza
        var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Raport transakcji");

        // Ustawiamy nagłówki dla sekcji zakupionych (lewa strona)
        ws.Cell("A1").Value = "Zakupione";
        ws.Cell("A2").Value = "Opis";
        ws.Cell("B2").Value = "Kwota (zł)";

        // Ustawiamy nagłówki dla sekcji sprzedanych (prawa strona)
        ws.Cell("D1").Value = "Sprzedane";
        ws.Cell("D2").Value = "Opis";
        ws.Cell("E2").Value = "Kwota (zł)";

        // Ustalamy maksymalną liczbę wierszy dla obu tabel
        int maxRows = Math.Max(purchaseTransactions.Count, soldTransactions.Count);
        int startDataRow = 3; // dane zaczynamy od wiersza 3

        // Wypełniamy obie sekcje transakcjami
        for (int i = 0; i < maxRows; i++)
        {
            int currentRow = startDataRow + i;
            if (i < purchaseTransactions.Count)
            {
                ws.Cell(currentRow, 1).Value = purchaseTransactions[i].Description;
                ws.Cell(currentRow, 2).Value = purchaseTransactions[i].Amount;
            }
            if (i < soldTransactions.Count)
            {
                ws.Cell(currentRow, 4).Value = soldTransactions[i].Description;
                ws.Cell(currentRow, 5).Value = soldTransactions[i].Amount;
            }
        }

        // Wiersz z sumami – umieszczamy je pod danymi
        int totalRow = startDataRow + maxRows;
        ws.Cell(totalRow, 1).Value = "Suma zakupionych:";
        ws.Cell(totalRow, 2).Value = purchaseSum;
        ws.Cell(totalRow, 4).Value = "Suma sprzedanych:";
        ws.Cell(totalRow, 5).Value = soldSum;

        // Wiersz z bilansem
        int balanceRow = totalRow + 2;
        ws.Cell(balanceRow, 1).Value = "Bilans (Sprzedane - Zakupione):";
        ws.Cell(balanceRow, 2).Value = balance;

        // Ustawienie formatu walutowego dla kolumn z kwotami
        ws.Column(2).Style.NumberFormat.Format = "#,##0.00 zł";
        ws.Column(5).Style.NumberFormat.Format = "#,##0.00 zł";

        // Dopasowanie szerokości kolumn
        ws.Columns().AdjustToContents();

        // Zapis do pliku Excel przy użyciu SaveFileDialog
        SaveFileDialog sfd = new()
        {
            Filter = "Plik Excel (*.xlsx)|*.xlsx"
        };

        if (sfd.ShowDialog() == DialogResult.OK)
        {
            workbook.SaveAs(sfd.FileName);
            MessageBox.Show("Plik Excel został zapisany.", "Sukces", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}

public enum TransactionType
{
    Sold,      // operacje sprzedaży (także „Zwrot kosztów” traktujemy jako przychód)
    Purchase   // operacje kupna
}

public class Transaction
{
    public string? Description { get; set; }
    // Kwota: sprzedaż – wartość dodatnia, kupno – wartość ujemna.
    public decimal Amount { get; set; }
    public TransactionType Type { get; set; }
}