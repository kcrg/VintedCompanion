using AngleSharp;
using AngleSharp.Dom;
using ClosedXML.Excel;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System.Diagnostics;
using System.Globalization;
using Color = QuestPDF.Infrastructure.Color;
using Colors = QuestPDF.Helpers.Colors;
using Document = QuestPDF.Fluent.Document;

namespace VintedCompanion;

public partial class Form1 : Form
{
    // Lista transakcji wyciągniętych z HTMLa
    private readonly List<Transaction> transactions = [];

    public Form1()
    {
        InitializeComponent();

        QuestPDF.Settings.License = LicenseType.Community;

        statusLabel.Text = "Status: Gotowy. Najpierw wybierz plik z Vinted, a potem kliknij przycisk.";
    }

    /// <summary>
    /// Proste, duże okienko informacyjne z czytelnym komunikatem.
    /// </summary>
    private void ShowInfo(string message, string title = "Informacja")
    {
        _ = MessageBox.Show(this, message, title, MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    // Jeden przycisk wykonuje najpierw parsowanie HTML, a potem eksport do Excela lub PDF
    private async void btnProcess_Click(object sender, EventArgs e)
    {
        // KROK 1 – wyjaśnienie, co się zaraz wydarzy
        ShowInfo(
            "KROK 1 z 2\n\n" +
            "Najpierw wybierz plik z historią operacji z Vinted.\n\n" +
            "Po kliknięciu „OK” pojawi się okienko z wyborem pliku.\n" +
            "Znajdź plik, kliknij go, a potem wybierz „Otwórz”.",
            "Wybór pliku z Vinted");

        statusLabel.Text = "Status: Oczekiwanie na wybór pliku HTML z Vinted...";
        Cursor.Current = Cursors.WaitCursor;

        try
        {
            // Wczytanie HTML – metoda zwraca true/false, czy operacja się powiodła
            bool htmlLoaded = await ProcessHtml();

            if (!htmlLoaded)
            {
                // Użytkownik anulował lub wystąpił błąd
                statusLabel.Text = "Status: Przerwano – nie wczytano żadnych danych z pliku HTML.";
                return;
            }

            if (transactions.Count == 0)
            {
                _ = MessageBox.Show(
                    this,
                    "Program nie znalazł żadnych operacji do eksportu.\n\n" +
                    "Upewnij się, że wybrałaś plik *.html z historią operacji Vinted (np. z zakładki „Portfel”).",
                    "Brak danych",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                statusLabel.Text = "Status: Wczytano plik, ale nie znaleziono operacji.";
                return;
            }

            // KROK 2 – informacja, że zaraz będzie wybór miejsca zapisu
            ShowInfo(
                "KROK 2 z 2\n\n" +
                "Dane z pliku zostały wczytane.\n\n" +
                "Za chwilę program zapyta, gdzie zapisać raport.\n" +
                "W następnym oknie wybierz miejsce (np. Pulpit), nazwę pliku\n" +
                "i kliknij „Zapisz”.",
                "Wybór miejsca zapisania pliku");

            DateTime reportMonth = DetermineReportMonthOrNow();

            if (radioButton1.Checked)
            {
                statusLabel.Text = "Status: Przygotowywanie raportu Excel...";
                ExportExcel(reportMonth);
            }
            else
            {
                statusLabel.Text = "Status: Przygotowywanie raportu PDF...";
                ExportPdf(reportMonth);
            }

            // Po tym momencie komunikaty końcowe są już w ExportExcel/ExportPdf
            // (pytanie czy otworzyć plik + statusLabel).
        }
        finally
        {
            transactions.Clear();
            Cursor.Current = Cursors.Default;
        }
    }

    /// <summary>
    /// Wczytuje i parsuje plik HTML.
    /// Zwraca:
    ///  - true  – jeśli udało się wczytać plik (nawet jeśli nie ma w nim operacji),
    ///  - false – jeśli użytkownik anulował lub wystąpił błąd.
    /// </summary>
    private async Task<bool> ProcessHtml()
    {
        using OpenFileDialog ofd = new()
        {
            Filter = "Pliki HTML (*.html)|*.html",
            Title = "Wybierz plik HTML z historią operacji Vinted"
        };

        DialogResult dialogResult = ofd.ShowDialog();
        if (dialogResult != DialogResult.OK)
        {
            // Użytkownik kliknął „Anuluj”
            transactions.Clear();
            _ = MessageBox.Show(
                this,
                "Nie wybrano żadnego pliku.\n\n" +
                "Jeśli chcesz wygenerować raport, kliknij przycisk jeszcze raz i wybierz plik.",
                "Operacja przerwana",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            return false;
        }

        statusLabel.Text = "Status: Wczytywanie danych z wybranego pliku HTML...";
        Application.DoEvents(); // żeby status od razu się odświeżył na ekranie

        string htmlContent;
        try
        {
            htmlContent = File.ReadAllText(ofd.FileName);
        }
        catch (Exception ex)
        {
            transactions.Clear();
            statusLabel.Text = "Status: Błąd podczas odczytu pliku HTML.";

            _ = MessageBox.Show(
                this,
                "Nie udało się odczytać wybranego pliku.\n\n" +
                "Sprawdź, czy plik istnieje i czy masz do niego dostęp.\n\n" +
                "Szczegóły techniczne (dla Kacpra):\n" + ex.Message,
                "Błąd odczytu pliku",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);

            return false;
        }

        IConfiguration config = Configuration.Default;
        IBrowsingContext context = BrowsingContext.New(config);
        AngleSharp.Dom.IDocument document = await context.OpenAsync(req => req.Content(htmlContent));

        transactions.Clear();

        // Szukamy elementów li z klasą "pile__element"
        IHtmlCollection<AngleSharp.Dom.IElement> items = document.QuerySelectorAll("li.pile__element");

        foreach (AngleSharp.Dom.IElement li in items)
        {
            AngleSharp.Dom.IElement? titleElement = li.QuerySelector(".web_ui__Cell__title");
            if (titleElement == null)
            {
                continue;
            }

            string operationTypeText = titleElement.TextContent.Trim();

            // Interesują nas tylko operacje: Sprzedane, Zakup oraz Zwrot kosztów
            if (operationTypeText is not "Sprzedane" and not "Zakup" and not "Zwrot kosztów")
            {
                continue;
            }

            AngleSharp.Dom.IElement? descElement = li.QuerySelector(".web_ui__Cell__body");
            string description = descElement != null ? descElement.TextContent.Trim() : "Brak opisu";

            AngleSharp.Dom.IElement? amountElement = li.QuerySelector("h2.web_ui__Text__text.web_ui__Text__title.web_ui__Text__right.web_ui__Text__parent");
            if (amountElement == null)
            {
                continue;
            }

            string amountText = amountElement.TextContent.Trim();
            // Usuwamy zbędne znaki – pozostawiamy cyfry oraz przecinek/kropkę
            string cleanAmount = new([.. amountText.Where(c => char.IsDigit(c) || c == ',' || c == '.')]);
            cleanAmount = cleanAmount.Replace(',', '.');

            if (!decimal.TryParse(cleanAmount, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal amount))
            {
                amount = 0;
            }

            // Wyciągamy datę – zakładamy, że data znajduje się w elemencie suffix jako ostatni węzeł tekstowy
            AngleSharp.Dom.IElement? suffixDiv = li.QuerySelector(".web_ui__Cell__suffix > div");
            string dateText = "";
            if (suffixDiv != null)
            {
                INode? textNode = suffixDiv.ChildNodes.FirstOrDefault(n => n.NodeType == NodeType.Text);
                if (textNode != null)
                {
                    dateText = textNode.TextContent.Trim();
                }
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
                {
                    amount = -amount; // zakupy – ujemne
                }
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

        statusLabel.Text = $"Status: Znaleziono {transactions.Count} operacji w pliku: {Path.GetFileName(ofd.FileName)}";

        if (transactions.Count == 0)
        {
            _ = MessageBox.Show(
                this,
                "W wybranym pliku nie znaleziono żadnych operacji Vinted.\n\n" +
                "Upewnij się, że jest to właściwy plik z historią operacji.",
                "Brak operacji w pliku",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        return true;
    }

    private void ExportExcel(DateTime reportMonth)
    {
        statusLabel.Text = $"Status: Eksport danych do pliku Excel {reportMonth:MMMM yyyy}...";

        string monthTitle = reportMonth.ToString("MMMM yyyy", Pl);

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
        IXLWorksheet ws = workbook.Worksheets.Add($"Raport {reportMonth:yyyy-MM}");

        // Główny nagłówek – wiersz 1
        ws.Cell("A1").Value = $"Raport operacji Vinted — {monthTitle}";
        _ = ws.Range("A1:K1").Merge();
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

        const int startDataRow = 5;
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

        _ = ws.Columns().AdjustToContents();

        SaveFileDialog sfd = new()
        {
            Filter = "Plik Excel (*.xlsx)|*.xlsx",
            Title = "Wybierz miejsce zapisania raportu Excel",
            FileName = $"vinted {monthTitle}".ToLower(Pl)
        };

        DialogResult saveResult = sfd.ShowDialog();
        if (saveResult != DialogResult.OK)
        {
            // Użytkownik zamknął okno / kliknął Anuluj
            statusLabel.Text = "Status: Zapis pliku Excel został przerwany – nie wybrano miejsca zapisania pliku.";

            _ = MessageBox.Show(
                this,
                "Zapis raportu został przerwany.\n\n" +
                "Jeśli chcesz zapisać raport, uruchom operację ponownie\n" +
                "i na koniec kliknij przycisk „Zapisz”.",
                "Zapis przerwany",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);

            return;
        }

        try
        {
            workbook.SaveAs(sfd.FileName);

            statusLabel.Text = $"Status: Zapisano plik Excel w {sfd.FileName}";

            // Tylko jeden komunikat końcowy: zapytanie o otwarcie
            DialogResult result = MessageBox.Show(
                this,
                "Plik Excel został zapisany.\n\nCzy chcesz go teraz otworzyć?",
                "Sukces",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                _ = Process.Start(new ProcessStartInfo
                {
                    FileName = sfd.FileName,
                    UseShellExecute = true
                });
            }
        }
        catch (Exception)
        {
            transactions.Clear();
            statusLabel.Text = "Status: Ostatni eksport danych do pliku Excel zakończony niepowodzeniem.";

            _ = MessageBox.Show(
                this,
                "Nie udało się otworzyć/zapisać pliku.\n\n" +
                "Zadzwoń do Kacpra, wyślij mu tekst z okienka, które pojawi się po tym jak wciśniesz OK " +
                "i rozwiniesz szczegóły w kolejnym okienku.",
                "Błąd",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);

            throw;
        }
    }

    private void ExportPdf(DateTime reportMonth)
    {
        if (transactions.Count == 0)
        {
            _ = MessageBox.Show("Brak danych do eksportu.", "Błąd",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        statusLabel.Text = $"Status: Eksport danych do pliku PDF {reportMonth:MMMM yyyy}...";


        // Polskie nazwy miesiąca, np. "kwiecień 2025"
        string monthTitle = reportMonth.ToString("MMMM yyyy", new CultureInfo("pl-PL"));

        // Obliczamy sumy i bilans
        decimal purchaseSum = Math.Abs(transactions
                                .Where(t => t.Type == TransactionType.Purchase)
                                .Sum(t => t.Amount));
        decimal soldSum = transactions
                                .Where(t => t.Type == TransactionType.Sold)
                                .Sum(t => t.Amount);
        decimal refundSum = transactions
                                .Where(t => t.Type == TransactionType.Refund)
                                .Sum(t => t.Amount);
        decimal balance = soldSum + refundSum - purchaseSum;

        using var sfd = new SaveFileDialog
        {
            Filter = "Plik PDF (*.pdf)|*.pdf",
            Title = "Wybierz miejsce zapisania raportu PDF",
            FileName = $"vinted {monthTitle}".ToLower(Pl)
        };

        DialogResult saveResult = sfd.ShowDialog();
        if (saveResult != DialogResult.OK)
        {
            statusLabel.Text = "Status: Zapis pliku PDF został przerwany – nie wybrano miejsca zapisania pliku.";

            _ = MessageBox.Show(
                this,
                "Zapis raportu został przerwany.\n\n" +
                "Jeśli chcesz zapisać raport, uruchom operację ponownie\n" +
                "i na koniec kliknij przycisk „Zapisz”.",
                "Zapis przerwany",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);

            return;
        }

        try
        {
            Document.Create(container =>
            {
                _ = container.Page(page =>
                {
                    page.Size(PageSizes.A4);
                    page.Margin(2, Unit.Centimetre);
                    page.DefaultTextStyle(x => x.FontSize(12));

                    // HEADER
                    page.Header()
                        .Height(80)
                        .Padding(5)
                        .Column(col =>
                        {
                            col.Spacing(2);

                            // Tytuł i miesiąc
                            _ = col.Item().Text("Raport operacji Vinted")
                                      .SemiBold()
                                      .FontSize(20)
                                      .AlignCenter();
                            _ = col.Item().Text(monthTitle)
                                      .FontSize(14)
                                      .AlignCenter();

                            // Sumy i bilans w jednym wierszu
                            col.Item().PaddingTop(5)
                                      .Row(row =>
                                      {
                                          _ = row.RelativeItem().Text($"Zakupione: {purchaseSum:0.00} zł")
                                                         .FontSize(11);
                                          _ = row.RelativeItem().Text($"Sprzedane: {soldSum:0.00} zł")
                                                         .FontSize(11);
                                          _ = row.RelativeItem().Text($"Zwroty: {refundSum:0.00} zł")
                                                         .FontSize(11);
                                          _ = row.RelativeItem().Text($"Bilans: {balance:0.00} zł")
                                                         .FontSize(11)
                                                         .SemiBold();
                                      });
                        });

                    // CONTENT: sekcje na osobnych stronach
                    page.Content()
                        .PaddingVertical(10)
                        .Column(col =>
                        {
                            AddSection(col, "Zakupione", TransactionType.Purchase, "#ADD8E6");
                            col.Item().PageBreak();

                            AddSection(col, "Sprzedane", TransactionType.Sold, "#C6EFCE");
                            col.Item().PageBreak();

                            AddSection(col, "Zwrot kosztów", TransactionType.Refund, "#FCE4D6");
                        });

                    // FOOTER: data + numeracja stron
                    page.Footer()
                        .Height(30)
                        .PaddingHorizontal(5)
                        .Row(row =>
                        {
                            // Lewa: data wygenerowania
                            row.RelativeItem()
                               .AlignLeft()
                               .Text(txt =>
                               {
                                   _ = txt.Span("Wygenerowano: ").FontSize(9);
                                   _ = txt.Span(DateTime.Now.ToString("dd.MM.yyyy HH:mm"))
                                      .FontSize(9);
                               });

                            // Prawa: numeracja stron X / Y
                            row.ConstantItem(100)
                               .AlignRight()
                               .Text(txt =>
                               {
                                   _ = txt.CurrentPageNumber();
                                   _ = txt.Span(" / ");
                                   _ = txt.TotalPages()
                                   .FontSize(9);
                               });
                        });
                });
            }).GeneratePdf(sfd.FileName);

            statusLabel.Text = $"Status: Zapisano plik PDF w {sfd.FileName}";

            DialogResult result = MessageBox.Show(
                this,
                "Plik PDF został zapisany.\n\nCzy chcesz go teraz otworzyć?",
                "Sukces",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                _ = Process.Start(new ProcessStartInfo
                {
                    FileName = sfd.FileName,
                    UseShellExecute = true
                });
            }
        }
        catch (Exception)
        {
            transactions.Clear();
            statusLabel.Text = "Status: Ostatni eksport danych do pliku PDF zakończony niepowodzeniem.";

            _ = MessageBox.Show(
                this,
                "Nie udało się otworzyć/zapisać pliku.\n\n" +
                "Zadzwoń do Kacpra, wyślij mu tekst z okienka, które pojawi się po tym jak wciśniesz OK " +
                "i rozwiniesz szczegóły w kolejnym okienku.",
                "Błąd",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);

            throw;
        }
    }

    // Metoda dodająca kolorowany nagłówek sekcji i tabelę
    private void AddSection(ColumnDescriptor column,
                            string title,
                            TransactionType type,
                            string headerHexColor)
    {
        var list = transactions.Where(t => t.Type == type).ToList();
        if (list.Count == 0)
        {
            return;
        }

        // Nagłówek sekcji (czarny tekst na jasnym tle)
        _ = column.Item()
              .Padding(5)
              .Background(Color.FromHex(headerHexColor))
              .Text(title)
              .SemiBold()
              .FontSize(14)
              .FontColor(Colors.Black);

        // Tabela: Opis | Kwota | Data
        column.Item().Table(table =>
        {
            table.ColumnsDefinition(cols =>
            {
                cols.RelativeColumn();
                cols.ConstantColumn(80);
                cols.ConstantColumn(80);
            });

            // Nagłówki tabeli
            table.Header(header =>
            {
                _ = header.Cell().Background(Colors.Grey.Lighten3)
                             .Padding(3)
                             .Text("Opis").SemiBold()
                             .FontColor(Colors.Black);
                _ = header.Cell().Background(Colors.Grey.Lighten3)
                             .Padding(3)
                             .AlignRight()
                             .Text("Kwota (zł)").SemiBold()
                             .FontColor(Colors.Black);
                _ = header.Cell().Background(Colors.Grey.Lighten3)
                             .Padding(3)
                             .AlignRight()
                             .Text("Data").SemiBold()
                             .FontColor(Colors.Black);
            });

            // Wiersze danych
            foreach (Transaction? t in list)
            {
                _ = table.Cell().Padding(3).Text(t.Description ?? "")
                       .FontColor(Colors.Black);
                _ = table.Cell().Padding(3).AlignRight().Text($"{t.Amount:0.00}")
                       .FontColor(Colors.Black);
                _ = table.Cell().Padding(3).AlignRight().Text(t.Date ?? "")
                       .FontColor(Colors.Black);
            }
        });
    }

    private void radioButton1_CheckedChanged(object sender, EventArgs e)
    {
        statusLabel.Text = radioButton1.Checked
            ? "Status: Wybrano raport Excel (do otwarcia w programie typu Excel)."
            : "Status: Wybrano raport PDF (do podglądu lub wydruku).";
    }

    private static readonly CultureInfo Pl = new("pl-PL");

    private DateTime DetermineReportMonthOrNow()
    {
        // Spróbuj sparsować daty transakcji i wybierz (Year, Month) z największą liczbą wpisów.
        // Gdy nic nie wyjdzie – bieżący miesiąc.
        DateTime now = DateTime.Now;
        var parsed = transactions
            .Select(t => TryParsePolishDate(t.Date, now))
            .Where(d => d.HasValue)
            .Select(d => d!.Value)
            .ToList();

        if (parsed.Count == 0)
        {
            return new DateTime(now.Year, now.Month, 1);
        }

        var topGroup = parsed
            .GroupBy(d => new { d.Year, d.Month })
            .OrderByDescending(g => g.Count())
            .ThenByDescending(g => g.Key.Year)
            .ThenByDescending(g => g.Key.Month)
            .First();

        return new DateTime(topGroup.Key.Year, topGroup.Key.Month, 1);
    }

    private static DateTime? TryParsePolishDate(string? raw, DateTime now)
    {
        if (string.IsNullOrWhiteSpace(raw))
        {
            return null;
        }

        string s = raw.Trim().ToLower(Pl);

        // proste relatywne
        if (s == "dzisiaj")
        {
            return now.Date;
        }

        if (s == "wczoraj")
        {
            return now.Date.AddDays(-1);
        }

        // popularne formaty z Vinted i podobne
        string[] formats =
        [
            "d MMMM yyyy",  // 9 maja 2025
            "d MMMM",       // 9 maja  (uzupełnimy rok)
            "dd.MM.yyyy",
            "d.M.yyyy",
            "dd.MM.yy",
            "d MMM yyyy"
        ];

        if (DateTime.TryParseExact(s, formats, Pl, DateTimeStyles.AllowWhiteSpaces, out DateTime dt))
        {
            if (dt.Year <= 1)
            {
                dt = new DateTime(now.Year, dt.Month, dt.Day);
            }

            return dt;
        }

        return DateTime.TryParse(s, Pl, DateTimeStyles.AllowWhiteSpaces, out dt) ? dt : null;
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
