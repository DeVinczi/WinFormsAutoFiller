using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using FormFiller.Constants;
using FormFiller.Helpers;
using FormFiller.Models;
using FormFiller.Models.TrainingEntity;
using FormFiller.Models.WorkerEntity;

using Newtonsoft.Json;

using OfficeOpenXml;

using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

using SeleniumExtras.WaitHelpers;

using System.Data;
using System.Text;

using WinFormsAutoFiller.Helpers;
using WinFormsAutoFiller.Models.KrsEntity;
using WinFormsAutoFiller.Models.PytaniaDoWnioskuEntity;
using WinFormsAutoFiller.Services.FormParts;
using WinFormsAutoFiller.Utilis;

using static WinFormsAutoFiller.Helpers.SeleniumHelpers;

namespace FormFiller.Services;

public interface ISeleniumService
{
    string DownloadAypRis();
    void LoginToPage();
    Task LoadForm(string businessAddresses, string city,
        string pkd, string nrRachunku, int liczbaZatrudnionych,
        ContactPerson contactPerson, DateTime startDate, DateTime endDate, string[] paths,
        string krs, string ceidg, string ris, string aypNip, Adres siedziba, string krsValue,
        bool czyRezerwowaPulaKFS, bool czyJestMikro);
    Task<string> DownloadComapnyByNipPdf(string nip);
    void ProvideWorkerInformation(ExcelWorkersRow row, string city);
    void LoadPlanowanyRealizator();
    void LoadInfomacjeDotyczaceKsztalcenia(ExcelWorksheet row, DataTable listaSzkolen, DataTable pracownicy, string uzasadnienie, Dictionary<string, string> wordHours);
    void LoadEndOfForm();
    void ValidateWorkers();
    void ZrobWydruk();
    void LoadZałączniki(string[] paths, string krsPath, string ceidgPath, string risPath, string aypNip);
}

public class SeleniumService : ISeleniumService
{
    private readonly IDriverService _driverService;
    private ChromeDriver _chromeDriver;
    private WebDriverWait _webDriverWait;

    public SeleniumService()
    {
        _driverService = DriverService.Instance;
        (_chromeDriver, _webDriverWait) = _driverService.GetDriver();
    }

    public string DownloadAypRis()
    {
        try
        {
            InitializeDriver();
            _chromeDriver.Navigate().GoToUrl("https://stor.praca.gov.pl/api/ris/szczegoly/pdf/60002");
            var path = Path.Combine(Path.GetTempPath() + "ayp");
            var files = Directory.GetFiles(path);

            if (files.Length == 0)
            {
                throw new Exception("AYP_RIS nie został pobrany!");
            }

            var latestFile = files
                   .Select(file => new FileInfo(file))
                   .OrderByDescending(fileInfo => fileInfo.LastWriteTime)
                   .First();

            string newFileName = Path.Combine(latestFile.DirectoryName, $"AYP_RIS" + ".pdf");
            File.Move(latestFile.FullName, newFileName);
            return newFileName;
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
            throw;
        }
    }

    public void LoginToPage()
    {
        try
        {
            InitializeDriver();
            if (IsElementPresent(By.CssSelector("header.govpl")))
            {
                Console.WriteLine("User is already logged in. Skipping login process.");
                return;
            }
            if (IsElementPresent(By.CssSelector("div.zielona-linia-kontener")))
            {
                return;
            }

            _chromeDriver.Navigate().GoToUrl("https://www.praca.gov.pl/api/eurzad/oauth2/authorization/login-praca-gov-pl");

            var loginButton = By.LinkText("Zaloguj się");
            _webDriverWait.Until(ExpectedConditions.ElementIsVisible(loginButton));

            _chromeDriver.FindElement(loginButton).Click();
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }

    private bool IsElementPresent(By by)
    {
        try
        {
            _chromeDriver.FindElement(by);
            return true;
        }
        catch (NoSuchElementException)
        {
            return false;
        }
    }

    public async Task LoadForm(string businessAddresses, string city, string pkd, string nrRachunku, int liczbaZatrudnionych, ContactPerson contactPerson, DateTime startDate, DateTime endDate,
        string[] paths, string krs, string ceidg, string ris, string aypNip, Adres siedziba, string krsValue, bool czyRezerwowaPulaKFS, bool czyJestMirko)
    {
        try
        {
            var wait = new WebDriverWait(_chromeDriver, TimeSpan.FromMinutes(5));
            try
            {
                InitializeDriver();
                var startForm = By.CssSelector("div.formularz fieldset");
                var uczestnicyForm = By.CssSelector("div[data-psz-kfs-uczestnicy]");

                wait.Until(ExpectedConditions.ElementExists(startForm));

                ExecuteKursyRadioButton(_chromeDriver, _webDriverWait);

                ErrorHandler.CatchError(() =>
                _chromeDriver.ExecuteScript(
                    "document.querySelector('button[data-ng-click=\"zapiszForme()\"]').click()"));

                wait.Until(ExpectedConditions.ElementIsVisible(uczestnicyForm));

                ErrorHandler.CatchError(() =>
                    ExecuteInputByName(_chromeDriver, _webDriverWait, "dane.miejscowosc", city)
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                $"{ex.Message}",
                "Uwaga",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            }

            var wnioskodawcaService = new WnioskodawcaService();
            // 1. CZĘŚĆ
            wnioskodawcaService.DaneIdentyfikacyjne_1_1(krsValue);
            wnioskodawcaService.AdresSiedziby_1_2();
            await wnioskodawcaService.MiejsceProwadzeniaDziałalności_1_3(businessAddresses, city);
            wnioskodawcaService.AdresDoKorespondencji_1_4();
            wnioskodawcaService.OznaczeniePKD_1_5(pkd);
            wnioskodawcaService.NumerRachunkuBankowego_1_6(nrRachunku);
            wnioskodawcaService.LiczbaZatrudnionychPracowikow_1_7(liczbaZatrudnionych, czyJestMirko);
            wnioskodawcaService.OsobaUprawnionaDoKontaktuZUrzedem_1_8(contactPerson);
            if (!czyRezerwowaPulaKFS)
            {
                _chromeDriver.ExecuteScript("document.querySelector(\"input[data-ng-model='dane.koszty.zrodloSrodkow']\").click()");
            }
            else
            {
                _chromeDriver.ExecuteScript("document.querySelector(\"input[value='REZERWA']\").click()");
            }

            ExecuteInputByName(_chromeDriver, _webDriverWait, "dane.termin.dataOd", startDate.ToShortDateString());
            ExecuteInputByName(_chromeDriver, _webDriverWait, "dane.termin.dataDo", endDate.ToShortDateString());


            LoadEndOfForm();
            LoadZałączniki(paths, krs, ceidg, ris, aypNip);

            var addWorkerButton = By.Id("tabela-uczestnik");
            wait.Until(ExpectedConditions.ElementIsVisible(addWorkerButton));
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
            throw;
        }
    }

    public void ProvideWorkerInformation(ExcelWorkersRow row, string city)
    {
        try
        {
            var isWorking = row.Data.TryGetValue(WorkersFormKeys.Nieobecny, out var workersForm);
            if (!isWorking)
            {
                return;
            }

            InitializeDriver();

            ExecuteInputByName(_chromeDriver, _webDriverWait, "dane.miejscowosc", city);

            _chromeDriver.ExecuteScript(
                "document.querySelector('div#tabela-uczestnik div.button-container button[data-ng-show=\"options.addButton\"]').click()");

            var kodUczestnika = _chromeDriver.FindElement(By.Id("kod-uczestnika"));
            var number = row.Data[WorkersFormKeys.Number];
            kodUczestnika.SendKeys((string)number);
            _chromeDriver.ExecuteScript("document.querySelector(\"input[value='PRACOWNIK']\").click()");

            var wait = new WebDriverWait(_chromeDriver, TimeSpan.FromMilliseconds(100));
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//div[contains(text(), 'Wiek')]")));
            _chromeDriver.ExecuteScript(
                "document.evaluate(\"//label[.//div[contains(text(),'Wiek')]]//div\", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click()");
            _chromeDriver.ExecuteScript(
                $"document.evaluate('//ul[contains(@class, \"ui-autocomplete\")]//li/a[text()=\"{Age.GetAge(int.Parse(((string?)row.Data[WorkersFormKeys.Wiek])!))}\"]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click()");

            _chromeDriver.ExecuteScript(
                "document.evaluate(\"//label[.//div[contains(text(),'Poziom wykształcenia')]]//div\", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click()");

            _chromeDriver.ExecuteScript(
                $"document.evaluate('//ul[contains(@class, \"ui-autocomplete\")]//li/a[text()=\"{Education.MapEducationLevel((string)row.Data[WorkersFormKeys.PoziomWyksztalcenia])}\"]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click()");

            _chromeDriver.ExecuteScript(
                "document.evaluate(\"//label[.//div[contains(text(),'Płeć')]]//div\", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click()");

            _chromeDriver.ExecuteScript(
                $"document.evaluate('//ul[contains(@class, \"ui-autocomplete\")]//li/a[text()=\"{Gender.MapGender((string)row.Data[WorkersFormKeys.Plec])}\"]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click()");

            _chromeDriver.ExecuteScript(
                "document.evaluate('//div[contains(@class, \"sg-autocomplete\") and .//input[contains(@data-ng-model, \"uczestnik.priorytet\")]]//span', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click()");
            var wait2 = new WebDriverWait(_chromeDriver, TimeSpan.FromMilliseconds(3000));

            ErrorHandler.CatchError(() =>
            {
                wait2.Until(ExpectedConditions.ElementIsVisible(By.XPath(
                    "//div[contains(@class, \"sg-autocomplete\") and .//input[contains(@data-ng-model, \"uczestnik.priorytet\")]]//span")));

                wait2.Until(ExpectedConditions.ElementToBeClickable(By.XPath(
                    $"//a[contains(text(), \"{KFSPriority.MapPriority((string)row.Data[WorkersFormKeys.Priorytet])}\")]")));
                //DO WYBRANIA:
                _chromeDriver.ExecuteScript(
                    $"document.evaluate('//a[contains(text(), \"{KFSPriority.MapPriority((string)row.Data[WorkersFormKeys.Priorytet])}\")]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click()");
            });

            _chromeDriver.ExecuteScript(
                "document.querySelector('input[data-ng-model=\"uczestnik.dofinansowanieKfs.wniosekZlozony\"][data-ng-value=\"false\"]').click()");
            //Wysokosc przyznanego w biezacym roku na tego uczestnika

            ErrorHandler.CatchError(() =>
            {
                ExecuteInputByName(_chromeDriver, _webDriverWait, "uczestnik.dofinansowanieKfs.przyznanaKwota",
                    PrzyznanaKwota.EnsureTwoDecimalPlaces(
                        (string)row.Data[WorkersFormKeys.KwotaOtrzymanegoDofinansowania]));
            });

            //Checkbox

            ExecuteCheckbox(_chromeDriver, _webDriverWait, "uczestnik.typKsztalcenia.awansZawodowy");
            ExecuteCheckbox(_chromeDriver, _webDriverWait, "uczestnik.typKsztalcenia.rozszerzenieObowiazkow");
            ExecuteCheckbox(_chromeDriver, _webDriverWait, "uczestnik.typKsztalcenia.kompetencjeZawodowe");


            //TO DO: DO SPRAWDZENIA
            if (DateTime.TryParse(row.Data[WorkersFormKeys.OkresZatrudnieniaDo].ToString(), out var parsed))
            {
                ExecuteCheckbox(_chromeDriver, _webDriverWait, "uczestnik.typKsztalcenia.przedluzenieZatrudnienia");
            }
            else
            {
                ExecuteCheckbox(_chromeDriver, _webDriverWait, "uczestnik.typKsztalcenia.utrzymanieZatrudnienia");
            }

            _chromeDriver.ExecuteScript("document.querySelector('textarea[data-ng-model=\"uczestnik.uzasadnieniePotrzeb\"]').value = 'Do uzupełnienia'");
            _chromeDriver.ExecuteScript("document.querySelector('textarea[data-ng-model=\"uczestnik.uzasadnieniePotrzeb\"]').dispatchEvent(new Event('input', { bubbles: true }))");

            //SAVE
            _chromeDriver.ExecuteScript(
                "document.querySelector('button[data-ng-click=\"zapiszUczestnika()\"]').click()");
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }

    public void LoadPlanowanyRealizator()
    {
        try
        {
            InitializeDriver();
            _chromeDriver.ExecuteScript(
                "document.querySelector('div#tabela-realizator div.button-container button[data-ng-show=\"options.addButton\"]').click()");
            Wait();
            var kodRealizatora = _chromeDriver.FindElement(By.Name("kodRealizatora"));
            kodRealizatora.SendKeys("AYP");

            ExecuteInputByName(_chromeDriver, _webDriverWait, "realizator.nazwa", "AYP Spółka z ograniczoną odpowiedzialnością");

            _chromeDriver.ExecuteScript("document.querySelector(\"input[value='KRAJOWY']\").click()");

            var inputElement = _chromeDriver.FindElements(By.CssSelector("[data-ng-model='kodPocztowy']"));

            string kodPocztowy = inputElement.FirstOrDefault(x => x.Displayed).GetAttribute("id");

            // Now use the dynamic ID to interact with the input field if needed
            var inputWithDynamicId = _chromeDriver.FindElement(By.Id(kodPocztowy));
            inputWithDynamicId.SendKeys("83-000");
            _chromeDriver.ExecuteScript(
                $$"""
                document.querySelector('input#{{kodPocztowy}}').dispatchEvent(new Event('input', {bubbles: true }))
                """
            );
            ExecuteInputByName(_chromeDriver, _webDriverWait, "realizator.adres.adresKrajowy.ulica", "Chopina");
            ExecuteInputByName(_chromeDriver, _webDriverWait, "realizator.adres.adresKrajowy.nrDomu", "28");
            ExecuteInputByName(_chromeDriver, _webDriverWait, "realizator.adres.adresKrajowy.nrLokalu", "7");
            ExecuteInputByName(_chromeDriver, _webDriverWait, "realizator.nip", "5842774905");
            ExecuteInputByName(_chromeDriver, _webDriverWait, "realizator.krs", "0000741233");
            ExecuteInputByName(_chromeDriver, _webDriverWait, "realizator.regon", "380820190");

            ExecuteInput(_chromeDriver, _webDriverWait, "realizator.pkd");

            Wait();
            SelectAutocompleteOptionWithRetryW(_chromeDriver, _webDriverWait, "8559B");

            _chromeDriver.ExecuteScript(
                """
               const selectElement = document.querySelector("select[data-ng-model='wybieranyCertyfikat']");

               selectElement.value = "13";
               selectElement.dispatchEvent(new Event("change", { bubbles: true }));
               selectElement.value = "11";
               selectElement.dispatchEvent(new Event("change", { bubbles: true }));
               selectElement.value = "8";
               selectElement.dispatchEvent(new Event("change", { bubbles: true })); 
               """);

            _chromeDriver.ExecuteScript(
                """
               document.evaluate('//input[contains(@data-ng-model, "realizator.certyfikaty.inny")]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click()
               """);

            var usługi = "DUNS, KRAZ, BUR";
            _chromeDriver.ExecuteScript(
                $"""
               document.evaluate('//textarea[contains(@data-ng-model, "realizator.certyfikaty.innyOpis")]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.value = '{usługi}'
               """
            );
            _chromeDriver.ExecuteScript(
                "document.evaluate('//textarea[contains(@data-ng-model, \"realizator.certyfikaty.innyOpis\")]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.dispatchEvent(new Event('input', { bubbles: true }));");
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }

    public void LoadInfomacjeDotyczaceKsztalcenia(ExcelWorksheet row, DataTable listaSzkolen, DataTable pracownicy, string uzasadnienie, Dictionary<string, string> wordHours)
    {
        try
        {
            InitializeDriver();
            var liczbaDni = listaSzkolen.Rows[0][3];
            for (int i = 1; i < listaSzkolen.Rows.Count; i++)
            {
                _chromeDriver.ExecuteScript(
                    """
                    document.querySelector('div#tabela-realizator-dzialania div.button-container button[data-ng-show=\"options.addButton\"]').click()
                    """);
                Wait();

                var szkolenie = listaSzkolen.Rows[i];

                var keyWord = szkolenie[TrainingFormKeys.TematSzkolenia].ToString()![..10];

                var kodDzialania = _chromeDriver.FindElement(By.Id("kod-dzialania"));
                kodDzialania.SendKeys(keyWord);

                ExecuteRadioButton(_chromeDriver, _webDriverWait, "KURS");
                var kształcenie = szkolenie[TrainingFormKeys.TematSzkolenia].ToString();

                ExecuteTextArea(_chromeDriver, _webDriverWait, "dzialanie.nazwa", kształcenie);
                ExecuteDropdown(_chromeDriver, _webDriverWait, "slowniki/OBSZARY_SZKOLEN", "dzialanie.tematyka", "Kompetencje");
                ExecuteCheckbox(_chromeDriver, _webDriverWait, "dzialanie.kompetencje.potwierdzenie.certyfikat");

                //radio button
                _chromeDriver.ExecuteScript(
                    """
                    document.querySelector("input[data-ng-value='false'][data-ng-model='dzialanie.kompetencje.podstawaPrawna']").click()
                    """);

                //Liczba z excela
                var wartoscSzkolenia =
                    PrzyznanaKwota.EnsureTwoDecimalPlaces(szkolenie[TrainingFormKeys.WartośćSzkoleniaNaOsobe].ToString());


                // Find the index and the matching key-value pair
                var matchingIndex = wordHours
                    .Select((kvp, index) => new { KeyValue = kvp.Key[(kvp.Key.LastIndexOf('\\') + 1)..], Index = kvp.Value }) // Project kvp with its index
                    .FirstOrDefault(x => x.KeyValue.StartsWith($"Program - {keyWord}", StringComparison.OrdinalIgnoreCase));
                //zmienie na wielkosc liter nie ma znaczenia
                //szukac po kontencie w pliku Program - 
                //nazwa nie dłuższa 

                if (matchingIndex is null)
                {
                    MessageBox.Show("Plik z liczbą godzin jest nieprawidłowo nazwany.");
                }

                ExecuteInputByName(_chromeDriver, _webDriverWait, "dzialanie.liczbaGodzin.ilosc", matchingIndex.Index);
                ExecuteInputByName(_chromeDriver, _webDriverWait, "dzialanie.cenaNetto", wartoscSzkolenia);
                ExecuteInputByName(_chromeDriver, _webDriverWait, "dzialanie.cenaBrutto", wartoscSzkolenia);
                _chromeDriver.ExecuteScript("document.querySelector('button[data-ng-disabled=\"dzialanie.porownanieOfert.length >= 20\"]').click()");
                ExecuteInputByName(_chromeDriver, _webDriverWait, "porownanie.nazwa", "Brak");
                ExecuteInputByName(_chromeDriver, _webDriverWait, "porownanie.liczbaGodzin", "1");
                ExecuteInputByName(_chromeDriver, _webDriverWait, "porownanie.cenaNetto", "0,00");
                ExecuteInputByName(_chromeDriver, _webDriverWait, "porownanie.cenaBrutto", "0,00");

                var doc = ReadWordToString(uzasadnienie);

                string escapedDoc = JsonConvert.SerializeObject(doc.Substring(0, 1999)).Trim();

                _chromeDriver.ExecuteScript($@"
                var textArea = document.querySelector('textarea[data-ng-model=""dzialanie.uzasadnienie""]');
                if (textArea) {{
                    textArea.value = {escapedDoc};  // Set the value of the textarea
                    var event = new Event('input', {{
                        'bubbles': true,
                        'cancelable': true
                    }});
                    textArea.dispatchEvent(event); 
                }}");

                _chromeDriver.ExecuteScript("document.querySelector('textarea[data-ng-model=\"dzialanie.uzasadnienie\"]').dispatchEvent(new Event('input', { bubbles: true }))");
                _chromeDriver.ExecuteScript("document.querySelector('textarea[data-ng-model=\"dzialanie.uzasadnienie\"]').dispatchEvent(new Event('input', { bubbles: true }))");

                PracownikPopupComparation(row, pracownicy, kształcenie);
            }

            _chromeDriver.ExecuteScript("""
                                        document.querySelector('div.panel-kontener button[data-ng-click="zapiszRealizatora()"]').click()
                                        """);
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex);
        }
    }
    public static string ReadWordToString(string filePath)
    {
        StringBuilder text = new StringBuilder();

        using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, false))
        {
            var body = wordDocument.MainDocumentPart.Document.Body;

            foreach (var paragraph in body.Elements<Paragraph>())
            {
                foreach (var run in paragraph.Elements<Run>())
                {
                    foreach (var textElement in run.Elements<Text>())
                    {
                        text.Append(textElement.Text);
                    }
                }
                text.AppendLine();
            }
        }

        return text.ToString();
    }

    private void PracownikPopupComparation(ExcelWorksheet row, DataTable pracownicy, string nazwa)
    {
        try
        {
            var getPracownicy = MapColumnByKey(pracownicy, "imię i nazwisko", nazwa).Where(x => x.Value == "1").ToDictionary();
            var workersList = new List<ExcelWorkersRow>();

            var worksheet = row.Workbook.Worksheets[0];
            var rowCount = worksheet.Dimension.Rows;
            var colCount = worksheet.Dimension.Columns;

            // Extract headers
            var headers = new List<string>();
            for (int col = 1; col <= colCount; col++)
            {
                headers.Add(worksheet.Cells[1, col].Text.Trim()); // Assume first row is the header
            }

            // Process all rows starting from the second row (skip headers)
            for (int r = 2; r <= rowCount; r++)
            {
                var excelRow = new ExcelWorkersRow();

                for (int col = 1; col <= colCount; col++)
                {
                    var header = headers[col - 1];
                    var cellValue = worksheet.Cells[r, col].Text.Trim();
                    excelRow.Data[header] = cellValue;
                }

                workersList.Add(excelRow);
            }

            _chromeDriver.ExecuteScript(
                """
            document.querySelector('div#tabela-dzialanie-uczestnik .group-right button[data-ng-click="button.handler()"]').click()
            """
            );

            foreach (var pracownicyKey in getPracownicy.Keys)
            {
                var matchedRecords = workersList
                    .Where(worker => (string)worker.Data.GetValueOrDefault(WorkersFormKeys.NazwiskoImie)! == pracownicyKey)
                    .ToList();

                if (matchedRecords.Any())
                {
                    foreach (var matchedRecord in matchedRecords)
                    {
                        // Log or use the matched record
                        Console.WriteLine($"Matched Key: {pracownicyKey}");
                        foreach (var kvp in matchedRecord.Data)
                        {
                            Console.WriteLine($"  {kvp.Key}: {kvp.Value}");
                        }

                        var x = matchedRecord.Data[WorkersFormKeys.Number].ToString();

                        // Perform actions for each matched record
                        _chromeDriver.ExecuteScript(
                            $"[...document.querySelectorAll('div.panel-kontener.formularz table tbody tr')].find(tr => [...tr.querySelectorAll('td')].some(td => td.textContent.includes(\"{x}\"))).querySelector('span').click();");

                    }
                }
            }

            _chromeDriver.ExecuteScript(
                "document.querySelector('div.okno-dialogowe button[data-ng-click=\"dodajOsobyZOkna(wybraniUczestnicy, dzialanie)\"]').click()");

            _chromeDriver.ExecuteScript("""
                                    document.querySelector('div.panel-kontener button[data-ng-click="zapiszDzialanie()"]').click()
                                    """);
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }
    public void SelectAutocompleteOptionWithRetry(string val, int maxRetries = 5, int retryDelayMs = 500)
    {
        try
        {
            _chromeDriver.ExecuteScript(
                $"document.evaluate('//input[@data-sg-autocomplete=\"slowniki/PKD_2007\" and @data-ng-model=\"realizator.pkd\"]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.value = '{val}'");

            _chromeDriver.ExecuteScript(
                "document.evaluate('//input[@data-sg-autocomplete=\"slowniki/PKD_2007\" and @data-ng-model=\"realizator.pkd\"]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.dispatchEvent(new Event('input', { bubbles: true }));");
            for (int attempt = 0; attempt < maxRetries; attempt++)
            {
                try
                {
                    var listItemClickScript = $@"
                var listItem = Array.from(document.querySelectorAll(""ul.ui-autocomplete li"")).findLast(el => el.textContent.trim().startsWith(""{val}""));
                if (listItem) {{
                    listItem.click();
                    return true;
                }}
                return false;
            ";

                    bool clicked = (bool)_chromeDriver.ExecuteScript(listItemClickScript);

                    if (clicked)
                    {
                        Console.WriteLine("Autocomplete option clicked successfully.");
                        break;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Attempt {attempt + 1} failed: {ex.Message}");
                }

                Thread.Sleep(retryDelayMs);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }

    private void Wait()
    {
        _chromeDriver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(10000);
    }
    private void InitializeDriver()
    {
        if (_chromeDriver == null)
        {
            (_chromeDriver, _webDriverWait) = _driverService.GetDriver();
        }
    }

    static Dictionary<string, string> MapColumnByKey(DataTable dataTable, string keyColumn, string valueColumn)
    {
        if (!dataTable.Columns.Contains(keyColumn))
            throw new ArgumentException($"Key column '{keyColumn}' does not exist.");
        if (!dataTable.Columns.Contains(valueColumn))
            throw new ArgumentException($"Value column '{valueColumn}' does not exist.");

        var mapping = new Dictionary<string, string>();

        foreach (DataRow row in dataTable.Rows)
        {
            var key = row[keyColumn]?.ToString();
            var value = row[valueColumn]?.ToString();

            if (!string.IsNullOrEmpty(key) && !mapping.ContainsKey(key))
            {
                mapping[key] = value ?? string.Empty;
            }
        }

        return mapping;
    }

    public async Task<string> DownloadComapnyByNipPdf(string nip)
    {
        InitializeDriver();
        await _chromeDriver.Navigate().GoToUrlAsync("https://aplikacja.ceidg.gov.pl/ceidg/ceidg.public.ui/search.aspx");
        var nipInput = "MainContentForm_txtNip";
        _webDriverWait.Until(ExpectedConditions.ElementIsVisible(By.Id(nipInput)));

        var element = _chromeDriver.FindElement(By.Id(nipInput));
        element.SendKeys(nip);

        var searchButton = "MainContentForm_btnInputSearch";
        _chromeDriver.FindElement(By.Id(searchButton)).Click();

        var details = By.Id("MainContentForm_DataListEntities_hrefDetails_0");
        _webDriverWait.Until(ExpectedConditions.ElementIsVisible(details));

        _chromeDriver.FindElement(details).Click();

        var downloadButton = By.Id("MainContentForm_btnPrint");
        _webDriverWait.Until(ExpectedConditions.ElementIsVisible(downloadButton));

        _chromeDriver.FindElement(downloadButton).Click();

        var downloadPdfButton = By.Id("MainContentForm_linkDownloadG");
        _webDriverWait.Until(ExpectedConditions.ElementIsVisible(downloadPdfButton));

        _chromeDriver.FindElement(downloadPdfButton).Click();

        var downloadedFile = await GetDownloadPathToNipFile(nip);
        if (!string.IsNullOrEmpty(downloadedFile?.Error?.Message))
        {
            return string.Empty;
        }

        return downloadedFile.Value;
    }

    private async Task<Result<string, Error>> GetDownloadPathToNipFile(string nip)
    {
        var downloadPath = Path.Combine(Path.GetTempPath() + "ayp");

        try
        {
            for (int i = 0; i < 3; i++)
            {
                var files = Directory.GetFiles(downloadPath);
                var latestFile = files
                        .Select(file => new FileInfo(file))
                        .OrderByDescending(fileInfo => fileInfo.LastWriteTime)
                        .Any(x => x.LastWriteTime >= DateTime.UtcNow.AddMinutes(-3));

                if (latestFile)
                {
                    break;
                }
                await Task.Delay(500);
            }

            if (Directory.Exists(downloadPath))
            {
                var files = Directory.GetFiles(downloadPath);
                if (files.Length > 0)
                {
                    var latestFile = files
                        .Select(file => new FileInfo(file))
                        .OrderByDescending(fileInfo => fileInfo.LastWriteTime)
                        .First();

                    string newFileName = Path.Combine(latestFile.DirectoryName, $"NIP_{nip}" + ".pdf");
                    File.Move(latestFile.FullName, newFileName);
                    return newFileName;
                }
                else
                {
                    Console.WriteLine("Wystąpił bład. Nie ma plików w folderze z NIP.");
                    return Errors.BrakPlikuNipError;
                }
            }
            return Errors.BrakPlikuNipError;
        }
        catch (Exception ex)
        {
            return ex.Message;
        }
    }

    public void LoadEndOfForm()
    {
        ExecuteCheckbox(_chromeDriver, _webDriverWait, "dane.oswiadczenie.pkt1");
        _chromeDriver.ExecuteScript("document.querySelector(\"input[data-ng-model='dane.oswiadczenie.pkt2']\").click()");
        ExecuteCheckbox(_chromeDriver, _webDriverWait, "dane.oswiadczenie.pkt3");
        ExecuteCheckbox(_chromeDriver, _webDriverWait, "dane.oswiadczenie.pkt4");
        ExecuteCheckbox(_chromeDriver, _webDriverWait, "dane.oswiadczenie.pkt5");
        ExecuteCheckbox(_chromeDriver, _webDriverWait, "dane.oswiadczenie.pkt6");
    }

    public void ValidateWorkers()
    {
        try
        {
            bool flag = true;
            do
            {
                var nodes = (long)_chromeDriver.ExecuteScript("return document.querySelectorAll('div#tabela-uczestnik td.tabela-selektor').length");

                if (nodes == 0)
                {
                    break;
                }
                for (int i = 0; i <= nodes; i++)
                {
                    _webDriverWait.Until(driver =>
                        _chromeDriver.ExecuteScript("return document.querySelector('button[data-ng-click=\"editRow()\"]')") != null);

                    _chromeDriver.ExecuteScript($"document.querySelectorAll('div#tabela-uczestnik td.tabela-selektor')[{i}].click()");
                    _chromeDriver.ExecuteScript("document.querySelector('button[data-ng-click=\"editRow()\"]').click()");

                    _chromeDriver.ExecuteScript("""var value = document.querySelector('input[data-ng-model=\"uczestnik.wydatki.srodkiKfs\"]').value;var parsedDouble = parseFloat(value.replace(',', '.').trim());var percent = parsedDouble * 0.20;document.querySelector('input[data-ng-model=\"uczestnik.wydatki.wkladWlasny\"]').value = percent.toString().replace('.', ',');""");
                    _chromeDriver.ExecuteScript("document.querySelector(\"input[data-ng-model='uczestnik.wydatki.wkladWlasny']\").dispatchEvent(new Event('input', {bubbles:true}));document.querySelector('button[data-ng-click=\"zapiszUczestnika()\"]').click();");
                    if ((i + 1) % 10 == 0 || i == nodes - 1)
                    {
                        var isButtonVisible = (bool)_chromeDriver.ExecuteScript(
                        """
                            var button = document.querySelector('button[ng-click="changePage(pageIndex + 1)"]');
                            return button.getAttribute('class') === 'disabled';
                        """
                        );
                        if (isButtonVisible)
                        {
                            flag = false;
                            _chromeDriver.ExecuteScript("document.querySelector('button[data-ng-click=\"editRow()\"]').click()");
                            _chromeDriver.ExecuteScript("document.querySelector('button[data-ng-click=\"zapiszUczestnika()\"]').click();");
                            break;
                        }
                        if (!isButtonVisible)
                        {
                            _chromeDriver.ExecuteScript("document.querySelector('button[ng-click=\"changePage(pageIndex + 1)\"]').click()");
                            _webDriverWait.Until(driver =>
                                (bool)_chromeDriver.ExecuteScript("return document.readyState === 'complete';")
                            );
                            var isButtonVisible2 = (bool)_chromeDriver.ExecuteScript(
                                """
                                var button = document.querySelector('button[ng-click="changePage(pageIndex + 1)"]');
                                return button.getAttribute('class') === 'disabled';
                                """);
                            if (isButtonVisible2)
                            {
                                break;
                            }
                        }
                    }
                }

            } while (flag);
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }


    public void ZrobWydruk()
    {
        _chromeDriver.ExecuteScript("document.querySelector(\"button[data-ng-click='$parent.wydruk()']\").click()");
    }

    public void LoadZałączniki(string[] paths, string krsPath, string ceidgPath, string risPath, string aypNip)
    {
        try
        {
            var p = PathHelpers.GetRightPath(paths);
            if (!string.IsNullOrEmpty(p?.Error?.Message))
            {
                //throw new Exception(p.Error.Message);
            }

            ExecuteDragAndDrop(_chromeDriver, _webDriverWait, "kontrolkaZalacznikow1", p.Value[PathHelpers.DE_MINIMIS].First());
            ExecuteDragAndDrop(_chromeDriver, _webDriverWait, "kontrolkaZalacznikow2", p.Value[PathHelpers.OŚWIADCZENIE_DE_MINIMIS].First());

            if (p.Value.ContainsKey(PathHelpers.PEŁNOMOCNICTWO))
            {
                try
                {
                    _chromeDriver.ExecuteScript("document.querySelectorAll('label.checkbox input[data-ng-model=\"zalacznik.dolaczony\"]')[5].click()");
                }
                catch
                {

                }
                ExecuteDragAndDropPELNOMOCNICTWO(_chromeDriver, _webDriverWait, "kontrolkaZalacznikow6", p.Value[PathHelpers.PEŁNOMOCNICTWO].First());
            }

            if (string.IsNullOrEmpty(krsPath))
            {
                ExecuteDragAndDrop(_chromeDriver, _webDriverWait, "kontrolkaZalacznikow3", ceidgPath);
            }
            else
            {
                ExecuteDragAndDrop(_chromeDriver, _webDriverWait, "kontrolkaZalacznikow3", krsPath);
            }

            if (p.Value[PathHelpers.PROGRAM].Count > 1 && p.Value.ContainsKey(PathHelpers.ZAKRES_EGZAMINU))
            {
                var array = p.Value[PathHelpers.PROGRAM].ToList();
                array.AddRange(p.Value[PathHelpers.ZAKRES_EGZAMINU]);
                ExecuteDragAndDrop(_chromeDriver, _webDriverWait, "kontrolkaZalacznikow4", array.ToArray());
            }
            else
            {
                ExecuteDragAndDrop(_chromeDriver, _webDriverWait, "kontrolkaZalacznikow4", p.Value[PathHelpers.PROGRAM].ToArray());
            }

            ExecuteDragAndDrop(_chromeDriver, _webDriverWait, "kontrolkaZalacznikow5", p.Value[PathHelpers.WZÓR_CERTYFIKATU].First());

            if (p.Value.ContainsKey(PathHelpers.OŚWIADCZENIE_CERTYFIKAT_KOMPLET))
            {
                _chromeDriver.ExecuteScript("document.querySelector('button[data-ng-click=\"dodajZalacznikUzytkownika()\"]').click()");
                var values = p.Value[PathHelpers.OŚWIADCZENIE_CERTYFIKAT_KOMPLET].ToList();
                values.Add(risPath);
                values.Add(aypNip);

                ExecuteAdditionalAttachment(_chromeDriver, _webDriverWait, values.ToList());
            }
            var dict = new Dictionary<int, string>();
            var count = 0;

            if (p.Value.ContainsKey(PathHelpers.NIP8))
            {
                dict.Add(count, p.Value[PathHelpers.NIP8].First());
            }
            if (p.Value.ContainsKey(PathHelpers.CIT))
            {
                dict.Add(++count, p.Value[PathHelpers.CIT].First());
            }
            if (p.Value.ContainsKey(PathHelpers.UMOWA_NAJMU_LOKALIZACJA))
            {
                dict.Add(++count, p.Value[PathHelpers.UMOWA_NAJMU_LOKALIZACJA].First());
            }

            _chromeDriver.ExecuteScript("document.querySelector('button[data-ng-click=\"dodajZalacznikUzytkownika()\"]').click()");

            ExecuteDragAndDropAdditional2(_chromeDriver, _webDriverWait, "", dict.Values.ToArray());
        }
        catch (Exception ex)
        {
            MessageBox.Show($"{ex.Message}", "Problem z dodawaniem załączników.", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}