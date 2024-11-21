using FormFiller.Constants;
using FormFiller.Helpers;
using FormFiller.Models;
using FormFiller.Models.TrainingEntity;
using FormFiller.Models.WorkerEntity;

using OfficeOpenXml;

using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

using SeleniumExtras.WaitHelpers;

using System.Data;
using System.Text.RegularExpressions;


namespace FormFiller.Services;

public interface ISeleniumService
{
    void LoginToPage();
    Task LoadWorkerForm();
    Task<string> DownloadComapnyByNipPdf(string nip);
    void ProvideWorkerInformation(ExcelWorkersRow row, string city);
    void LoadPlanowanyRealizator();
    void LoadInfomacjeDotyczaceKsztalcenia(ExcelWorksheet row, DataTable? listaSzkolen, DataTable? pracownicy);
}

public class SeleniumService : ISeleniumService, IDisposable
{
    private readonly IDriverService _driverService;
    private ChromeDriver _chromeDriver;
    private WebDriverWait _webDriverWait;

    public SeleniumService()
    {
        _driverService = DriverService.Instance;
        (_chromeDriver, _webDriverWait) = _driverService.GetDriver();
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

    public async Task LoadWorkerForm()
    {
        try
        {
            InitializeDriver();
            var startForm = By.CssSelector("div[psz-kfs-uczestnicy]");
            var uczestnicyForm = By.CssSelector("div[data-psz-kfs-uczestnicy]");

            var wait = new WebDriverWait(_chromeDriver, TimeSpan.FromMinutes(5));
            wait.Until(ExpectedConditions.ElementIsVisible(startForm));

            ExecuteKursyRadioButton();

            _chromeDriver.ExecuteScript(
                "document.querySelector('button[data-ng-click=\"zapiszForme()\"]').click()");

            wait.Until(ExpectedConditions.ElementIsVisible(uczestnicyForm));

            var addWorkerButton = By.Id("tabela-uczestnik");
            wait.Until(ExpectedConditions.ElementIsVisible(addWorkerButton));
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
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

            ExecuteInputByName("dane.miejscowosc", city);

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
                ExecuteInputByName("uczestnik.dofinansowanieKfs.przyznanaKwota",
                    PrzyznanaKwota.EnsureTwoDecimalPlaces(
                        (string)row.Data[WorkersFormKeys.KwotaOtrzymanegoDofinansowania]));
            });

            //Checkbox

            ExecuteCheckbox("uczestnik.typKsztalcenia.awansZawodowy");
            ExecuteCheckbox("uczestnik.typKsztalcenia.rozszerzenieObowiazkow");
            ExecuteCheckbox("uczestnik.typKsztalcenia.kompetencjeZawodowe");
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

            ExecuteInputByName("realizator.nazwa", "AYP Spółka z ograniczoną odpowiedzialnością");

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
            ExecuteInputByName("realizator.adres.adresKrajowy.ulica", "Chopina");
            ExecuteInputByName("realizator.adres.adresKrajowy.nrDomu", "28");
            ExecuteInputByName("realizator.adres.adresKrajowy.nrLokalu", "7");
            ExecuteInputByName("realizator.nip", "5842774905");
            ExecuteInputByName("realizator.krs", "0000741233");
            ExecuteInputByName("realizator.regon", "380820190");

            ExecuteInput("realizator.pkd");

            Wait();
            SelectAutocompleteOptionWithRetry();

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

    public void LoadInfomacjeDotyczaceKsztalcenia(ExcelWorksheet row, DataTable? listaSzkolen, DataTable? pracownicy)
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
                string pattern = @"\b[A-Z][a-z]*\b";

                var keyWord = GetKeyWord(szkolenie[TrainingFormKeys.TematSzkolenia].ToString(), pattern);

                var kodDzialania = _chromeDriver.FindElement(By.Id("kod-dzialania"));
                kodDzialania.SendKeys(keyWord);

                ExecuteRadioButton("KURS");
                var kształcenie = szkolenie[TrainingFormKeys.TematSzkolenia].ToString();

                ExecuteTextArea("dzialanie.nazwa", kształcenie);
                ExecuteDropdown("slowniki/OBSZARY_SZKOLEN", "dzialanie.tematyka", "Kompetencje");
                ExecuteCheckbox("dzialanie.kompetencje.potwierdzenie.certyfikat");

                //radio button
                _chromeDriver.ExecuteScript(
                    """
                    document.querySelector("input[data-ng-value='false'][data-ng-model='dzialanie.kompetencje.podstawaPrawna']").click()
                    """);

                //Liczba z excela
                var wartoscSzkolenia =
                    PrzyznanaKwota.EnsureTwoDecimalPlaces(szkolenie[TrainingFormKeys.WartośćSzkoleniaNaOsobe].ToString());

                var liczbaGodzin = CalculateTime.Calculate(liczbaDni.ToString());

                ExecuteInputByName("dzialanie.liczbaGodzin.ilosc", liczbaGodzin);
                ExecuteInputByName("dzialanie.cenaNetto", wartoscSzkolenia);
                ExecuteInputByName("dzialanie.cenaBrutto", wartoscSzkolenia);

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

    private void PracownikPopupComparation(ExcelWorksheet row, DataTable? pracownicy, string nazwa)
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
    public void SelectAutocompleteOptionWithRetry(int maxRetries = 5, int retryDelayMs = 500)
    {
        try
        {
            // Set the input value and trigger the input event
            _chromeDriver.ExecuteScript(
                "document.evaluate('//input[@data-sg-autocomplete=\"slowniki/PKD_2007\" and @data-ng-model=\"realizator.pkd\"]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.value = '8559B'");

            // Trigger the input event to show the suggestions
            _chromeDriver.ExecuteScript(
                "document.evaluate('//input[@data-sg-autocomplete=\"slowniki/PKD_2007\" and @data-ng-model=\"realizator.pkd\"]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.dispatchEvent(new Event('input', { bubbles: true }));");
            // Retry finding and clicking the list item up to `maxRetries` times
            for (int attempt = 0; attempt < maxRetries; attempt++)
            {
                try
                {
                    // Attempt to select the item
                    var listItemClickScript = $@"
                const listItem = Array.from(document.querySelectorAll(""ul.ui-autocomplete li"")).find(el => el.textContent.trim() === ""8559B - Pozostałe pozaszkolne formy edukacji, gdzie indziej niesklasyfikowane"");
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

                // Wait a short period before retrying
                Thread.Sleep(retryDelayMs);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }

    private void ExecuteInputByName(string name, string inputValue)
    {
        try
        {
            _chromeDriver.ExecuteScript(
                $"""
            document.querySelector("input[data-ng-model='{name}']").value = '{inputValue}'
            """
            );

            _chromeDriver.ExecuteScript(
                $$"""
            document.querySelector("input[data-ng-model='{{name}}']").dispatchEvent(new Event('input', { bubbles: true }))
            """
            );
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }

    private void ExecuteRadioButton(string radioButtonValue)
    {
        try
        {
            _chromeDriver.ExecuteScript(
                $"""
            const xpath = "//div[@class='wiersz']/label/input[@name='dzialanieTyp' and @value='{radioButtonValue}']";
            document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click();
            """);
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }

    void ExecuteKursyRadioButton()
    {
        try
        {
            _chromeDriver.ExecuteScript(
                $"""
            const firstXpath = "//div[@class='wiersz']/label/input[@name='potrzebySzkoleniowe' and @value='false']";
            document.evaluate(firstXpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click();
            """);
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }

    private void ExecuteCheckbox(string model)
    {
        try
        {
            _chromeDriver.ExecuteScript(
                $"""
            document.querySelector("input[data-ng-model='{model}'").click()
            """);
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }
    private void ExecuteTextArea(string model, string textArea)
    {
        try
        {
            var sanitized = JavaScriptHelper.EscapeForJavaScript(textArea);
            _chromeDriver.ExecuteScript(
                $"""
                 document.evaluate('//textarea[contains(@data-ng-model, "{model}")]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.value = '{sanitized}'
                 """);
            _chromeDriver.ExecuteScript(
                $$"""
                  document.evaluate('//textarea[contains(@data-ng-model, "{{model}}")]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.dispatchEvent(new Event('input', { bubbles: true }))
                  """
            );
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }
    private void ExecuteInput(string model)
    {
        try
        {
            _chromeDriver.ExecuteScript(
                $"""
             document.evaluate('//div[contains(@class, "sg-autocomplete") and .//input[contains(@data-ng-model, "{model}")]]//span', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click()
             """);
            _chromeDriver.ExecuteScript(
                $$"""
              document.evaluate('//div[contains(@class, "sg-autocomplete") and .//input[contains(@data-ng-model, "{{model}}")]]//span', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.dispatchEvent(new Event('input', { bubbles: true }))
              """);
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }

    private void ExecuteDropdown(string autocomplete, string model, string inputValue)
    {
        try
        {
            WebDriverWait wait = new WebDriverWait(_chromeDriver, TimeSpan.FromSeconds(10));

            // Wait for the input element to be visible
            IWebElement inputElement = wait.Until(driver =>
                driver.FindElement(By.CssSelector($"[data-sg-autocomplete='{autocomplete}']")));

            // Set value in the input field
            _chromeDriver.ExecuteScript($$"""
                                             document.querySelector('[data-sg-autocomplete="{{autocomplete}}"]').value = '{{inputValue}}';
                                         """);
            wait.Until(d =>
                d.FindElement(
                    By.CssSelector("div.wiersz-szeroka-etykieta [data-ng-click=\"toggleSuggestionsClick($event)\"]")));

            _chromeDriver.ExecuteScript(
                """
                  document.querySelector('div.wiersz-szeroka-etykieta [data-ng-click="toggleSuggestionsClick($event)"]').click()
                  """
                );
            Task.Delay(500).Wait();
            // Click the desired dropdown option
            _chromeDriver.ExecuteScript("Array.from(document.querySelectorAll(\"ul.ui-autocomplete li\")).find(el => el.textContent.trim() == \"Kompetencje cyfrowe\").click()");
        }
        catch (WebDriverTimeoutException)
        {
            Console.WriteLine("Timeout: Element or dropdown options did not load in time.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred in ExecuteDropdown: {ex.Message}");
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

    static string GetKeyWord(string sentence, string pattern)
    {
        var regex = new Regex(pattern);
        var matches = regex.Matches(sentence);

        if (matches.Count > 0)
        {
            // Return the first capitalized word
            return matches[0].Value;
        }
        else
        {
            return "Do wstawienia";
        }
    }

    static Dictionary<string, string> MapColumnByKey(DataTable dataTable, string keyColumn, string valueColumn)
    {
        // Check if columns exist
        if (!dataTable.Columns.Contains(keyColumn))
            throw new ArgumentException($"Key column '{keyColumn}' does not exist.");
        if (!dataTable.Columns.Contains(valueColumn))
            throw new ArgumentException($"Value column '{valueColumn}' does not exist.");

        // Create the dictionary
        var mapping = new Dictionary<string, string>();

        foreach (DataRow row in dataTable.Rows)
        {
            var key = row[keyColumn]?.ToString();
            var value = row[valueColumn]?.ToString();

            if (!string.IsNullOrEmpty(key) && !mapping.ContainsKey(key))
            {
                mapping[key] = value ?? string.Empty; // Use empty string if value is null
            }
        }

        return mapping;
    }

    public void Dispose()
    {
        _chromeDriver?.Quit();
        _chromeDriver?.Dispose();
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

        string downloadFolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads";

        if (Directory.Exists(downloadFolder))
        {
            // Get all files in the download folder
            var files = Directory.GetFiles(downloadFolder);

            if (files.Length > 0)
            {
                // Get the most recently downloaded file (last modified or created)
                var latestFile = files
                    .Select(file => new FileInfo(file))  // Convert file paths to FileInfo objects
                    .OrderByDescending(fileInfo => fileInfo.LastWriteTime)  // Sort by last write time (modification time)
                    .First();  // Take the most recent file

                string newFileName = Path.Combine(latestFile.DirectoryName, $"NIP_{nip}" + latestFile.Extension);
                File.Move(latestFile.FullName, newFileName);
                return newFileName;
            }
            else
            {
                Console.WriteLine("No files found in the download folder.");
            }
        }
        return string.Empty;
    }
}