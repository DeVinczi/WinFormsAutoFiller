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

using WinFormsAutoFiller.Helpers;
using WinFormsAutoFiller.Models.PytaniaDoWnioskuEntity;
using WinFormsAutoFiller.Utilis;


namespace FormFiller.Services;

public interface ISeleniumService
{
    void LoginToPage();
    Task LoadForm(string businessAddresses, string city, string pkd, string nrRachunku, int liczbaZatrudnionych, ContactPerson contactPerson, DateTime startDate, DateTime endDate, string[] paths, string krs, string ceidg);
    Task<string> DownloadComapnyByNipPdf(string nip);
    void ProvideWorkerInformation(ExcelWorkersRow row, string city);
    void LoadPlanowanyRealizator();
    void LoadInfomacjeDotyczaceKsztalcenia(ExcelWorksheet row, DataTable? listaSzkolen, DataTable? pracownicy);
    void LoadEndOfForm();
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

    public async Task LoadForm(string businessAddresses, string city, string pkd, string nrRachunku, int liczbaZatrudnionych, ContactPerson contactPerson, DateTime startDate, DateTime endDate, string[] paths, string krs, string ceidg)
    {
        try
        {
            InitializeDriver();
            var startForm = By.CssSelector("div.formularz fieldset");
            var uczestnicyForm = By.CssSelector("div[data-psz-kfs-uczestnicy]");

            var wait = new WebDriverWait(_chromeDriver, TimeSpan.FromMinutes(5));
            wait.Until(ExpectedConditions.ElementIsVisible(startForm));

            ExecuteKursyRadioButton();

            _chromeDriver.ExecuteScript(
                "document.querySelector('button[data-ng-click=\"zapiszForme()\"]').click()");

            wait.Until(ExpectedConditions.ElementIsVisible(uczestnicyForm));

            var element = _chromeDriver.FindElement(By.CssSelector("label.nazwaUrzedu-gf"));
            element.SendKeys(city);

            var address = FuzzHelpers.GetAddress(businessAddresses, city);
            //wyciagnac kod pocztowy
            var kodPocztowy = _chromeDriver.ExecuteAsyncScript($"document.querySelector('label.kod-pocztowy-gf input').value = '{address}'");
            _chromeDriver.ExecuteAsyncScript("document.querySelector('label.kod-pocztowy-gf input').dispatchEvent(new Event('input', {bubbles: true })) ");

            ExecuteInputByName("dane.podmiot.adresSiedziby.ulica", address);
            ExecuteInputByName("dane.podmiot.adresSiedziby.nrDomu", address);
            if (address is not null)
            {
                ExecuteInputByName("dane.podmiot.adresSiedziby.nrLokalu", address);
            }

            //Wyciagnąć miasto
            var text = _chromeDriver.ExecuteAsyncScript("document.querySelector('label input#miejscowoscSelectId4sgSelectInputId').value").ToString().Trim();
            if (text != city)
            {
                ExecuteCheckbox("miejsceDzialalnosciInne");

                var kodPocztowyPowiazanyZPUP = _chromeDriver.ExecuteAsyncScript($"document.querySelector('label.kod-pocztowy-gf input#kodPocztowyInputId1').value = '{address}'");
                _chromeDriver.ExecuteAsyncScript("document.querySelector('label.kod-pocztowy-gf input#kodPocztowyInputId1').dispatchEvent(new Event('input', {bubbles: true })) ");

                ExecuteInputByName("dane.podmiot.miejsceDzialalnosci.ulica", address);
                ExecuteInputByName("dane.podmiot.miejsceDzialalnosci.nrDomu", address);
                if (address is not null)
                {
                    ExecuteInputByName("dane.podmiot.miejsceDzialalnosci.nrLokalu", address);
                }
            }

            SelectAutocompleteOptionWithRetryPKD(pkd);
            ExecuteInputByName("podmiotNrKonta", nrRachunku);
            ExecuteInputByName("dane.podmiot.zatrudnieni.liczba", liczbaZatrudnionych.ToString());
            if (liczbaZatrudnionych > 10)
            {
                ExecuteCheckbox("dane.podmiot.zatrudnieni.mikro");
            }

            ExecuteInputByName("dane.podmiot.osobaKontakt.imie", contactPerson.Name.Split(' ')[0]);
            ExecuteInputByName("dane.podmiot.osobaKontakt.nazwisko", contactPerson.Name.Split(' ')[1]);
            ExecuteInputByName("dane.podmiot.osobaKontakt.stanowisko", contactPerson.Position);
            ExecuteInputByName("dane.podmiot.osobaKontakt.telefon", contactPerson.Phone);
            ExecuteInputByName("dane.podmiot.osobaKontakt.adresEMail", contactPerson.Email);

            _chromeDriver.ExecuteAsyncScript("document.querySelector(\"input[data-ng-model='dane.koszty.zrodloSrodkow']\").click()");

            ExecuteInputByName("dane.termin.dataOd", startDate.ToString());
            ExecuteInputByName("dane.termin.dataDo", endDate.ToString());

            var p = PathHelpers.GetRightPath(paths);
            if (!string.IsNullOrEmpty(p?.Error?.Message))
            {
                throw new Exception(p.Error.Message);
            }
            if (p.Value.ContainsKey(PathHelpers.OŚWIADCZENIE_DE_MINIMIS))
            {
                ExecuteDragAndDrop("kontrolkaZalacznikow1", p.Value[PathHelpers.OŚWIADCZENIE_DE_MINIMIS].First());
            }
            else { ExecuteDragAndDrop("kontrolkaZalacznikow1", p.Value[PathHelpers.DE_MINIMIS].First()); }

            ExecuteDragAndDrop("kontrolkaZalacznikow2", p.Value[PathHelpers.OŚWIADCZENIE_CERTYFIKAT_KOMPLET].First());

            if (string.IsNullOrEmpty(krs))
            {
                ExecuteDragAndDrop("kontrolkaZalacznikow3", ceidg);
            }
            else
            {
                ExecuteDragAndDrop("kontrolkaZalacznikow3", krs);
            }
            if (p.Value[PathHelpers.PROGRAM].Count > 1)
            {
                ExecuteDragAndDrop("kontrolkaZalacznikow4", p.Value[PathHelpers.PROGRAM].ToArray());
            }
            else
            {
                ExecuteDragAndDrop("kontrolkaZalacznikow4", p.Value[PathHelpers.PROGRAM].First());
            }

            ExecuteDragAndDrop("kontrolkaZalacznikow5", p.Value[PathHelpers.WZÓR_CERTYFIKATU].First());

            if (p.Value.ContainsKey(PathHelpers.PEŁNOMOCNICTWO))
            {
                ExecuteCheckbox("zalacznik.dolaczony");
                ExecuteDragAndDrop("kontrolkaZalacznikow6", p.Value[PathHelpers.PEŁNOMOCNICTWO].First());
            }
            var dict = new Dictionary<int, string>();
            var count = 0;

            _webDriverWait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("button[data-ng-click=\"dodajZalacznikUzytkownika()\"]")));

            if (p.Value.ContainsKey(PathHelpers.NIP8))
            {
                _chromeDriver.ExecuteScript("document.querySelector('button[data-ng-click=\"dodajZalacznikUzytkownika()\"]').click()");
                dict.Add(count, p.Value[PathHelpers.NIP8].First());
            }
            if (p.Value.ContainsKey(PathHelpers.CIT))
            {
                _chromeDriver.ExecuteScript("document.querySelector('button[data-ng-click=\"dodajZalacznikUzytkownika()\"]').click()");
                dict.Add(count, p.Value[PathHelpers.CIT].First());
            }
            if (p.Value.ContainsKey(PathHelpers.UMOWA_NAJMU_LOKALIZACJA))
            {
                _chromeDriver.ExecuteScript("document.querySelector('button[data-ng-click=\"dodajZalacznikUzytkownika()\"]').click()");
                dict.Add(count, p.Value[PathHelpers.UMOWA_NAJMU_LOKALIZACJA].First());
            }
            ExecuteAdditionalAttachment(dict.Values.ToList());

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

                var keyWord = szkolenie[TrainingFormKeys.TematSzkolenia].ToString()![..10];

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

    public void SelectAutocompleteOptionWithRetryPKD(string val, int maxRetries = 5, int retryDelayMs = 500)
    {
        try
        {
            // Set the input value and trigger the input event
            _chromeDriver.ExecuteScript(
                $"document.evaluate('//input[@data-sg-autocomplete=\"slowniki/PKD_2007\" and @data-ng-model=\"dane.podmiot.pkd\"]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.value = '{val}'");

            // Trigger the input event to show the suggestions
            _chromeDriver.ExecuteScript(
                "document.evaluate('//input[@data-sg-autocomplete=\"slowniki/PKD_2007\" and @data-ng-model=\"dane.podmiot.pkd\"]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.dispatchEvent(new Event('input', { bubbles: true }));");
            // Retry finding and clicking the list item up to `maxRetries` times
            for (int attempt = 0; attempt < maxRetries; attempt++)
            {
                try
                {
                    // Attempt to select the item
                    var listItemClickScript = $@"
                var listItem2 = Array.from(document.querySelectorAll(""ul.ui-autocomplete li"")).find(el => el.textContent.trim() === ""{val}"");
                if (listItem2) {{
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
            var xpath = "//div[@class='wiersz']/label/input[@name='dzialanieTyp' and @value='{radioButtonValue}']";
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
                mapping[key] = value ?? string.Empty;
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
        ExecuteCheckbox("dane.oswiadczenie.pkt1");
        _chromeDriver.ExecuteScript("document.querySelector(\"input[data-ng-model='dane.oswiadczenie.pkt2']\").click()");
        ExecuteCheckbox("dane.oswiadczenie.pkt3");
        ExecuteCheckbox("dane.oswiadczenie.pkt4");
        ExecuteCheckbox("dane.oswiadczenie.pkt5");
        ExecuteCheckbox("dane.oswiadczenie.pkt6");
    }

    void ExecuteDragAndDrop(string dropzoneId, string path)
    {
        var dropzoneElement = _webDriverWait.Until(d => d.FindElement(By.CssSelector($"div#{dropzoneId}")));
        var name = path[(path.LastIndexOf('\\') + 1)..];
        var bytes = File.ReadAllBytes(path);
        var base64String = Convert.ToBase64String(bytes);

        _chromeDriver.ExecuteScript(
            $$"""
              var {{dropzoneId}} = document.querySelector('div#{{dropzoneId}} div#dropzone');
              if (dropzone) {
                  var byteCharacters = atob("{{base64String}}");
                  var byteArray = new Uint8Array(byteCharacters.length);
              
                  for (var i = 0; i < byteCharacters.length; i++) {
                      byteArray[i] = byteCharacters.charCodeAt(i);
                  }
              
                  var file = new File([byteArray], "{{name}}", { type: 'application/pdf' });
                  var dataTransfer = new DataTransfer();
                  dataTransfer.items.add(file);
                  var event = new DragEvent('drop', {
                      dataTransfer: dataTransfer
                  });
                  {{dropzoneId}}.dispatchEvent(event);
              } else {
                  throw new Error('Dropzone element not found');
              }
              """);
    }

    void ExecuteDragAndDrop(string dropzoneId, string[] paths)
    {
        var dropzoneElement = _webDriverWait.Until(d => d.FindElement(By.CssSelector($"div#{dropzoneId}")));

        var filesData = paths.Select(path =>
        {
            var name = path[(path.LastIndexOf('\\') + 1)..];
            var bytes = File.ReadAllBytes(path);
            var base64String = Convert.ToBase64String(bytes);
            return new { name, base64String };
        }).ToArray();

        var jsFilesArray = filesData.Select(file =>
            $$"""
          {
              name: "{{file.name}}",
              content: "{{file.base64String}}"
          }
        """
        ).Aggregate((a, b) => $"{a},{b}");

        _chromeDriver.ExecuteScript(
            $$"""
          var dropzone = document.querySelector('div#{{dropzoneId}} div#dropzone');
          if (dropzone) {
              var files = [{{jsFilesArray}}].map(file => {
                  var byteCharacters = atob(file.content);
                  var byteArray = new Uint8Array(byteCharacters.length);

                  for (var i = 0; i < byteCharacters.length; i++) {
                      byteArray[i] = byteCharacters.charCodeAt(i);
                  }

                  return new File([byteArray], file.name, { type: 'application/pdf' });
              });

              var dataTransfer = new DataTransfer();
              files.forEach(file => dataTransfer.items.add(file));

              var event = new DragEvent('drop', {
                  dataTransfer: dataTransfer
              });
              dropzone.dispatchEvent(event);
          } else {
              throw new Error('Dropzone element not found');
          }
          """
        );
    }

    void ExecuteAdditionalAttachment(List<string> paths)
    {
        var elements = _chromeDriver.FindElements(By.CssSelector("div[data-ng-repeat='zalacznik in zalaczniki track by $index']"));
        if (elements.Count > 1)
        {
            var elementsToAdd = elements.Where(x => x.Text.StartsWith("Opis załącznika:")).ToList();
            var constInt = 7;

            for (int i = 0; i < elementsToAdd.Count; i++)
            {
                var name = Path.GetFileName(paths[i]);
                var textarea = elementsToAdd[i].FindElement(By.CssSelector("textarea[data-ng-model='zalacznik.opis']"));
                textarea.Clear();
                textarea.SendKeys($"{name}");

                ExecuteDragAndDrop($"kontrolkaZalacznikow{constInt}", paths[i]);
                ++constInt;
            }
        }
        else
        {
            Console.WriteLine("Brak Elementów w Załącznikach!");
        }
    }
}