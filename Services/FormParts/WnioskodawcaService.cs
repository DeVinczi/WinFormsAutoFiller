using FormFiller.Services;

using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

using WinFormsAutoFiller.Models;
using WinFormsAutoFiller.Models.PytaniaDoWnioskuEntity;

using static WinFormsAutoFiller.Helpers.SeleniumHelpers;

namespace WinFormsAutoFiller.Services.FormParts
{
    public interface IWnioskodawcaService
    {
        void DaneIdentyfikacyjne_1_1(string krs);
        void AdresSiedziby_1_2();
        Task MiejsceProwadzeniaDziałalności_1_3(string businessAddresses, string city);
        void AdresDoKorespondencji_1_4();
        void OznaczeniePKD_1_5(string pkd);
        void NumerRachunkuBankowego_1_6(string nrRachunku);
        void LiczbaZatrudnionychPracowikow_1_7(int liczbaZatrudnionych, bool czyJestMikro);
        void OsobaUprawnionaDoKontaktuZUrzedem_1_8(ContactPerson contactPerson);
    }

    internal class WnioskodawcaService : IWnioskodawcaService
    {
        private readonly IDriverService _driverService;
        private ChromeDriver _chromeDriver;
        private WebDriverWait _webDriverWait;
        private bool _testMode;

        public WnioskodawcaService(bool testMode = false)
        {
            _driverService = FormFiller.Services.DriverService.Instance;
            (_chromeDriver, _webDriverWait) = _driverService.GetDriver();
            _testMode = testMode;
        }

        public void DaneIdentyfikacyjne_1_1(string krs)
        {
            try
            {
                if (_testMode)
                {
                    _chromeDriver.ExecuteScript("var el = document.querySelector('[data-ng-if]');var scope = angular.element(el).scope();scope.dane.podmiot.typPodmiotu = 'PODMIOT_GOSPODARCZY';scope.$apply();");
                }

                if (string.IsNullOrEmpty(krs))
                {
                    ExecuteDropdownForDaneIdentyfikacyjne(_chromeDriver, _webDriverWait, false, default);
                    return;
                }

                ExecuteDropdownForDaneIdentyfikacyjne(_chromeDriver, _webDriverWait, true, krs);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Uwaga", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public void AdresSiedziby_1_2() { }

        public async Task MiejsceProwadzeniaDziałalności_1_3(string businessAddresses, string city)
        {
            try
            {
                var addressModel = AddressHelpers.GetAddresses(businessAddresses);

                //Jak czesto jest to samo miasto, a rózne miejsce prowadzenia dzialalności 

                if (_testMode)
                {
                    _chromeDriver.ExecuteScript("document.querySelector('label input#miejscowoscSelectId4sgSelectInputId').value = 'Gdańsk'");
                }

                var text = _chromeDriver.FindElement(By.CssSelector("label input#miejscowoscSelectId4sgSelectInputId"));
                var textValue = text.GetAttribute("value");

                if (!textValue.Trim().Equals(city.Trim(), StringComparison.CurrentCultureIgnoreCase))
                {
                    ExecuteCheckbox(_chromeDriver, _webDriverWait, "dane.podmiot.miejsceDzialalnosciInne");

                    var addressMatched = await RetryAndGetAddressMatch(_chromeDriver, _webDriverWait, addressModel);

                    var streetService = new StreetMapService();
                    var details = await streetService.GetCityDetailsByName(addressMatched.City, addressMatched.PostalCode);
                    await Task.Delay(1000);
                    await ExecuteDropdownWithSelect_1_3(_chromeDriver, _webDriverWait, details);

                    ExecuteInputByName(_chromeDriver, _webDriverWait, "dane.podmiot.miejsceDzialalnosci.ulica", addressMatched.Street);
                    ExecuteInputByName(_chromeDriver, _webDriverWait, "dane.podmiot.miejsceDzialalnosci.nrDomu", addressMatched.HouseNumber);
                    if (!string.IsNullOrWhiteSpace(addressMatched.FlatNumber))
                    {
                        ExecuteInputByName(_chromeDriver, _webDriverWait, "dane.podmiot.miejsceDzialalnosci.nrLokalu", addressMatched.FlatNumber);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Uwaga", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public void AdresDoKorespondencji_1_4() { }

        public void OznaczeniePKD_1_5(string pkd)
        {
            SelectAutocompleteOptionWithRetryPKD(_chromeDriver, _webDriverWait, pkd);
        }
        public void NumerRachunkuBankowego_1_6(string nrRachunku)
        {
            ExecuteInputByName(_chromeDriver, _webDriverWait, "dane.podmiot.nrKonta", nrRachunku);
        }
        public void LiczbaZatrudnionychPracowikow_1_7(int liczbaZatrudnionych, bool czyJestMikro)
        {
            ExecuteInputByName(_chromeDriver, _webDriverWait, "dane.podmiot.zatrudnieni.liczba", liczbaZatrudnionych.ToString());
            if (czyJestMikro)
            {
                ExecuteCheckbox(_chromeDriver, _webDriverWait, "dane.podmiot.zatrudnieni.mikro");
            }
        }

        public void OsobaUprawnionaDoKontaktuZUrzedem_1_8(ContactPerson contactPerson)
        {
            ExecuteInputByName(_chromeDriver, _webDriverWait, "dane.podmiot.osobaKontakt.imie", contactPerson.Name.Split(' ')[0]);
            ExecuteInputByName(_chromeDriver, _webDriverWait, "dane.podmiot.osobaKontakt.nazwisko", contactPerson.Name.Split(' ')[1]);
            ExecuteInputByName(_chromeDriver, _webDriverWait, "dane.podmiot.osobaKontakt.stanowisko", contactPerson.Position);
            ExecuteInputByName(_chromeDriver, _webDriverWait, "dane.podmiot.osobaKontakt.telefon", contactPerson.Phone);
            ExecuteInputByName(_chromeDriver, _webDriverWait, "dane.podmiot.osobaKontakt.adresEMail", contactPerson.Email);
        }
    }
}
