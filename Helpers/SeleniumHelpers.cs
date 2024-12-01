using FormFiller.Helpers;

using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

using WinFormsAutoFiller.Services;

namespace WinFormsAutoFiller.Helpers;

internal static class SeleniumHelpers
{
    public static void ExecuteDragAndDrop(ChromeDriver chromeDriver, WebDriverWait webDriverWait, string dropzoneId, string path)
    {
        try
        {
            var dropzoneElement = webDriverWait.Until(d => d.FindElement(By.CssSelector($"div#{dropzoneId}")));
            var name = path[(path.LastIndexOf('\\') + 1)..];
            var bytes = File.ReadAllBytes(path);
            var base64String = Convert.ToBase64String(bytes);

            chromeDriver.ExecuteScript(
                $$"""
                var {{dropzoneId}} = document.querySelector('div#{{dropzoneId}} div#dropzone');
                if ({{dropzoneId}}) {
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
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
            MessageBox.Show(
                $"{ex.Message}. {ex.StackTrace}",
                "Uwaga",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );

            throw;
        }
    }

    public static void ExecuteDragAndDrop(ChromeDriver chromeDriver, WebDriverWait webDriverWait, string dropzoneId, string[] paths)
    {
        try
        {
            var dropzoneElement = webDriverWait.Until(d => d.FindElement(By.CssSelector($"div#{dropzoneId}")));

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

            chromeDriver.ExecuteScript(
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
        catch (Exception ex)
        {
            MessageBox.Show(
                $"{ex.Message}",
                "Uwaga",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );

            throw;
        }
    }

    public static void ExecuteDragAndDropAdditional(ChromeDriver chromeDriver, WebDriverWait webDriverWait, string dropzoneId, string[] paths)
    {
        try
        {
            var id = chromeDriver.ExecuteScript("return document.querySelectorAll('div[data-ng-model=\"zalacznik.zalaczniki\"]')[6].id");
            chromeDriver.ExecuteScript("document.querySelector('textarea[data-ng-model=\"zalacznik.opis\"]').value = 'KRS AYP + RIS + CERTYFIKATY'");
            chromeDriver.ExecuteScript("document.querySelector('textarea[data-ng-model=\"zalacznik.opis\"]').dispatchEvent(new Event('input'), {bubbles:true})");
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

            chromeDriver.ExecuteScript(
                $$"""
          var dropzone = document.querySelector('div#{{id}} div#dropzone');
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
        catch (Exception ex)
        {
            MessageBox.Show(
                $"{ex.Message}. {ex.StackTrace}",
                "Uwaga",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
            throw;
        }
    }

    public static void ExecuteDragAndDropAdditional2(ChromeDriver chromeDriver, WebDriverWait webDriverWait, string dropzoneId, string[] paths)
    {
        try
        {
            var names = string.Join(" + ", paths.Select(x => Path.GetFileNameWithoutExtension(x)));
            var id = chromeDriver.ExecuteScript("return document.querySelectorAll('div[data-ng-model=\"zalacznik.zalaczniki\"]')[7].id");
            chromeDriver.ExecuteScript($"document.querySelectorAll('textarea[data-ng-model=\"zalacznik.opis\"]')[1].value = '{names}'");
            chromeDriver.ExecuteScript("document.querySelectorAll('textarea[data-ng-model=\"zalacznik.opis\"]')[1].dispatchEvent(new Event('input', {bubbles:true}))");
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

            chromeDriver.ExecuteScript(
                $$"""
            var dropzone = document.querySelector('div#{{id}} div#dropzone');
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
        catch (Exception ex)
        {
            MessageBox.Show(
                $"{ex.Message}. {ex.StackTrace}",
                "Uwaga",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
            throw;
        }
    }

    public static void ExecuteAdditionalAttachment(ChromeDriver chromeDriver, WebDriverWait webDriverWait, List<string> paths)
    {
        string pathNames = string.Join(" + ", paths.Select(x => Path.GetFileName(x)));

        chromeDriver.ExecuteScript($"document.querySelectorAll('div[data-ng-model=\"zalacznik.zalaczniki\"]')[6].parentNode.parentNode.querySelector(\"textarea[data-ng-model='zalacznik.opis']\").value = '{pathNames}'");
        ExecuteDragAndDropAdditional(chromeDriver, webDriverWait, "", paths.ToArray());
    }

    public static void ExecuteDragAndDropPELNOMOCNICTWO(ChromeDriver chromeDriver, WebDriverWait webDriverWait, string dropzoneId, string path)
    {
        try
        {
            var dropzoneElement = chromeDriver.ExecuteScript("return document.querySelectorAll('label.checkbox input[data-ng-model=\"zalacznik.dolaczony\"]')[5].parentNode.parentNode.parentNode.querySelector('div[data-ng-model=\"zalacznik.zalaczniki\"]').id");
            var name = path[(path.LastIndexOf('\\') + 1)..];
            var bytes = File.ReadAllBytes(path);
            var base64String = Convert.ToBase64String(bytes);

            chromeDriver.ExecuteScript(
                $$"""
              var dropzone = document.querySelector('div#{{dropzoneElement}} div#dropzone');
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
                  dropzone.dispatchEvent(event);
              } else {
                  throw new Error('Dropzone element not found');
              }
              """);
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
            MessageBox.Show(
                $"{ex.Message}. {ex.StackTrace}",
                "Uwaga",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
        }
    }
    public static void SelectAutocompleteOptionWithRetryPKD(ChromeDriver chromeDriver, WebDriverWait webDriverWait, string val, int maxRetries = 5, int retryDelayMs = 500)
    {
        try
        {
            chromeDriver.ExecuteScript(
                $"document.querySelector('input[data-ng-model=\"dane.podmiot.pkd\"]').value = '{val}'");

            for (int attempt = 0; attempt < maxRetries; attempt++)
            {
                try
                {
                    chromeDriver.ExecuteScript("document.querySelector('input[data-ng-model=\"dane.podmiot.pkd\"]').dispatchEvent(new Event('input', { bubbles: true }))");

                    chromeDriver.ExecuteScript("document.querySelector('input[data-ng-model=\"dane.podmiot.pkd\"]').closest('div.sg-autocomplete')?.querySelector('span.sg-autocomplete-ikona').click()");

                    var listItemClickScript = $@"
                    var listItem2 = Array.from(document.querySelectorAll(""ul.ui-autocomplete li"")).find(el => el.textContent.trim().startsWith(""{val}""));
                    if (listItem2) {{
                        listItem2.click();
                        return true;
                    }}
                    return false;
            ";

                    bool clicked = (bool)chromeDriver.ExecuteScript(listItemClickScript);

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

                Task.Delay(400).Wait();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
            MessageBox.Show(
                $"{ex.Message}. {ex.StackTrace}",
                "Uwaga",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
        }
    }
    public static void SelectAutocompleteOptionWithRetryW(ChromeDriver chromeDriver, WebDriverWait webDriverWait, string val, int maxRetries = 5, int retryDelayMs = 500)
    {
        try
        {
            chromeDriver.ExecuteScript(
                $"document.querySelector('input[data-ng-model=\"realizator.pkd\"]').value = '{val}'");

            for (int attempt = 0; attempt < maxRetries; attempt++)
            {
                try
                {
                    chromeDriver.ExecuteScript("document.querySelector('input[data-ng-model=\"realizator.pkd\"]').dispatchEvent(new Event('input', { bubbles: true }))");

                    chromeDriver.ExecuteScript("document.querySelector('input[data-ng-model=\"realizator.pkd\"]').closest('div.sg-autocomplete')?.querySelector('span.sg-autocomplete-ikona').click()");

                    var listItemClickScript = $@"
                    var listItem = Array.from(document.querySelectorAll(""ul.ui-autocomplete li"")).findLast(el => el.textContent.trim().startsWith(""{val}""));
                    if (listItem) {{
                        listItem.click();
                        return true;
                    }}
                    return false;
            ";

                    bool clicked = (bool)chromeDriver.ExecuteScript(listItemClickScript);

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

                Task.Delay(400).Wait();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
            MessageBox.Show(
                $"{ex.Message}. {ex.StackTrace}",
                "Uwaga",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
        }
    }
    public static void ExecuteInputByName(ChromeDriver chromeDriver, WebDriverWait webDriverWait, string name, string inputValue)
    {
        try
        {
            chromeDriver.ExecuteScript(
                $"""
            document.querySelector("input[data-ng-model='{name}']").value = '{inputValue}'
            """
            );

            chromeDriver.ExecuteScript(
                $$"""
            document.querySelector("input[data-ng-model='{{name}}']").dispatchEvent(new Event('input', { bubbles: true }))
            """
            );
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
            MessageBox.Show(
                $"{ex.Message}. {ex.StackTrace}",
                "Uwaga",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
        }
    }
    public static void ExecuteInputKodPocztowy(ChromeDriver chromeDriver, string inputValue)
    {
        try
        {
            chromeDriver.ExecuteScript(
                $"""
            document.querySelector("input#kodPocztowyInputId1").value = '{inputValue}'
            """
            );

            chromeDriver.ExecuteScript(
                $$"""
            document.querySelector("input#kodPocztowyInputId1").dispatchEvent(new Event('input', { bubbles: true }))
            """
            );
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
            MessageBox.Show(
                $"{ex.Message}. {ex.StackTrace}",
                "Uwaga",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
        }
    }

    public static void ExecuteRadioButton(ChromeDriver chromeDriver, WebDriverWait webDriverWait, string radioButtonValue)
    {
        try
        {
            chromeDriver.ExecuteScript(
                $"""
            var xpath = "//div[@class='wiersz']/label/input[@name='dzialanieTyp' and @value='{radioButtonValue}']";
            document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click();
            """);
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
            MessageBox.Show(
                $"{ex.Message}. {ex.StackTrace}",
                "Uwaga",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
        }
    }

    public static void ExecuteKursyRadioButton(ChromeDriver chromeDriver, WebDriverWait webDriverWait)
    {
        try
        {
            chromeDriver.ExecuteScript(
                $"""
            var firstXpath = "//div[@class='wiersz']/label/input[@name='potrzebySzkoleniowe' and @value='false']";
            document.evaluate(firstXpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click();
            """);
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
            MessageBox.Show(
                $"{ex.Message}. {ex.StackTrace}",
                "Uwaga",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
        }
    }

    public static void ExecuteCheckbox(ChromeDriver chromeDriver, WebDriverWait webDriverWait, string model)
    {
        try
        {
            chromeDriver.ExecuteScript(
                $"""
            document.querySelector("input[data-ng-model='{model}'").click()
            """);
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
            MessageBox.Show(
                $"{ex.Message}. {ex.StackTrace}",
                "Uwaga",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
        }
    }
    public static void ExecuteTextArea(ChromeDriver chromeDriver, WebDriverWait webDriverWait, string model, string textArea)
    {
        try
        {
            var sanitized = JavaScriptHelper.EscapeForJavaScript(textArea);
            chromeDriver.ExecuteScript(
                $"""
                 document.evaluate('//textarea[contains(@data-ng-model, "{model}")]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.value = '{sanitized}'
                 """);
            chromeDriver.ExecuteScript(
                $$"""
                  document.evaluate('//textarea[contains(@data-ng-model, "{{model}}")]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.dispatchEvent(new Event('input', { bubbles: true }))
                  """
            );
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
            MessageBox.Show(
                $"{ex.Message}. {ex.StackTrace}",
                "Uwaga",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
        }
    }
    public static void ExecuteInput(ChromeDriver chromeDriver, WebDriverWait webDriverWait, string model)
    {
        try
        {
            chromeDriver.ExecuteScript(
                $"""
             document.evaluate('//div[contains(@class, "sg-autocomplete") and .//input[contains(@data-ng-model, "{model}")]]//span', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click()
             """);
            chromeDriver.ExecuteScript(
                $$"""
              document.evaluate('//div[contains(@class, "sg-autocomplete") and .//input[contains(@data-ng-model, "{{model}}")]]//span', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.dispatchEvent(new Event('input', { bubbles: true }))
              """);
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
            MessageBox.Show(
                $"{ex.Message}. {ex.StackTrace}",
                "Uwaga",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
        }
    }

    public static void ExecuteDropdown(ChromeDriver chromeDriver, WebDriverWait webDriverWait, string autocomplete, string model, string inputValue)
    {
        try
        {
            WebDriverWait wait = new WebDriverWait(chromeDriver, TimeSpan.FromSeconds(10));

            IWebElement inputElement = wait.Until(driver =>
                driver.FindElement(By.CssSelector($"[data-sg-autocomplete='{autocomplete}']")));

            chromeDriver.ExecuteScript($$"""
                                             document.querySelector('[data-sg-autocomplete="{{autocomplete}}"]').value = '{{inputValue}}';
                                         """);
            wait.Until(d =>
                d.FindElement(
                    By.CssSelector("div.wiersz-szeroka-etykieta [data-ng-click=\"toggleSuggestionsClick($event)\"]")));

            chromeDriver.ExecuteScript(
                """
                  document.querySelector('div.wiersz-szeroka-etykieta [data-ng-click="toggleSuggestionsClick($event)"]').click()
                  """
                );
            Task.Delay(500).Wait();

            chromeDriver.ExecuteScript("Array.from(document.querySelectorAll(\"ul.ui-autocomplete li\")).find(el => el.textContent.trim() == \"Kompetencje cyfrowe\").click()");
        }
        catch (WebDriverTimeoutException e)
        {
            Console.WriteLine("Timeout: Element or dropdown options did not load in time.");
            MessageBox.Show(
                $"{e.Message}",
                "Uwaga",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred in ExecuteDropdown: {ex.Message}");
            MessageBox.Show(
                $"{ex.Message}. {ex.StackTrace}",
                "Uwaga",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
        }
    }

    public static void ExecuteDropdownForDaneIdentyfikacyjne(ChromeDriver chromeDriver, WebDriverWait webDriverWait, bool hasKrs, string krs)
    {
        try
        {
            const string KRS = "KRS";
            const string CEIDG = "CEIDG";

            var choice = hasKrs switch
            {
                true => KRS,
                false => CEIDG,
            };

            chromeDriver.ExecuteScript("document.querySelector('div[data-ng-if] span[class=\"sg-autocomplete-ikona\"]').click()");
            var isDropdownShown = (bool)chromeDriver.ExecuteScript($"return Array.from(document.querySelectorAll(\"ul.ui-autocomplete li\")).some(el => el.textContent.trim().includes(\"{choice}\"));");
            webDriverWait.Until(x => isDropdownShown);
            chromeDriver.ExecuteScript($"Array.from(document.querySelectorAll(\"ul.ui-autocomplete li\")).find(el => el.textContent.trim().includes(\"{choice}\")).click();");

            if (hasKrs)
            {
                ExecuteInputByName(chromeDriver, webDriverWait, "dane.podmiot.organizacja.rejestr.numerKrs", krs);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"{ex.Message}. {ex.StackTrace}",
                "Uwaga",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
            throw;
        }
    }

    public async static Task<Models.Address> RetryAndGetAddressMatch(
        ChromeDriver chromeDriver,
        WebDriverWait webDriverWait,
        List<Models.Address> addresses)
    {
        try
        {
            var maxRetries = 3; // Retry each address up to 3 times
            var wait = new WebDriverWait(chromeDriver, TimeSpan.FromSeconds(2)); // Increase timeout to 5 seconds

            foreach (var address in addresses)
            {
                for (int retry = 0; retry < maxRetries; retry++)
                {
                    try
                    {
                        ExecuteInputKodPocztowy(chromeDriver, address.PostalCode);

                        // Delay to allow the DOM to update
                        await Task.Delay(500);

                        var isInvalid = await Task.Run(() =>
                        {
                            try
                            {
                                return wait.Until(driver =>
                                {
                                    var element = driver.FindElement(By.CssSelector("div[data-ng-message=\"obslugiwanePowiatyValidate\"]"));
                                    return element.Displayed && !string.IsNullOrWhiteSpace(element.Text);
                                });
                            }
                            catch (WebDriverTimeoutException)
                            {
                                return false;
                            }
                        });

                        if (!isInvalid)
                        {
                            return address; // Found a valid address
                        }
                    }
                    catch (Exception innerEx)
                    {
                        // Log or handle retries
                        Console.WriteLine($"Retry {retry + 1} for postal code {address.PostalCode} failed: {innerEx.Message}");
                        await Task.Delay(250 * (retry + 1)); // Incremental delay for retries
                    }
                }
            }

            throw new Exception("Nie ma kodu pocztowego pasujacego do danego PUP");
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"{ex.Message}. {ex.StackTrace}",
                "Uwaga",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);

            throw;
        }
    }

    public static async Task ExecuteDropdownWithSelect_1_3(ChromeDriver chromeDriver, WebDriverWait webDriverWait, StreetMapApiModel model)
    {
        try
        {
            if (model.CzyWieś == true)
            {
                var gmina = model.Gmina.Trim()[5..].Trim();
                var powiat = model.Powiat.Trim()[6..].Trim();
                var województwo = model.Województwo.Trim()[11..].Trim();

                chromeDriver.ExecuteScript("document.querySelector('input#powiatSelectId6sgSelectInputId').click()");
                var isDropdownShown = (bool)chromeDriver.ExecuteScript($"return Array.from(document.querySelectorAll(\"ul.ui-autocomplete li\")).some(el => el.textContent.trim().includes(\"{powiat}\"));");
                webDriverWait.Until(x => isDropdownShown);
                chromeDriver.ExecuteScript($"Array.from(document.querySelectorAll(\"ul.ui-autocomplete li\")).find(el => el.textContent.trim().includes(\"{powiat}\")).click();");

                chromeDriver.ExecuteScript("document.querySelector('input#gminaSelectId7sgSelectInputId').click()");
                var isDropdownShown2 = (bool)chromeDriver.ExecuteScript($"return Array.from(document.querySelectorAll(\"ul.ui-autocomplete li\")).some(el => el.textContent.trim().startsWith(\"{gmina} (gmina wiejska)\"));");
                webDriverWait.Until(x => isDropdownShown2);
                chromeDriver.ExecuteScript($"Array.from(document.querySelectorAll(\"ul.ui-autocomplete li\")).find(el => el.textContent.trim().includes(\"{gmina} (gmina wiejska)\")).click();");

                await Task.Delay(500);
                chromeDriver.ExecuteScript("document.querySelector('input#miejscowoscSelectId8sgSelectInputId').click()");
                var isDropdownShown3 = (bool)chromeDriver.ExecuteScript($"return Array.from(document.querySelectorAll(\"ul.ui-autocomplete li\")).some(el => el.textContent.trim().startsWith(\"{model.Miasto}\"));");
                webDriverWait.Until(x => isDropdownShown3);
                chromeDriver.ExecuteScript($"Array.from(document.querySelectorAll(\"ul.ui-autocomplete li\")).find(el => el.textContent.trim() === \"{model.Miasto}\").click()");

                return;
            }

        }
        catch (Exception ex)
        {
            MessageBox.Show(
               $"{ex.Message}. {ex.StackTrace}",
               "Uwaga",
               MessageBoxButtons.OK,
               MessageBoxIcon.Warning);

            throw;
        }
    }
}
