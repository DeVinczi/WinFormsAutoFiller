using FormFiller.Constants;
using FormFiller.Models.TrainingEntity;
using FormFiller.Services;

using FuzzySharp;

using System.Data;
using System.Text.RegularExpressions;

using WinFormsAutoFiller.Encryption;
using WinFormsAutoFiller.Helpers;
using WinFormsAutoFiller.Models.KrsEntity;
using WinFormsAutoFiller.Models.PytaniaDoWnioskuEntity;
using WinFormsAutoFiller.Services;
using WinFormsAutoFiller.Utilis;

using Color = System.Drawing.Color;
using Font = System.Drawing.Font;
using Size = System.Drawing.Size;

namespace WinFormsAutoFiller
{
    public partial class Form1 : Form
    {
        private string file1Path;
        private string file2Path;
        private ProgressBar progressBar;
        private string[] selectedFiles;
        private Dictionary<string, string> filePaths;
        private Button openButton;
        private string city;
        private ComboBox comboBoxFiles;
        private Label labelUzasadnienia;
        string directoryPath = Path.Combine(Application.StartupPath, "Uzasadnienia");

        public Form1()
        {
            InitializeComponent();
            Text = "Uzupe³nianie formularza KFS";
            Size = new Size(1400, 800);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;

            labelUzasadnienia = new Label();
            labelUzasadnienia.Text = "Uzasadnienia";
            labelUzasadnienia.Font = new Font("Segoe UI", 12);// Set the text for the label
            labelUzasadnienia.Location = new System.Drawing.Point(10, 370); // Position the label
            labelUzasadnienia.Width = 300;  // Set width of the label

            comboBoxFiles = new ComboBox();
            comboBoxFiles.Location = new System.Drawing.Point(10, 400);  // Position the ComboBox below the label
            comboBoxFiles.Width = 380;  // Set width of the ComboBox

            // Load files into the ComboBox
            LoadFiles();

            Label titleLabel = new Label
            {
                Text = "Uzupe³nianie formularza KFS",
                Font = new Font("Segoe UI", 18, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Top,
                Height = 60
            };

            Label descriptionLabel = new Label
            {
                Text = "Po przyciœniêciu przycisku \"Za³aduj pliki do formularza KFS\" otworzy Ci siê Google Chrome.",
                Font = new Font("Segoe UI", 12),
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Top,
                Height = 40
            };

            Label descriptionLabel2 = new Label
            {
                Text = "Zaloguj siê do pracuj.gov.pl",
                Font = new Font("Segoe UI", 12),
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Top,
                Height = 40
            };

            Label descriptionLabel3 = new Label
            {
                Text = "Przejdz do formularza PSZ-KFS i poczekaj a¿ za³aduj¹ siê dane.",
                Font = new Font("Segoe UI", 12),
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Top,
                Height = 40
            };

            Button file1Button = CreateFileButton("Wybierz plik Excel z danymi ogólnymi", 1);
            Button file2Button = CreateFileButton("Wybierz plik Excel z wycen¹", 2);
            Button file3Button = CreateFileButton2("Za³¹czniki", 3);  // Third file button

            Label file1Label = CreateFileLabel(1);
            Label file2Label = CreateFileLabel(2);
            Label file3Label = CreateFileLabel2(3);  // Third file label

            Button uploadButton = new Button
            {
                Text = "Za³aduj pliki do formularza KFS",
                Dock = DockStyle.Bottom,
                Height = 50,
                Font = new Font("Segoe UI", 14),
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            progressBar = new ProgressBar
            {
                Dock = DockStyle.Bottom,
                Height = 20,
                Minimum = 0,
                Maximum = 100,
                Step = 1
            };


            uploadButton.FlatAppearance.BorderSize = 0;
            uploadButton.Click += UploadButton_Click;

            // Add controls to form
            Controls.Add(uploadButton);
            Controls.Add(file2Label);
            Controls.Add(file2Button);
            Controls.Add(file1Label);
            Controls.Add(file1Button);
            Controls.Add(descriptionLabel3);
            Controls.Add(descriptionLabel2);
            Controls.Add(descriptionLabel);
            Controls.Add(titleLabel);
            Controls.Add(file3Label);   // Add third file label
            Controls.Add(file3Button);  // Add third file button
            Controls.Add(progressBar);
            Controls.Add(labelUzasadnienia);
            Controls.Add(comboBoxFiles);

            // Set background color
            BackColor = Color.FromArgb(240, 240, 240);
            file3Button.AllowDrop = true;
            file3Button.DragEnter += File3Button_DragEnter;
            file3Button.DragDrop += File3Button_DragDrop;
        }
        private void LoadFiles()
        {
            // Define the path to the "Uzasadnienia" folder
            string directoryPath = Path.Combine(Application.StartupPath, "Uzasadnienia");

            // Check if the directory exists
            if (Directory.Exists(directoryPath))
            {
                // Get all the files in the directory
                string[] files = Directory.GetFiles(directoryPath);

                // Add file names to the ComboBox
                foreach (string file in files)
                {
                    // Add only the file name, not the full path
                    comboBoxFiles.Items.Add(Path.GetFileName(file));
                }
            }
            else
            {
                MessageBox.Show("The 'Uzasadnienia' directory does not exist.");
            }
        }
        private Button CreateFileButton(string text, int fileNumber)
        {
            Button button = new Button
            {
                Text = text,
                Size = new Size(400, 35),
                Location = new Point(500, 220 + (fileNumber - 1) * 80),
                Font = new Font("Segoe UI", 10),
                BackColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            button.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
            button.Click += (sender, e) => FileButton_Click(fileNumber);
            return button;
        }
        private Button CreateFileButton2(string text, int fileNumber)
        {
            Button button = new Button
            {
                Text = text,
                Size = new Size(400, 35),
                Location = new Point(500, 220 + (fileNumber - 1) * 80),
                Font = new Font("Segoe UI", 10),
                BackColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            button.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
            button.Click += (sender, e) => FileButton_Click_2(fileNumber);
            return button;
        }

        private Label CreateFileLabel(int fileNumber)
        {
            return new Label
            {
                Text = GUIMessage.BrakWybranegoPliku,
                Size = new Size(500, 20),
                Location = new Point(450, 260 + (fileNumber - 1) * 80),
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.Gray,
                TextAlign = ContentAlignment.MiddleCenter
            };
        }

        private Label CreateFileLabel2(int fileNumber)
        {
            Label label = new Label
            {
                Text = GUIMessage.BrakWybranegoPliku,
                Size = new Size(600, 200),
                Location = new Point(400, 260 + (fileNumber - 1) * 80),
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.Gray,
                TextAlign = ContentAlignment.TopCenter,
                AutoSize = false,
            };

            // Create a ToolTip for the label
            ToolTip toolTip = new ToolTip
            {
                ToolTipIcon = ToolTipIcon.Info, // You can choose other icons
                ToolTipTitle = "Pliki wybrane"  // Title for the tooltip
            };

            label.MouseHover += (sender, e) =>
            {
                // Check if the label's text exceeds its bounds
                if (label.Text.Length > 100) // Adjust this threshold as needed
                {
                    // Show the tooltip with the full text
                    toolTip.SetToolTip(label, label.Text);
                }
            };

            return label;
        }
        private void FileButton_Click_2(int fileNumber)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Title = $"Wybierz pliki: {fileNumber}",
                Multiselect = true,
            };

            // Show the dialog and process the selected files
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                selectedFiles = openFileDialog.FileNames;  // Store selected files
                filePaths = AddFilePaths(openFileDialog.FileNames);

                // Display the selected file names in the label
                ((Label)Controls[9]).Text = string.Join(Environment.NewLine, selectedFiles.Select(file => Path.GetFileName(file)));
            }
        }

        private Dictionary<string, string> AddFilePaths(string[] fileNames)
        {
            var dict = new Dictionary<string, string>();
            for (int i = 0; i < fileNames.Length; i++)
            {
                dict.TryAdd($"{i}", fileNames[i]);
            }
            return dict;
        }

        private void File3Button_DragEnter(object sender, DragEventArgs e)
        {
            // Check if the data being dragged contains file paths
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;  // Show the copy cursor
            }
        }

        private void File3Button_DragDrop(object sender, DragEventArgs e)
        {
            // Get the file paths from the dragged files
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            selectedFiles = files;  // Store the file paths

            // Display the selected file names in the label
            var z = string.Join(Environment.NewLine, selectedFiles.Select(file => Path.GetFileName(file)));
            ((Label)Controls[9]).Text = z;
        }

        private void FileButton_Click(int fileNumber)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xls;*.xlsm;*.xlsb;*.xltx;*.xlw",
                Title = $"Wybierz plik: {fileNumber}"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                if (fileNumber == 1)
                {
                    try
                    {
                        var checkName = RegexHelpers.CheckDaneOgolneRegex(filePath);
                        if (!string.IsNullOrEmpty(checkName?.Error?.Code))
                        {
                            MessageBox.Show(checkName.Error.Message, checkName.Error.Code, MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        ((Label)Controls[3]).Text = Path.GetFileName(filePath);
                        file1Path = filePath;
                    }
                    catch
                    {
                        ((Label)Controls[3]).Text = GUIMessage.BrakWybranegoPliku;
                        file1Path = null;
                    }
                }
                else
                {
                    try
                    {
                        var cityAndDate = RegexHelpers.GetName(filePath);
                        if (!string.IsNullOrEmpty(cityAndDate?.Error?.Code))
                        {
                            MessageBox.Show(cityAndDate.Error.Message, cityAndDate.Error.Code, MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        city = cityAndDate.Value.City;
                        file2Path = filePath;
                        ((Label)Controls[1]).Text = Path.GetFileName(filePath);
                    }
                    catch
                    {
                        ((Label)Controls[1]).Text = GUIMessage.BrakWybranegoPliku;
                        file2Path = null;
                    }
                }
            }
        }

        private async void UploadButton_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Path.GetTempPath() + "ayp");
            if (Directory.Exists(path))
            {
                var files = Directory.GetFiles(path);

                foreach (var file in files)
                {
                    File.Delete(file);
                }
            }

            var program = RegexHelpers.GetProgram([.. filePaths.Values]);
            var reader = new FileReader();
            var hours = reader.FindWordInWordDocument(program).Value;
            if (!int.TryParse(hours, out var hoursParsed))
            {
                MessageBox.Show("Czas trwania kursu jest nieprawid³owy.", "B³¹d", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var result = reader.FindObjectInExcelDocument(file1Path);
            if (!string.IsNullOrEmpty(result?.Error?.Code))
            {
                MessageBox.Show(result.Error.Message, "B³¹d", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (string.IsNullOrEmpty(file1Path) || string.IsNullOrEmpty(file2Path))
            {
                MessageBox.Show("Prosze za³aduj dwa pliki Excel.", "B³¹d", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                var httpClientService = new HttpClientService();
                ISeleniumService seleniumService = new SeleniumService();
                string pathToPdfKrs = string.Empty;
                string pathToPdfNip = string.Empty;

                var ris = seleniumService.DownloadAypRis();
                var aypkrs = AESBruteForceDecryption.Encrypt("0000741233");
                var aypkrsfile = await httpClientService.PostKrsAsync($"https://prs-openapi2-prs-prod.apps.ocp.prod.ms.gov.pl/api/wyszukiwarka/OdpisPelny/pdf", aypkrs, "0000741233");
                var siedziba = new Adres();

                if (OperationHelpers.ValidateKRS(result.Value.KRS))
                {
                    var krs = AESBruteForceDecryption.Encrypt(result.Value.KRS);
                    pathToPdfKrs = await httpClientService.PostKrsAsync($"https://prs-openapi2-prs-prod.apps.ocp.prod.ms.gov.pl/api/wyszukiwarka/OdpisPelny/pdf", krs, result.Value.KRS);
                    var s = await httpClientService.GetAsync($"https://api-krs.ms.gov.pl/api/krs/odpisaktualny/{result.Value.KRS}?rejestr=P");
                    siedziba = s.Odpis.Dane.Dzial1.SiedzibaIAdres.Adres;
                }
                else
                {
                    pathToPdfNip = await seleniumService.DownloadComapnyByNipPdf(result.Value.NIP);
                }

                var pkd = Regex.Match(result.Value.PKD.Replace(".", ""), @"\d+[A-Z]").Value;
                var nrRachunku = Regex.Replace(result.Value.BankAccountDetails.AccountNumber, @"[^\d\s]", "").Trim();
                var uzasadnienie = Path.Combine(directoryPath, comboBoxFiles.SelectedItem.ToString());
                var isProcessed = await ProcessExcelFiles
                    (file1Path, file2Path, result.Value.AdditionalBusinessAddresses, city, pkd, nrRachunku, result.Value.EmploymentData.ContractEmployees, result.Value.PUPContact, selectedFiles, pathToPdfKrs, pathToPdfNip, ris, aypkrsfile, siedziba, uzasadnienie);
                if (!string.IsNullOrEmpty(isProcessed?.Error?.Code))
                {
                    MessageBox.Show($"Wyst¹pi³ b³¹d: {isProcessed.Error.Message}", "B³¹d", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                MessageBox.Show("Pliki zosta³y dodane pomyœlnie do formularza.", "Sukces", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Wyst¹pi³ b³¹d: {ex.Message}", "B³¹d", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            ((Label)Controls[1]).Text = GUIMessage.BrakWybranegoPliku;
            ((Label)Controls[3]).Text = GUIMessage.BrakWybranegoPliku;
            progressBar.Value = 0;
        }

        private async Task<Result<bool, Error>> ProcessExcelFiles(
            string filePath1, string filePath2, string businessAddresses, string city, string pkd, string nrRachunku, int liczbaZatrudnionych, ContactPerson contactPerson, string[] paths, string krs, string ceidg, string ris, string aypNip, Adres siedziba, string uzasadnienie)
        {
            try
            {
                var fileReader = new FileReader();
                ISeleniumService seleniumService = new SeleniumService();

                UpdateProgress(10);
                seleniumService.LoginToPage();

                List<string> tematySzkoleñ = [];
                var listaSzkoleñ = await fileReader.ReadExcelFileAsync(filePath2, TrainingFormPatterns.Patterns, "lista szkoleñ");
                List<string> models = [];
                var dataExcelCities = await fileReader.GetExcelWorksheetNames(filePath1);

                var dataWithExcelWorksheets = dataExcelCities
                    .Select(x => new { Item = x, Name = RegexHelpers.GetNameFromExcel(x) })
                    .Where(result => result.Name != null)
                    .Select(result => result.Name)
                    .ToList();
                dataWithExcelWorksheets.Add(new Models.RegexEntity.CityAndDateModel { City = "Kolonia Karna", StartDate = DateTime.UtcNow });
                var SataWithExcelWorksheets = dataExcelCities
                    .Select(z => new { Item = z, Name = RegexHelpers.GetNameFromExcel(z) }).ToList();

                var start = dataWithExcelWorksheets.OrderBy(x => x.StartDate).First().StartDate;
                string regexPattern = @"[A-Z¥ÆÊ£ÑÓŒ¯a-z¹æê³ñóœŸ¿]+(?:[_\s\.])?\d{4}(?:[_\s\.])?\d{2}(?:[_\s\.])?\d{2}"; // Regex to match CITY_YYYY.MM.DD
                string matchedFromFile = RegexHelpers.ExtractMatch(file1Path, regexPattern);
                var worksheetMatch = string.Empty;
                if (matchedFromFile != null)
                {
                    Console.WriteLine($"Found in file path: {matchedFromFile}");

                    worksheetMatch = RegexHelpers.FindMatchInWorksheets(matchedFromFile, SataWithExcelWorksheets.Select(x => x.Item).ToList());

                    Console.WriteLine(worksheetMatch != null
                        ? $"Match found in worksheets: {worksheetMatch}"
                        : "No match found in worksheets.");
                }
                else
                {
                    Console.WriteLine("No matching pattern found in file path.");
                }
                var getName = RegexHelpers.GetName(filePath1);
                var dataTables = await fileReader.ReadExcelFileAsync(filePath1, WorkerFormPatterns.Patterns, worksheetMatch);
                var dataTableAdditionalAllWorkers = await fileReader.ReadExcelFileAsync(filePath1, WorkerFormPatterns.Patterns, "Dane ogólne");
                var dataTablesCombinded = Clone(dataTables, dataTableAdditionalAllWorkers);
                await seleniumService.LoadForm(businessAddresses, city, pkd, nrRachunku, liczbaZatrudnionych, contactPerson, start, start, paths, krs, ceidg, ris, aypNip, siedziba);

                UpdateProgress(20);

                IFileWriter fileWriter = new FormFiller.Services.FileWriter();

                var pracownicyTable = await fileReader.ReadExcelFileAsync(filePath2, null, "lista osób");

                //var tmp = await fileWriter.ExcelWriterAsyncReversed(dataTables, filePath1);

                //for (var i = tmp.Item2.Rows.Count(); i > 1; i--)
                //{
                //    var row = fileReader.ReadExcelRow(tmp.Item1);
                //    seleniumService.ProvideWorkerInformation(row, city);
                //    await fileWriter.DeleteExcelRowAsync(tmp.Item1, i);
                //    UpdateProgress(30 + (int)((tmp.Item2.Rows.Count() - i) / (float)tmp.Item2.Rows.Count() * 20));
                //}

                var workersData = new List<WorkerData>();

                // Extract the first column (name and surname)
                var names = pracownicyTable.AsEnumerable()
                    .Select(row => row[0]?.ToString() ?? string.Empty) // First column as names
                    .ToList();

                // Extract column names for data columns (skipping the second column)
                var dataColumnNames = pracownicyTable.Columns
                    .Cast<DataColumn>()
                    .Skip(2) // Skip the second column
                    .Select(col => col.ColumnName)
                    .ToList();

                // Match names with their corresponding column data
                for (int i = 0; i < names.Count; i++) // Loop over rows
                {
                    var worker = new WorkerData { Name = names[i] }; // Create worker with name

                    for (int colIndex = 2; colIndex < pracownicyTable.Columns.Count; colIndex++) // Start from third column
                    {
                        var columnName = dataColumnNames[colIndex - 2]; // Adjust index to match dataColumnNames
                        var cellValue = pracownicyTable.Rows[i][colIndex]?.ToString()?.ToLower(); // Get cell value as string

                        // Interpret cell value as bool: "yes"/"true"/"1" -> true, otherwise false
                        var isTakingAction = cellValue == "yes" || cellValue == "true" || cellValue == "1";

                        worker.ColumnData[columnName] = isTakingAction;
                    }

                    workersData.Add(worker);
                }


                UpdateProgress(70);
                var matchingWorker = new List<DataRow>();
                var tmp2 = await fileWriter.ExcelWriterAsync(dataTablesCombinded, filePath1);
                var pracownicy = tmp2.Item2.Rows.Count();
                {
                    try
                    {
                        for (int i = 1; i < dataTablesCombinded.Rows.Count; i++)
                        {
                            var rowP = dataTablesCombinded.Rows[i];
                            var p = dataTablesCombinded.Rows[i][1]?.ToString()?.ToLower();
                            foreach (var w in workersData)
                            {
                                var ratio = Fuzz.Ratio(w.Name.ToLower(), p);
                                if (w.Name.ToLower() != p)
                                {
                                    if (ratio > 95)
                                    {
                                        matchingWorker.Add(rowP);
                                        break;
                                    }
                                }

                                if (w.Name.ToLower() == p)
                                {
                                    matchingWorker.Add(rowP);
                                    break;
                                }
                            }
                        }
                    }
                    catch
                    {

                    }
                    UpdateProgress(60);
                    foreach (var workerr in matchingWorker)
                    {
                        var row = fileReader.ReadExcelRow(workerr);
                        seleniumService.ProvideWorkerInformation(row, city);
                    }

                    UpdateProgress(70);

                    seleniumService.LoadPlanowanyRealizator();

                    seleniumService.LoadInfomacjeDotyczaceKsztalcenia(tmp2.Item2, listaSzkoleñ, pracownicyTable, uzasadnienie);

                    UpdateProgress(90);

                    seleniumService.ValidateWorkers();

                    //var worksheetNames = await fileReader.GetExcelWorksheetNames(filePath1);

                    //string pattern = @"^(?=.*_)(?=.*\.)(?=.*\d{4}).*$";
                    //var regex = new Regex(pattern);
                    //var matchingWorksheets = worksheetNames.Where(x => regex.IsMatch(x)).ToList();

                    var data = await fileReader.ReadExcelFileAsync(tmp2.Item1, WorkerFormPatterns.Patterns, null);
                    if (data.Rows.Count == 0)
                    {
                        File.Delete(tmp2.Item1);
                    }

                    UpdateProgress(100);
                    seleniumService.ZrobWydruk();
                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return Errors.ChromeProccessingError;
            }
        }
        List<string> FindMatchingWorkers(List<WorkerData> workersData, List<string> externalNames)
        {
            var workerNamesSet = new HashSet<string>(workersData.Select(w => w.Name));
            return externalNames.Where(name => workerNamesSet.Contains(name)).ToList();
        }
        private void UpdateProgress(int value)
        {
            if (progressBar.InvokeRequired)
            {
                progressBar.Invoke(new Action(() => progressBar.Value = value));
            }
            else
            {
                progressBar.Value = value;
            }
        }

        private DataTable Clone(DataTable dataTables, DataTable dataTableAdditionalAllWorkers)
        {
            var mergedDataTable = dataTables.Clone(); // Skopiuj strukturê tabeli.

            foreach (DataRow row in dataTables.Rows)
            {
                // Pobierz wartoœæ z kolumny NazwiskoImie.
                var nazwiskoImie = row[WorkersFormKeys.NazwiskoImie]?.ToString();

                if (!string.IsNullOrEmpty(nazwiskoImie))
                {
                    // ZnajdŸ dopasowany wiersz w dodatkowej tabeli.
                    var matchingRow = dataTableAdditionalAllWorkers
                        .AsEnumerable()
                        .FirstOrDefault(r => r[WorkersFormKeys.NazwiskoImie]?.ToString() == nazwiskoImie);

                    if (matchingRow != null)
                    {
                        // Dodaj now¹ kolumnê do mergedDataTable, jeœli nie istnieje.
                        if (!mergedDataTable.Columns.Contains(WorkersFormKeys.KwotaOtrzymanegoDofinansowania))
                        {
                            mergedDataTable.Columns.Add(WorkersFormKeys.KwotaOtrzymanegoDofinansowania, typeof(string));
                        }

                        // Skopiuj oryginalny wiersz do nowej tabeli.
                        var newRow = mergedDataTable.NewRow();
                        newRow.ItemArray = row.ItemArray;

                        // Dodaj wartoœæ KwotaOtrzymanegoDofinansowania.
                        newRow[WorkersFormKeys.KwotaOtrzymanegoDofinansowania] =
                            matchingRow[WorkersFormKeys.KwotaOtrzymanegoDofinansowania];

                        mergedDataTable.Rows.Add(newRow);
                    }
                }
            }
            return mergedDataTable;

        }
    }

    public class WorkerData
    {
        public string Name { get; set; } = string.Empty;
        public Dictionary<string, bool> ColumnData { get; set; } = new();
    }
}
