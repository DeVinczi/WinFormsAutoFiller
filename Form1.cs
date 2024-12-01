using AutoFiller.Models;

using FormFiller.Constants;
using FormFiller.Models.TrainingEntity;
using FormFiller.Models.WorkerEntity;
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
        private List<string> selectedFiles = [];
        private Dictionary<string, string> filePaths;
        private string city;
        private ComboBox comboBoxFiles;
        private ComboBox workSheets;
        private Label labelUzasadnienia;
        private Label labelZak³adka;
        private CheckBox czyJestMikro;
        string directoryPath = Path.Combine(Application.StartupPath, "Uzasadnienia");

        public Form1()
        {
            InitializeComponent();
            Text = "Uzupe³nianie formularza KFS";
            Size = new Size(1400, 800);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            AutoScaleMode = AutoScaleMode.Dpi;

            labelUzasadnienia = new Label();
            labelUzasadnienia.Text = "Uzasadnienia";
            labelUzasadnienia.Font = new Font("Segoe UI", 12);// Set the text for the label
            labelUzasadnienia.Location = new System.Drawing.Point(10, 370); // Position the label
            labelUzasadnienia.Width = 300;  // Set width of the label

            comboBoxFiles = new ComboBox();
            comboBoxFiles.Location = new System.Drawing.Point(10, 400);  // Position the ComboBox below the label
            comboBoxFiles.Width = 380;  // Set width of the ComboBox

            workSheets = new ComboBox();
            workSheets.Location = new Point(920, 230);
            workSheets.Width = 300;
            labelZak³adka = new Label();
            labelZak³adka.Text = "Zak³adki:";
            labelZak³adka.Font = new Font("Segoe UI", 11);// Set the text for the label
            labelZak³adka.Location = new System.Drawing.Point(920, 200); // Position the label
            labelZak³adka.Width = 300;

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

            czyJestMikro = new CheckBox
            {
                Text = "Czy jest mikroprzedsiêbiorc¹?",
                AutoSize = true, // Automatically size the checkbox to fit the text
                Font = new Font("Segoe UI", 9), // Stylish font
                Location = new Point(940, 320),
                FlatStyle = FlatStyle.Flat, // Flat style for a modern look
                CheckAlign = ContentAlignment.MiddleLeft, // Align checkbox to the right of text
                TextAlign = ContentAlignment.MiddleRight, // Align text to the left
                Checked = true
            };

            czyJestMikro.Click += (sender, e) =>
            {
                czyJestMikro.Checked = !czyJestMikro.Checked;
            };

            czyJestMikro.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
            czyJestMikro.FlatAppearance.BorderSize = 1;

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
            Controls.Add(workSheets);
            Controls.Add(labelZak³adka);
            Controls.Add(czyJestMikro);

            // Set background color
            BackColor = Color.FromArgb(240, 240, 240);
            file3Button.AllowDrop = true;
            file3Button.DragEnter += File3Button_DragEnter;
            file3Button.DragDrop += File3Button_DragDrop;

            czyJestMikro.AutoCheck = false;
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
            button.AllowDrop = true;
            button.DragEnter += (sender, e) => Button_DragEnter(sender, e, fileNumber);
            button.DragDrop += (sender, e) => Button_DragDrop(sender, e, fileNumber);
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

            ToolTip toolTip = new ToolTip
            {
                ToolTipIcon = ToolTipIcon.Info,
                ToolTipTitle = "Pliki wybrane"
            };

            label.MouseHover += (sender, e) =>
            {
                if (label.Text.Length > 100)
                {
                    toolTip.SetToolTip(label, label.Text);
                }
            };

            return label;
        }

        private void Button_DragEnter(object sender, DragEventArgs e, int fileNumber)
        {
            // Check if the dragged data contains file paths and is an Excel file
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

                // Check if the file is an Excel file (.xls or .xlsx)
                if (files.Length == 1 && IsExcelFile(files[0]))
                {
                    e.Effect = DragDropEffects.Copy;  // Show the copy cursor
                }
                else
                {
                    e.Effect = DragDropEffects.None;  // Show no effect if it's not an Excel file
                }
            }
            else
            {
                e.Effect = DragDropEffects.None;  // Show no effect if it's not a file
            }
        }

        private bool IsExcelFile(string filePath)
        {
            string[] allowedExtensions = { ".xlsx", ".xls", ".xlsm", ".xlsb", ".xltx", ".xlw" };
            return allowedExtensions.Contains(Path.GetExtension(filePath).ToLower());
        }

        private async void Button_DragDrop(object sender, DragEventArgs e, int fileNumber)
        {
            try
            {
                // Get the file paths from the dragged files
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

                // Only process the drop if it's a single file (Excel)
                if (files.Length == 1 && IsExcelFile(files[0]))
                {
                    if (fileNumber == 1)
                    {
                        var fileReader = new FileReader();
                        var getNames = await fileReader.GetExcelWorksheetNames(files[0]);
                        if (getNames.Count == 0)
                        {
                            MessageBox.Show("Excel nie ma zak³adek do pokazania", "B³¹d", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        foreach (var name in getNames)
                        {
                            workSheets.Items.Add(name);
                        }
                        file1Path = files[0];
                        ((Label)Controls[3]).Text = Path.GetFileName(files[0]);
                    }
                    if (fileNumber == 2)
                    {
                        file2Path = files[0];
                        ((Label)Controls[1]).Text = Path.GetFileName(files[0]);
                    }
                }
            }
            catch
            {
                throw;
            }
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
                selectedFiles = openFileDialog.FileNames.ToList();  // Store selected files
                filePaths = AddFilePaths(openFileDialog.FileNames);

                // Display the selected file names in the label
                ((Label)Controls[9]).Text = string.Join(Environment.NewLine, selectedFiles.Select(file => Path.GetFileName(file)));
            }
        }
        private void LoadFiles()
        {
            string directoryPath = Path.Combine(Application.StartupPath, "Uzasadnienia");

            if (Directory.Exists(directoryPath))
            {
                string[] files = Directory.GetFiles(directoryPath);

                foreach (string file in files)
                {
                    comboBoxFiles.Items.Add(Path.GetFileName(file));
                }
            }
            else
            {
                MessageBox.Show("Folder 'Uzasadnienia' nie istnieje.");
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
            selectedFiles.AddRange(files);  // Store the file paths

            filePaths = AddFilePaths(selectedFiles.ToArray());
            // Display the selected file names in the label
            var z = string.Join(Environment.NewLine, selectedFiles.Select(file => Path.GetFileName(file)));
            ((Label)Controls[9]).Text = z;
        }

        private async void FileButton_Click(int fileNumber)
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
                        var fileReader = new FileReader();
                        var getNames = await fileReader.GetExcelWorksheetNames(filePath);
                        if (getNames.Count == 0)
                        {
                            MessageBox.Show("Excel nie ma zak³adek do pokazania", "B³¹d", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        foreach (var name in getNames)
                        {
                            workSheets.Items.Add(name);
                        }
                        ((Label)Controls[3]).Text = Path.GetFileName(filePath);
                        file1Path = filePath;
                    }
                    catch
                    {
                        throw;
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

                        file2Path = filePath;
                        ((Label)Controls[1]).Text = Path.GetFileName(filePath);
                    }
                    catch
                    {
                    }
                }
            }
        }

        private async void UploadButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(workSheets.SelectedItem.ToString()))
                {
                    MessageBox.Show("Nie wybra³eœ zak³adki", "B³¹d", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (string.IsNullOrEmpty(comboBoxFiles.SelectedItem.ToString()))
                {
                    MessageBox.Show("Nie wybra³eœ uzasadnienia", "B³¹d", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                city = RegexHelpers.GetNameFromExcel(workSheets.SelectedItem.ToString()).City;
                var path = Path.Combine(Path.GetTempPath() + "ayp");
                if (Directory.Exists(path))
                {
                    var files = Directory.GetFiles(path);

                    foreach (var file in files)
                    {
                        File.Delete(file);
                    }
                }

                var reader = new FileReader();
                var wordReader = new DocumentReader();

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

                    if (!string.IsNullOrEmpty(result.Value.KRS) && result.Value.KRS.Any(char.IsDigit))
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
                    var krsValue = result.Value.KRS ?? string.Empty;
                    var isProcessed = await ProcessExcelFiles
                        (file1Path, file2Path, result.Value.AdditionalBusinessAddresses, city, pkd, nrRachunku, result.Value.EmploymentData.ContractEmployees, result.Value.PUPContact, selectedFiles.ToArray(), pathToPdfKrs, pathToPdfNip, ris, aypkrsfile, siedziba, uzasadnienie, workSheets.SelectedItem.ToString(), krsValue);
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

                file1Path = null;
                file2Path = null;
                filePaths = null;

                comboBoxFiles.SelectedIndex = -1;
                workSheets.SelectedIndex = -1;
                workSheets.Items.Clear();
                ((Label)Controls[1]).Text = GUIMessage.BrakWybranegoPliku;
                ((Label)Controls[3]).Text = GUIMessage.BrakWybranegoPliku;
                ((Label)Controls[9]).Text = GUIMessage.BrakWybranegoPliku;
                progressBar.Value = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "B³¹d", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }

        private async Task<Result<bool, Error>> ProcessExcelFiles(
            string filePath1, string filePath2, string businessAddresses, string city, string pkd, string nrRachunku, int liczbaZatrudnionych,
            ContactPerson contactPerson, string[] paths, string krs, string ceidg, string ris, string aypNip, Adres siedziba, string uzasadnienie, string nazwaZak³adki, string krsValue)
        {
            try
            {
                var fileReader = new FileReader();
                var wordReader = new DocumentReader();
                ISeleniumService seleniumService = new SeleniumService();

                UpdateProgress(10);
                seleniumService.LoginToPage();

                List<string> tematySzkoleñ = [];
                var listaSzkoleñ = await fileReader.ReadExcelFileAsync(filePath2, TrainingFormPatterns.Patterns, "lista szkoleñ");
                List<string> models = [];
                var dataExcelCities = await fileReader.GetExcelWorksheetNames(filePath1);
                var programs = RegexHelpers.GetProgram([.. filePaths.Values]);
                var hoursDict = new Dictionary<string, string>();
                foreach (var program in programs)
                {
                    var hours = await wordReader.FindWordInWordDocumentAsync(program.Key);
                    if (!string.IsNullOrEmpty(hours?.Error?.Message))
                    {
                        throw new Exception(hours.Error.Message);
                    }

                    if (!int.TryParse(hours.Value, out var hoursParsed))
                    {
                        MessageBox.Show("Czas trwania kursu jest nieprawid³owy.", "B³¹d", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        throw new Exception("Czas trwania kursu jest nieprawid³owy.");
                    }

                    hoursDict.Add(program.Key, hours.Value);
                }

                var getName = RegexHelpers.GetName(filePath1);
                var dataTables = await fileReader.ReadExcelFileAsync(filePath1, WorkerFormPatterns.Patterns, nazwaZak³adki);
                var dataTableAdditionalAllWorkers = await fileReader.ReadExcelFileAsync(filePath1, WorkerFormPatterns.Patterns, "Dane ogólne");
                List<DateTime> list = [];
                for (int i = 1; i < listaSzkoleñ.Rows.Count; i++)
                {
                    var x = listaSzkoleñ.Rows[i][TrainingFormKeys.TematSzkolenia].ToString();
                    var s = ColumnHelpers.ReadWorksheetToDataTable(filePath1, nazwaZak³adki);
                    var columnData = ColumnHelpers.FetchColumnData(s, x);
                    foreach (var col in columnData)
                    {
                        var (rangeCount, dates) = RegexHelpers.ExtractDatesWithRangeInfo(col);
                        list.AddRange(dates);
                    }
                }
                var start = list.OrderBy(x => x.Date).First();
                var end = list.OrderByDescending(x => x.Date).First();
                var dataTablesCombinded = FileReader.Clone(dataTables, dataTableAdditionalAllWorkers);

                var columnWithPulaKFS = ColumnHelpers.FetchColumnData(dataTables, "Priorytet");
                var ishigherThan10 = columnWithPulaKFS.Select(KFSPriority.MapPriority).Select(x => int.Parse(x.TrimEnd('\\'))).ToList();
                var ishigher = ishigherThan10.Any(number => number <= 9);

                await seleniumService.LoadForm(businessAddresses, city, pkd, nrRachunku, liczbaZatrudnionych, contactPerson, start, end, paths, krs, ceidg, ris, aypNip, siedziba, krsValue, ishigher, czyJestMikro.Checked);

                UpdateProgress(20);

                IFileWriter fileWriter = new FormFiller.Services.FileWriter();

                var pracownicyTable = await fileReader.ReadExcelFileAsync(filePath2, null, "lista osób");

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

                    seleniumService.LoadInfomacjeDotyczaceKsztalcenia(tmp2.Item2, listaSzkoleñ, pracownicyTable, uzasadnienie, hoursDict);

                    UpdateProgress(90);
                    await Task.Delay(1000);
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
            catch
            {
                throw;
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

        private static void GetErrorsHandled(Error error)
        {
            if (!string.IsNullOrEmpty(error.Code))
            {
                MessageBox.Show(error.Message, "B³¹d", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
    }
}
