using FormFiller.Constants;
using FormFiller.Models.TrainingEntity;
using FormFiller.Services;

using System.Data;
using System.Text.RegularExpressions;

using WinFormsAutoFiller.Encryption;
using WinFormsAutoFiller.Helpers;
using WinFormsAutoFiller.Models.PytaniaDoWnioskuEntity;
using WinFormsAutoFiller.Services;
using WinFormsAutoFiller.Utilis;

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

        public Form1()
        {
            InitializeComponent();
            Text = "Uzupe³nianie formularza KFS";
            Size = new Size(1400, 800);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;

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

            // Set background color
            BackColor = Color.FromArgb(240, 240, 240);
            file3Button.AllowDrop = true;
            file3Button.DragEnter += File3Button_DragEnter;
            file3Button.DragDrop += File3Button_DragDrop;
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

            //if (string.IsNullOrEmpty(file1Path) || string.IsNullOrEmpty(file2Path))
            //{
            //    MessageBox.Show("Prosze za³aduj dwa pliki Excel.", "B³¹d", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    return;
            //}

            try
            {
                var httpClientService = new HttpClientService();
                ISeleniumService seleniumService = new SeleniumService();
                string pathToPdfKrs;
                string pathToPdfNip;
                if (!OperationHelpers.ValidateKRS(result.Value.KRS))
                {
                    var krs = AESBruteForceDecryption.Encrypt(result.Value.KRS);
                    pathToPdfKrs = await httpClientService.PostKrsAsync($"https://prs-openapi2-prs-prod.apps.ocp.prod.ms.gov.pl/api/wyszukiwarka/OdpisPelny/pdf", krs, result.Value.KRS);

                }
                else
                {
                    pathToPdfNip = await seleniumService.DownloadComapnyByNipPdf("5842859530");
                }
                var pkd = Regex.Match(result.Value.PKD.Replace(".", ""), @"\d+[A-Z]").Value;
                var nrRachunku = Regex.Replace(result.Value.BankAccountDetails.AccountNumber, @"[^\d\s]", "").Trim();

                var isProcessed = await ProcessExcelFiles(file1Path, file2Path, result.Value.AdditionalBusinessAddresses, city, pkd, nrRachunku, result.Value.EmploymentData.ContractEmployees, );
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
            ((Label)Controls[1]).Text = GUIMessage.BrakWybranegoPliku;
            ((Label)Controls[3]).Text = GUIMessage.BrakWybranegoPliku;
            progressBar.Value = 0;
        }

        private async Task<Result<bool, Error>> ProcessExcelFiles(string filePath1, string filePath2, string businessAddresses, string city, string pkd, string nrRachunku, int liczbaZatrudnionych, ContactPerson contactPerson, DateTime startDate, DateTime endDate, string[] paths, string krs, string ceidg)
        {
            try
            {
                var fileReader = new FileReader();
                ISeleniumService seleniumService = new SeleniumService();

                UpdateProgress(10);
                seleniumService.LoginToPage();
                await seleniumService.LoadForm(businessAddresses, city, pkd, nrRachunku, liczbaZatrudnionych, contactPerson, startDate, endDate, paths, krs, ceidg);

                var dataTables = await fileReader.ReadExcelFileAsync(filePath1, WorkerFormPatterns.Patterns, "Dane ogólne");
                UpdateProgress(20);

                IFileWriter fileWriter = new FormFiller.Services.FileWriter();
                var tmp = await fileWriter.ExcelWriterAsyncReversed(dataTables, filePath1);

                for (var i = tmp.Item2.Rows.Count(); i > 1; i--)
                {
                    var row = fileReader.ReadExcelRow(tmp.Item1);
                    seleniumService.ProvideWorkerInformation(row, city);
                    await fileWriter.DeleteExcelRowAsync(tmp.Item1, i);
                    UpdateProgress(30 + (int)((tmp.Item2.Rows.Count() - i) / (float)tmp.Item2.Rows.Count() * 20));
                }

                var data = await fileReader.ReadExcelFileAsync(tmp.Item1, WorkerFormPatterns.Patterns, null);
                if (data.Rows.Count == 0)
                {
                    File.Delete(tmp.Item1);
                }

                UpdateProgress(60);

                seleniumService.LoadPlanowanyRealizator();

                UpdateProgress(70);

                var tmp2 = await fileWriter.ExcelWriterAsync(dataTables, filePath1);
                var pracownicy = tmp2.Item2.Rows.Count();

                var listaSzkoleñ = await fileReader.ReadExcelFileAsync(filePath2, TrainingFormPatterns.Patterns, "lista szkoleñ");
                var pracownicyTable = await fileReader.ReadExcelFileAsync(filePath2, null, "lista osób");

                UpdateProgress(80);

                seleniumService.LoadInfomacjeDotyczaceKsztalcenia(tmp2.Item2, listaSzkoleñ, pracownicyTable);

                File.Delete(tmp2.Item1);

                UpdateProgress(90);

                var worksheetNames = await fileReader.GetExcelWorksheetNames(filePath1);

                string pattern = @"^(?=.*_)(?=.*\.)(?=.*\d{4}).*$";
                var regex = new Regex(pattern);
                var matchingWorksheets = worksheetNames.Where(x => regex.IsMatch(x)).ToList();

                UpdateProgress(100);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return Errors.ChromeProccessingError;
            }
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
    }
}
