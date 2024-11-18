using FormFiller.Constants;
using FormFiller.Models.TrainingEntity;
using FormFiller.Services;

using System.Data;
using System.Text.RegularExpressions;

namespace WinFormsAutoFiller
{
    public partial class Form1 : Form
    {
        private string file1Path;
        private string file2Path;
        private ProgressBar progressBar;

        public Form1()
        {
            InitializeComponent();
            Text = "Uzupe³nianie formularza KFS";
            Size = new Size(800, 500);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;

            // Create and configure controls
            Label titleLabel = new Label
            {
                Text = "Uzupe³nianie formularza KFS",
                Font = new Font("Segoe UI", 18, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Top,
                Height = 50
            };

            Label descriptionLabel = new Label
            {
                Text = "Po przyciœniêciu przycisku \"Za³aduj pliki do formularza KFS\" otworzy Ci siê Google Chrome.",
                Font = new Font("Segoe UI", 10),
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Top
            };

            Label descriptionLabel2 = new Label
            {
                Text = "Zaloguj siê do pracuj.gov.pl",
                Font = new Font("Segoe UI", 10),
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Top
            };
            Label descriptionLabel3 = new Label
            {
                Text = "Przejdz do formularza PSZ-KFS i poczekaj a¿ za³aduj¹ siê dane.",
                Font = new Font("Segoe UI", 10),
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Top
            };

            Button file1Button = CreateFileButton("Wybierz plik Excel z danymi ogólnymi", 1);
            Button file2Button = CreateFileButton("Wybierz plik Excel z list¹ pracowników", 2);

            Label file1Label = CreateFileLabel(1);
            Label file2Label = CreateFileLabel(2);

            Button uploadButton = new Button
            {
                Text = "Za³aduj pliki do formularza KFS",
                Dock = DockStyle.Bottom,
                Height = 40,
                Font = new Font("Segoe UI", 12),
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
            Controls.Add(progressBar);


            // Set background color
            BackColor = Color.FromArgb(240, 240, 240);
        }

        private Button CreateFileButton(string text, int fileNumber)
        {
            Button button = new Button
            {
                Text = text,
                Size = new Size(400, 35),
                Location = new Point(200, 170 + (fileNumber - 1) * 80),
                Font = new Font("Segoe UI", 10),
                BackColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            button.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
            button.Click += (sender, e) => FileButton_Click(fileNumber);
            return button;
        }

        private Label CreateFileLabel(int fileNumber)
        {
            return new Label
            {
                Text = "Brak wybranego pliku.",
                Size = new Size(400, 20),
                Location = new Point(200, 210 + (fileNumber - 1) * 80),
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.Gray,
                TextAlign = ContentAlignment.MiddleCenter
            };
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
                    file1Path = filePath;
                    ((Label)Controls[3]).Text = Path.GetFileName(filePath);
                }
                else
                {
                    file2Path = filePath;
                    ((Label)Controls[1]).Text = Path.GetFileName(filePath);
                }
            }
        }

        private async void UploadButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(file1Path) || string.IsNullOrEmpty(file2Path))
            {
                MessageBox.Show("Prosze za³aduj dwa pliki Excel.", "B³ad", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                await ProcessExcelFiles(file1Path, file2Path);
                MessageBox.Show("Pliki zosta³y dodane pomyœlnie do formularza.", "Sukces", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Wyst¹pi³ b³ad: {ex.Message}", "B³ad", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            // Reset file paths and labels after processing
            file1Path = null;
            file2Path = null;
            ((Label)Controls[1]).Text = "Brak wybranego pliku.";
            ((Label)Controls[3]).Text = "Brak wybranego pliku.";
            progressBar.Value = 0;
        }

        private async Task ProcessExcelFiles(string filePath1, string filePath2)
        {
            try
            {
                // Simulating the processing of the provided code
                var fileReader = new FileReader();
                ISeleniumService seleniumService = new SeleniumService();

                UpdateProgress(10);
                seleniumService.LoginToPage();
                seleniumService.LoadWorkerForm();

                var dataTables = await fileReader.ReadExcelFileAsync(filePath1, WorkerFormPatterns.Patterns, "Dane ogólne");
                UpdateProgress(20);

                IFileWriter fileWriter = new FormFiller.Services.FileWriter();
                var tmp = await fileWriter.ExcelWriterAsyncReversed(dataTables, filePath1);

                for (var i = tmp.Item2.Rows.Count(); i > 1; i--)
                {
                    var row = fileReader.ReadExcelRow(tmp.Item1);
                    seleniumService.ProvideWorkerInformation(row);
                    await fileWriter.DeleteExcelRowAsync(tmp.Item1, i);
                    UpdateProgress(30 + (int)((tmp.Item2.Rows.Count() - i) / (float)tmp.Item2.Rows.Count() * 20));
                }

                var data = fileReader.ReadExcelFileAsync(tmp.Item1, WorkerFormPatterns.Patterns, null).Result;
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
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
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
