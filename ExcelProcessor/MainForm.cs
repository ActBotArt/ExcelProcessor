using System;
using System.Drawing;
using System.Windows.Forms;
using System.IO;

namespace ExcelProcessor
{
    public partial class MainForm : Form
    {
        private string projectPath;
        private Button btnSelectFiles;
        private Button btnCompile;
        private ListBox listFiles;
        private Label lblStatus;
        private Label lblTitle;
        private TableLayoutPanel mainLayout;

        public MainForm()
        {
            InitializeComponents();
            CreateProjectFolder();
        }

        private void InitializeComponents()
        {
            // Настройка формы
            this.Text = "Excel в SQL конвертер";
            this.Size = new Size(800, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimizeBox = true;
            this.MaximizeBox = false;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;

            // Создание компонентов
            mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                Padding = new Padding(10),
                RowStyles = {
                    new RowStyle(SizeType.Absolute, 40),  // Заголовок
                    new RowStyle(SizeType.Absolute, 50),  // Кнопки
                    new RowStyle(SizeType.Percent, 100),  // Список
                    new RowStyle(SizeType.Absolute, 30)   // Статус
                }
            };

            lblTitle = new Label
            {
                Text = "Конвертер Excel файлов в SQL скрипт",
                Font = new Font("Segoe UI", 14, FontStyle.Bold),
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };

            var buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false
            };

            btnSelectFiles = new Button
            {
                Text = "Выбрать Excel файлы",
                Size = new Size(200, 40),
                Font = new Font("Segoe UI", 10),
                Margin = new Padding(0, 0, 10, 0)
            };
            btnSelectFiles.Click += BtnSelectFiles_Click;

            btnCompile = new Button
            {
                Text = "Создать SQL скрипт",
                Size = new Size(200, 40),
                Font = new Font("Segoe UI", 10),
                Enabled = false
            };
            btnCompile.Click += BtnCompile_Click;

            buttonPanel.Controls.AddRange(new Control[] { btnSelectFiles, btnCompile });

            listFiles = new ListBox
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 10),
                BorderStyle = BorderStyle.FixedSingle
            };

            lblStatus = new Label
            {
                AutoSize = true,
                Font = new Font("Segoe UI", 9),
                Dock = DockStyle.Fill
            };

            // Добавление компонентов в layout
            mainLayout.Controls.Add(lblTitle, 0, 0);
            mainLayout.Controls.Add(buttonPanel, 0, 1);
            mainLayout.Controls.Add(listFiles, 0, 2);
            mainLayout.Controls.Add(lblStatus, 0, 3);

            this.Controls.Add(mainLayout);
        }

        private void CreateProjectFolder()
        {
            try
            {
                string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                projectPath = Path.Combine(desktop, $"ExcelToSQL_{DateTime.Now:yyyyMMdd_HHmmss}");

                if (!Directory.Exists(projectPath))
                {
                    Directory.CreateDirectory(projectPath);
                }

                UpdateStatus($"Создана рабочая папка: {projectPath}");
            }
            catch (Exception ex)
            {
                string errorMessage = $"Ошибка при создании рабочей папки:\n{ex.Message}";
                MessageBox.Show(errorMessage, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus(errorMessage, true);
                throw;
            }
        }

        private void BtnSelectFiles_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Multiselect = true;
                ofd.Filter = "Excel Files (*.xlsx, *.xls)|*.xlsx;*.xls|All Files (*.*)|*.*";
                ofd.Title = "Выберите Excel файлы";
                ofd.CheckFileExists = true;
                ofd.CheckPathExists = true;

                try
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        listFiles.Items.Clear();
                        foreach (string sourceFile in ofd.FileNames)
                        {
                            try
                            {
                                if (!File.Exists(sourceFile))
                                {
                                    throw new FileNotFoundException($"Файл не найден: {sourceFile}");
                                }

                                string fileName = Path.GetFileName(sourceFile);
                                string destPath = Path.Combine(projectPath, fileName);

                                // Проверяем, существует ли директория назначения
                                if (!Directory.Exists(projectPath))
                                {
                                    Directory.CreateDirectory(projectPath);
                                }

                                File.Copy(sourceFile, destPath, true);
                                listFiles.Items.Add(fileName);
                                UpdateStatus($"Добавлен файл: {fileName}");
                            }
                            catch (Exception ex)
                            {
                                string errorMessage = $"Ошибка при копировании файла {Path.GetFileName(sourceFile)}:\n{ex.Message}";
                                MessageBox.Show(errorMessage, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                UpdateStatus(errorMessage, true);
                            }
                        }
                        btnCompile.Enabled = listFiles.Items.Count > 0;
                    }
                }
                catch (Exception ex)
                {
                    string errorMessage = $"Ошибка при выборе файлов:\n{ex.Message}";
                    MessageBox.Show(errorMessage, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UpdateStatus(errorMessage, true);
                }
            }
        }

        private void BtnCompile_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                btnCompile.Enabled = false;
                btnSelectFiles.Enabled = false;
                UpdateStatus("Создание SQL скрипта...");

                var processor = new ExcelFileProcessor(projectPath);
                string sqlFilePath = processor.ConvertToSql();

                UpdateStatus("SQL скрипт успешно создан!");

                var result = MessageBox.Show(
                    $"SQL скрипт создан успешно!\nПуть к файлу: {sqlFilePath}\n\nОткрыть папку с результатом?",
                    "Готово",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information
                );

                if (result == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start("explorer.exe", projectPath);
                }
            }
            catch (Exception ex)
            {
                string errorMessage = $"Ошибка при конвертации:\n{ex.Message}";
                MessageBox.Show(errorMessage, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus(errorMessage, true);
            }
            finally
            {
                Cursor = Cursors.Default;
                btnCompile.Enabled = listFiles.Items.Count > 0;
                btnSelectFiles.Enabled = true;
            }
        }

        private void UpdateStatus(string message, bool isError = false)
        {
            if (lblStatus.InvokeRequired)
            {
                lblStatus.Invoke(new Action(() => UpdateStatus(message, isError)));
                return;
            }

            lblStatus.ForeColor = isError ? Color.Red : Color.Black;
            lblStatus.Text = message;
            Application.DoEvents();
        }
    }
}