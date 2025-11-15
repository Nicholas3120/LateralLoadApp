using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace LateralLoadApp
{
    public partial class MainForm : Form
    {
        // ======= UI Controls =======
        TextBox txtFolder, txtColFile, txtWallFile, txtCoorFile, txtZValue, txtOutputFile, txtOutput;
        Button btnBrowseFolder, btnRun, btnBrowseCol, btnBrowseWall, btnBrowseCoor;
        Label lblLoading;

        // Spinner animation
        bool spinnerRunning = false;
        string[] spinnerChars = new[] { "⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏" };
        int spinnerIndex = 0;
        Timer spinnerTimer;

        public MainForm()
        {
            // Step 1: Build UI
            BuildUI();

            // Step 2: Initialize Spinner for loading animation
            InitSpinner();

            // Step 3: Initialize default values for folder and output file
            InitDefaultValues();
        }

        // =============================
        // BUILD UI PROGRAMMATICALLY
        // =============================
        private void BuildUI()
        {
            this.Text = "Lateral Load Processor";
            this.Width = 900;
            this.Height = 650;
            this.StartPosition = FormStartPosition.CenterScreen;

            int xLabel = 20;
            int xInput = 200;
            int y = 20;
            int dy = 35;

            // Folder Path Label + TextBox + Browse Button
            CreateLabel("Folder Path:", xLabel, y);
            txtFolder = CreateTextbox(xInput, y, 350);
            btnBrowseFolder = CreateButton("Browse folder...", 600, y - 2, 140, btnBrowseFolder_Click);
            y += dy;

            // Column File Label + TextBox + Browse Button
            CreateLabel("Column File:", xLabel, y);
            txtColFile = CreateTextbox(xInput, y, 350);
            btnBrowseCol = CreateButton("Browse...", 600, y - 2, 120, btnBrowseCol_Click);
            y += dy;

            // Wall File Label + TextBox + Browse Button
            CreateLabel("Wall File:", xLabel, y);
            txtWallFile = CreateTextbox(xInput, y, 350);
            btnBrowseWall = CreateButton("Browse...", 600, y - 2, 120, btnBrowseWall_Click);
            y += dy;

            // Coordinate File Label + TextBox + Browse Button
            CreateLabel("Coordinate File:", xLabel, y);
            txtCoorFile = CreateTextbox(xInput, y, 350);
            btnBrowseCoor = CreateButton("Browse...", 600, y - 2, 120, btnBrowseCoor_Click);
            y += dy;

            // Z Elevation Label + TextBox
            CreateLabel("Z Elevation (mm):", xLabel, y);
            txtZValue = CreateTextbox(xInput, y, 150);
            y += dy;

            // Output File Label + TextBox
            CreateLabel("Output Filename:", xLabel, y);
            txtOutputFile = CreateTextbox(xInput, y, 350);
            y += dy;

            // Run Button + Loading Spinner Label
            btnRun = CreateButton("Run", xInput, y, 120, btnRun_Click);
            lblLoading = new Label
            {
                Left = xInput + 140,
                Top = y + 5,
                Width = 250,
                Text = ""
            };
            this.Controls.Add(lblLoading);

            y += dy + 10;

            // Output TextBox (ReadOnly)
            txtOutput = new TextBox
            {
                Left = 20,
                Top = y,
                Width = 830,
                Height = 320,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true
            };
            this.Controls.Add(txtOutput);
        }

        private Label CreateLabel(string text, int x, int y)
        {
            var lbl = new Label
            {
                Text = text,
                Left = x,
                Top = y + 5,
                AutoSize = true
            };
            this.Controls.Add(lbl);
            return lbl;
        }

        private TextBox CreateTextbox(int x, int y, int width)
        {
            var txt = new TextBox
            {
                Left = x,
                Top = y,
                Width = width
            };
            this.Controls.Add(txt);
            return txt;
        }

        private Button CreateButton(string text, int x, int y, int width, EventHandler onClick)
        {
            var btn = new Button
            {
                Text = text,
                Left = x,
                Top = y,
                Width = width
            };
            btn.Click += onClick;
            this.Controls.Add(btn);
            return btn;
        }

        // =============================
        // SPINNER
        // =============================
        private void InitSpinner()
        {
            spinnerTimer = new Timer();
            spinnerTimer.Interval = 100;
            spinnerTimer.Tick += SpinnerTimer_Tick;
        }

        private void SpinnerTimer_Tick(object sender, EventArgs e)
        {
            if (!spinnerRunning)
            {
                lblLoading.Text = "";
                spinnerTimer.Stop();
                return;
            }

            lblLoading.Text = $"Processing... {spinnerChars[spinnerIndex % spinnerChars.Length]}";
            spinnerIndex++;
        }

        private void InitDefaultValues()
        {
            txtFolder.Text = @"D:\Lateral Loads";  // Set a default folder path (changeable)
            txtOutputFile.Text = "unique_points.xlsx";
        }

        // =============================
        // FILE BROWSE HELPER METHODS
        // =============================
        private void btnBrowseFolder_Click(object sender, EventArgs e)
        {
            using (var dlg = new FolderBrowserDialog())
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                    txtFolder.Text = dlg.SelectedPath;
            }
        }

        private void BrowseExcelFileFor(TextBox targetTextBox)
        {
            using (var ofd = new OpenFileDialog())
            {
                ofd.Title = "Select Excel file";
                ofd.Filter = "Excel Files|*.xlsx;*.xlsm;*.xls|All Files|*.*";

                if (!string.IsNullOrWhiteSpace(txtFolder.Text) && Directory.Exists(txtFolder.Text))
                {
                    ofd.InitialDirectory = txtFolder.Text;
                }

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    string fullPath = ofd.FileName;
                    string dir = Path.GetDirectoryName(fullPath);
                    string fileName = Path.GetFileName(fullPath);

                    txtFolder.Text = dir;
                    targetTextBox.Text = fileName;
                }
            }
        }

        private void btnBrowseCol_Click(object sender, EventArgs e) => BrowseExcelFileFor(txtColFile);
        private void btnBrowseWall_Click(object sender, EventArgs e) => BrowseExcelFileFor(txtWallFile);
        private void btnBrowseCoor_Click(object sender, EventArgs e) => BrowseExcelFileFor(txtCoorFile);

        // =============================
        // MAIN DATA PROCESSOR
        // =============================
        private (double Fx, double Fy, double Fz, double Mx, double My, string OutputPath)
            ProcessData(
                string folderPath,
                string colFile,
                string wallFile,
                string coorFile,
                double zCoor,
                string outputFilename)
        {
            string colPath = Path.Combine(folderPath, colFile);
            string wallPath = Path.Combine(folderPath, wallFile);
            string coorPath = Path.Combine(folderPath, coorFile);

            foreach (var f in new[] { colPath, wallPath, coorPath })
                if (!File.Exists(f))
                    throw new FileNotFoundException($"File not found: {f}");

            var colDt = ReadExcelToDataTable(colPath);
            var wallDt = ReadExcelToDataTable(wallPath);
            var coorDt = ReadExcelToDataTable(coorPath);

            var colData = AddCoords(colDt, coorDt);
            var wallData = AddCoords(wallDt, coorDt);

            var allData = colData.Concat(wallData)
                                 .Where(p => p.Z.HasValue)
                                 .ToList();

            if (!allData.Any())
                throw new Exception("No coordinate data found.");

            double tol = 1e-6;
            var filtered = allData
                .Where(p => p.Z.HasValue && Math.Abs(p.Z.Value - zCoor) < tol)
                .ToList();

            if (!filtered.Any())
                throw new Exception($"No points found at Z = {zCoor}");

            var grouped = filtered
                .GroupBy(p => new { X = p.X ?? 0.0, Y = p.Y ?? 0.0, Z = p.Z ?? 0.0 })
                .Select(g => new
                {
                    g.Key.X,
                    g.Key.Y,
                    g.Key.Z,
                    Fx = g.Sum(p => p.Fx),
                    Fy = g.Sum(p => p.Fy),
                    Fz = g.Sum(p => p.Fz),
                    Mx = g.Sum(p => p.Mx),
                    My = g.Sum(p => p.My),
                    Mz = g.Sum(p => p.Mz)
                })
                .ToList();

            double FxTotal = grouped.Sum(r => r.Fx);
            double FyTotal = grouped.Sum(r => r.Fy);
            double FzTotal = grouped.Sum(r => r.Fz);
            double MYTotal = -grouped.Sum(r => r.Fz * r.X) / 1000.0 - grouped.Sum(r => r.My);
            double MXTotal = -grouped.Sum(r => r.Fz * r.Y) / 1000.0 - grouped.Sum(r => r.Mx);

            string outputPath = Path.Combine(folderPath, outputFilename);
            if (!outputPath.ToLower().EndsWith(".xlsx"))
                outputPath += ".xlsx";

            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");

                ws.Cell(1, 1).Value = "X";
                ws.Cell(1, 2).Value = "Y";
                ws.Cell(1, 3).Value = "Z";
                ws.Cell(1, 4).Value = "Fx";
                ws.Cell(1, 5).Value = "Fy";
                ws.Cell(1, 6).Value = "Fz";
                ws.Cell(1, 7).Value = "Mx";
                ws.Cell(1, 8).Value = "My";
                ws.Cell(1, 9).Value = "Mz";

                int row = 2;
                foreach (var r in grouped)
                {
                    ws.Cell(row, 1).Value = r.X;
                    ws.Cell(row, 2).Value = r.Y;
                    ws.Cell(row, 3).Value = r.Z;
                    ws.Cell(row, 4).Value = r.Fx;
                    ws.Cell(row, 5).Value = r.Fy;
                    ws.Cell(row, 6).Value = r.Fz;
                    ws.Cell(row, 7).Value = r.Mx;
                    ws.Cell(row, 8).Value = r.My;
                    ws.Cell(row, 9).Value = r.Mz;
                    row++;
                }

                wb.SaveAs(outputPath);
            }

            return (FxTotal, FyTotal, FzTotal, MXTotal, MYTotal, outputPath);
        }

        // =============================
        // READ EXCEL TO DATATABLE
        // =============================
        private DataTable ReadExcelToDataTable(string path)
        {
            var dt = new DataTable();

            using (var workbook = new XLWorkbook(path))
            {
                var ws = workbook.Worksheet(1);
                var range = ws.RangeUsed();
                int rows = range.RowCount();
                int cols = range.ColumnCount();

                for (int c = 1; c <= cols; c++)
                    dt.Columns.Add($"Col{c}", typeof(object));

                for (int r = 1; r <= rows; r++)
                {
                    var row = dt.NewRow();
                    for (int c = 1; c <= cols; c++)
                        row[c - 1] = range.Cell(r, c).Value;
                    dt.Rows.Add(row);
                }
            }

            return dt;
        }

        // =============================
        // MAP COORDINATES
        // =============================
        private List<ForcePoint> AddCoords(DataTable df, DataTable coorDf)
        {
            var list = new List<ForcePoint>();

            foreach (DataRow row in df.Rows)
            {
                var p = new ForcePoint
                {
                    Fx = ToFloatSafe(row.ItemArray.Length > 7 ? row[7] : null) ?? 0.0,
                    Fy = ToFloatSafe(row.ItemArray.Length > 8 ? row[8] : null) ?? 0.0,
                    Fz = ToFloatSafe(row.ItemArray.Length > 9 ? row[9] : null) ?? 0.0,
                    Mx = ToFloatSafe(row.ItemArray.Length > 10 ? row[10] : null) ?? 0.0,
                    My = ToFloatSafe(row.ItemArray.Length > 11 ? row[11] : null) ?? 0.0,
                    Mz = ToFloatSafe(row.ItemArray.Length > 12 ? row[12] : null) ?? 0.0
                };

                string joint = row.ItemArray.Length > 4 ? row[4]?.ToString() : null;
                if (string.IsNullOrEmpty(joint))
                {
                    list.Add(p);
                    continue;
                }

                p.Joint = joint;

                var matches = coorDf.AsEnumerable()
                    .Where(r => r.ItemArray.Length > 1 &&
                                (r[1]?.ToString() ?? "") == joint)
                    .ToList();

                if (matches.Any())
                {
                    var r0 = matches[0];
                    p.X = ToFloatSafe(r0.ItemArray.Length > 5 ? r0[5] : null);
                    p.Y = ToFloatSafe(r0.ItemArray.Length > 6 ? r0[6] : null);
                    p.Z = ToFloatSafe(r0.ItemArray.Length > 7 ? r0[7] : null);
                }

                list.Add(p);
            }

            return list;
        }

        // =============================
        // SAFE NUMERIC PARSER
        // =============================
        private double? ToFloatSafe(object value)
        {
            try
            {
                if (value == null || value is DBNull) return null;

                string s = value.ToString().ToLower();
                if (string.IsNullOrWhiteSpace(s)) return null;

                s = s.Replace("~", "").Replace("mm", "");
                s = s.Replace(",", "").Replace(" ", "").Trim();

                if (double.TryParse(s,
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out double result))
                {
                    return result;
                }

                return null;
            }
            catch
            {
                return null;
            }
        }

        // =============================
        // RUN BUTTON LOGIC
        // =============================
        private async void btnRun_Click(object sender, EventArgs e)
        {
            btnRun.Enabled = false;
            btnRun.Text = "Processing...";
            spinnerRunning = true;
            spinnerIndex = 0;
            spinnerTimer.Start();

            try
            {
                string folder = txtFolder.Text.Trim();
                string colFile = txtColFile.Text.Trim();
                string wallFile = txtWallFile.Text.Trim();
                string coorFile = txtCoorFile.Text.Trim();
                string outputFile = txtOutputFile.Text.Trim();
                string zText = txtZValue.Text.Trim();

                if (string.IsNullOrWhiteSpace(folder) || string.IsNullOrWhiteSpace(colFile) || string.IsNullOrWhiteSpace(wallFile) ||
                    string.IsNullOrWhiteSpace(coorFile) || string.IsNullOrWhiteSpace(outputFile) || string.IsNullOrWhiteSpace(zText))
                {
                    MessageBox.Show("Please fill in all required fields.", "Missing Data", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (!double.TryParse(zText, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double zValue))
                {
                    MessageBox.Show("Invalid Z elevation value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Process the data in a background thread
                var result = await Task.Run(() =>
                    ProcessData(folder, colFile, wallFile, coorFile, zValue, outputFile));

                string text =
                    $"Total Fx = {result.Fx:F2}{Environment.NewLine}" +
                    $"Total Fy = {result.Fy:F2}{Environment.NewLine}" +
                    $"Total Fz = {result.Fz:F2}{Environment.NewLine}" +
                    $"Moment about X = {result.Mx:F2}{Environment.NewLine}" +
                    $"Moment about Y = {result.My:F2}{Environment.NewLine}" +
                    $"{Environment.NewLine}" +
                    $"Output saved to:{Environment.NewLine}{result.OutputPath}";

                txtOutput.ReadOnly = false;
                txtOutput.Text = text;
                txtOutput.ReadOnly = true;

                MessageBox.Show("Process completed and Excel exported.", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                spinnerRunning = false;
                btnRun.Enabled = true;
                btnRun.Text = "Run";
            }
        }
    }
}
