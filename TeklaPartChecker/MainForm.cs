using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Tekla.Structures.Drawing;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Model;
using Tekla.Structures.Solid;
using WinFormsPoint = System.Drawing.Point;
using WinFormsSize = System.Drawing.Size;


namespace TeklaPartChecker
{
    public class MainForm : Form
    {

        private bool AreBoundingBoxesIntersecting(Solid solidA, Solid solidB)
        {
            var minA = solidA.MinimumPoint;
            var maxA = solidA.MaximumPoint;

            var minB = solidB.MinimumPoint;
            var maxB = solidB.MaximumPoint;

            return (minA.X <= maxB.X && maxA.X >= minB.X) &&
                   (minA.Y <= maxB.Y && maxA.Y >= minB.Y) &&
                   (minA.Z <= maxB.Z && maxA.Z >= minB.Z);
        }






        private Button runButton;

        public MainForm()
        {

            this.MinimumSize = new System.Drawing.Size(500, 500);
            this.FormBorderStyle = FormBorderStyle.Sizable;


            Text = "Tekla Part Checker";
            Size = new System.Drawing.Size(500, 500);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            BackColor = System.Drawing.Color.FromArgb(30, 30, 30); // dark gray
            ForeColor = System.Drawing.Color.White;
            Font = new System.Drawing.Font("Segoe UI", 9F);

            runButton = new Button();
            runButton.Text = "📋 Run Check";
            runButton.Width = 240;
            runButton.Height = 50;
            runButton.Top = 40;
            runButton.Left = (this.ClientSize.Width - runButton.Width) / 2;
            runButton.Click += RunScript;

            // Dark style for button
            runButton.FlatStyle = FlatStyle.Flat;
            runButton.BackColor = System.Drawing.Color.FromArgb(50, 50, 50);
            runButton.ForeColor = System.Drawing.Color.White;
            runButton.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;

            Controls.Add(runButton);

            //Refresh button
            Button refreshButton = new Button();
            refreshButton.Text = "🔄 Refresh";
            refreshButton.Width = 240;
            refreshButton.Height = 40;
            refreshButton.Top = 100;
            refreshButton.Left = (this.ClientSize.Width - refreshButton.Width) / 2;
            refreshButton.Click += RunScript;

            Controls.Add(refreshButton);

            // About button (top-right corner)
            Button aboutButton = new Button();
            aboutButton.Text = "ℹ️";
            aboutButton.Width = 30;
            aboutButton.Height = 30;
            aboutButton.Top = 10;
            aboutButton.Left = this.ClientSize.Width - aboutButton.Width - 10; // right align
            aboutButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            aboutButton.Click += ShowAboutDialog;
            aboutButton.FlatStyle = FlatStyle.Flat;
            aboutButton.BackColor = System.Drawing.Color.FromArgb(50, 50, 50);
            aboutButton.ForeColor = System.Drawing.Color.White;
            aboutButton.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;

            Controls.Add(aboutButton);

            // Drawing check button
            Button drawingAuditButton = new Button();
            drawingAuditButton.Text = "🖊️ Check Drawings";
            drawingAuditButton.Width = 240;
            drawingAuditButton.Height = 40;
            drawingAuditButton.Top = 150;
            drawingAuditButton.Left = (this.ClientSize.Width - drawingAuditButton.Width) / 2;
            drawingAuditButton.Click += CheckDrawingsForEmptyDrawnBy;

            drawingAuditButton.FlatStyle = FlatStyle.Flat;
            drawingAuditButton.BackColor = Color.FromArgb(50, 50, 50);
            drawingAuditButton.ForeColor = Color.White;
            drawingAuditButton.FlatAppearance.BorderColor = Color.DimGray;

            Controls.Add(drawingAuditButton);




        }


        private void RunScript(object sender, EventArgs e)
        {
            var log = new List<string>();

            // Create a temporary working directory
            string tempDir = Path.Combine(Path.GetTempPath(), "TeklaExport_" + Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempDir);

            string csvPath = Path.Combine(tempDir, "tekla_parts_report.csv");
            string summaryPath = Path.Combine(tempDir, "tekla_parts_report.summary.txt");
            string logPath = Path.Combine(tempDir, "tekla_gui_log.txt");

            // Ask user where to save the final ZIP file
            string zipPath;
            using (SaveFileDialog dialog = new SaveFileDialog())
            {
                dialog.Title = "Save ZIP Report Bundle";
                dialog.Filter = "ZIP files (*.zip)|*.zip";
                dialog.FileName = "tekla_report_bundle.zip";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    zipPath = dialog.FileName;
                }
                else
                {
                    return;
                }
            }

            try
            {
                log.Add("========================================");
                log.Add($"[{DateTime.Now}] Starting Tekla Part Check...");

                var model = new Model();
                int attempts = 0;

                while (!model.GetConnectionStatus() && attempts < 10)
                {
                    Thread.Sleep(500);
                    model = new Model();
                    attempts++;
                }

                if (!model.GetConnectionStatus())
                    throw new Exception("Could not connect to Tekla Structures model.");

                var modelInfo = model.GetInfo();
                log.Add($"Connected to model: {modelInfo.ModelName}");

                var selector = model.GetModelObjectSelector();
                var enumerator = selector.GetAllObjects();

                var parts = new List<Tekla.Structures.Model.Part>();
while (enumerator.MoveNext())
{
    if (enumerator.Current is Tekla.Structures.Model.Part part)
    {
        // Only add main parts, skip secondary/child parts
        var assembly = part.GetAssembly();
        if (assembly != null && assembly.GetMainPart() is Tekla.Structures.Model.Part mainPart && mainPart.Identifier.ID == part.Identifier.ID)
        {
            parts.Add(part);
        }
    }
}

                log.Add($"Found {parts.Count} parts.");

                int clashCount = 0;
                var clashingPairs = new List<string>();

                for (int i = 0; i < parts.Count; i++)
                {
                    var partA = parts[i];
                    Solid solidA = partA.GetSolid();

                    for (int j = i + 1; j < parts.Count; j++)
                    {
                        var partB = parts[j];

                        // Skip if same part or not valid
                        if (partA.Identifier.ID == partB.Identifier.ID)
                            continue;

                        Solid solidB = partB.GetSolid();

                        // Broad-phase: bounding box test first
                        if (!AreBoundingBoxesIntersecting(solidA, solidB))
                            continue;

                        try
                        {
                            // var intersection = solidA.Intersect(solidB);
                            // if (intersection != null && intersection.Volume() > 0)
                            // {
                            //     clashCount++;
                            //     clashingPairs.Add($"⚠️ {partA.Name} intersects {partB.Name}");
                            // }

                            // Use bounding box intersection only, or implement a more advanced check if needed
                            if (AreBoundingBoxesIntersecting(solidA, solidB))
                            {
                                clashCount++;
                                clashingPairs.Add($"⚠️ {partA.Name} intersects {partB.Name}");
                            }
                        }
                        catch
                        {
                            // Ignore intersection errors, e.g. if solids are complex or invalid
                        }
                    }
                }


                log.Add($"Detected {clashCount} clashing part pairs.");
                log.AddRange(clashingPairs);




                // Write CSV
                using (var writer = new StreamWriter(csvPath, false, Encoding.UTF8))
                {
                    writer.WriteLine("Part Name;Profile;Status");
                    foreach (var part in parts)
                    {
                        string name = part.Name;
                        string profile = part.Profile?.ProfileString ?? "Unknown";
                        string status = string.IsNullOrWhiteSpace(name) ? "Empty Name" : "OK";
                        writer.WriteLine($"{name};{profile};{status}");
                    }
                }

                int emptyNameCount = parts.Count(p => string.IsNullOrWhiteSpace(p.Name));
                int okCount = parts.Count - emptyNameCount;

                // Count missing "Drawn By" in drawings
                int missingDrawnByCount = 0;
                try
                {
                    var drawingHandler = new Tekla.Structures.Drawing.DrawingHandler();
                    if (drawingHandler.GetConnectionStatus())
                    {
                        var drawingsEnum = drawingHandler.GetDrawings();
                        while (drawingsEnum.MoveNext())
                        {
                            var drawing = drawingsEnum.Current as Tekla.Structures.Drawing.Drawing;
                            if (drawing == null) continue;
                            string drawnBy = "";
                            drawing.GetUserProperty("DRAWN_BY", ref drawnBy);
                            if (string.IsNullOrWhiteSpace(drawnBy))
                                missingDrawnByCount++;
                        }
                    }
                }
                catch { /* Ignore drawing errors for summary */ }

                var profileGroups = parts
                    .GroupBy(p => p.Profile?.ProfileString ?? "Unknown")
                    .Select(g => $"{g.Key}: {g.Count()} parts");

                string summary = $"📋 Part Summary:\n" +
                                 $"• Total parts: {parts.Count}\n" +
                                 $"• OK: {okCount}\n" +
                                 $"• Empty names: {emptyNameCount}\n" +
                                 $"• Drawings missing 'Drawn By': {missingDrawnByCount}\n\n" +
                                 $"🔢 Profiles:\n{string.Join("\n", profileGroups)}";

                summary += $"\n🛑 Clashing Parts: {clashCount}";
                if (clashingPairs.Count > 0)
                {
                    summary += $"\n\nIntersections:\n" + string.Join("\n", clashingPairs.Take(10));
                    if (clashingPairs.Count > 10)
                        summary += $"\n...and {clashingPairs.Count - 10} more.";
                }

                File.WriteAllText(summaryPath, summary);
                log.Add($"Saved part summary to: {summaryPath}");

                MessageBox.Show(summary, "Part Summary", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Write log before zipping
                File.WriteAllLines(logPath, log);

                // Create ZIP
                if (File.Exists(zipPath)) File.Delete(zipPath);

                string modelFolderName = modelInfo.ModelName?.Trim().Replace(" ", "_") ?? "Model";
                using (FileStream zipToOpen = new FileStream(zipPath, FileMode.Create))
                using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
                {
                    archive.CreateEntryFromFile(csvPath, $"{modelFolderName}/{Path.GetFileName(csvPath)}");
                    archive.CreateEntryFromFile(summaryPath, $"{modelFolderName}/{Path.GetFileName(summaryPath)}");
                    archive.CreateEntryFromFile(logPath, $"{modelFolderName}/{Path.GetFileName(logPath)}");
                }


                log.Add($"📦 Created ZIP bundle: {zipPath}");

                // Clean up temp files
                Directory.Delete(tempDir, true);

                MessageBox.Show($"✅ Exported {parts.Count} parts.\n\n📦 ZIP saved at:\n{zipPath}", "Success");
            }
            catch (Exception ex)
            {
                log.Add("ERROR:");
                log.Add(ex.Message);
                log.Add(ex.StackTrace);
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                try
                {
                    File.AppendAllLines(logPath, log);
                }
                catch { }
            }
        }



        private void ShowAboutDialog(object sender, EventArgs e)
        {
            string message = "Tekla Part Checker\n" +
                             "Version 1.2.5\n" +
                             "Created by HoldFort S.A.";
            MessageBox.Show(message, "About", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void AddVersionLabel()
        {
            Label versionLabel = new Label
            {
                Text = "v1.2.5",
                AutoSize = true,
                ForeColor = Color.Gray,
                BackColor = Color.Transparent,
                Font = new Font("Segoe UI", 8, FontStyle.Regular),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };

            this.Controls.Add(versionLabel);

            // Force layout so we can get the size
            versionLabel.PerformLayout();
            versionLabel.Refresh();

            // Set position manually using measured size
            int x = this.ClientSize.Width - versionLabel.PreferredWidth - 10;
            int y = this.ClientSize.Height - versionLabel.PreferredHeight - 10;
            versionLabel.Location = new System.Drawing.Point(x, y);
        }




        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // MainForm
            // 
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Name = "MainForm";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.ResumeLayout(false);

        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            AddVersionLabel();
        }


        private void CheckDrawingsForEmptyDrawnBy(object sender, EventArgs e)
        {
            try
            {
                var drawingHandler = new Tekla.Structures.Drawing.DrawingHandler();

                if (!drawingHandler.GetConnectionStatus())
                {
                    MessageBox.Show("Could not connect to Tekla drawing handler.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var drawingsEnum = drawingHandler.GetDrawings();
                var missingDrawnByList = new List<string>();
                int totalDrawings = 0;

                while (drawingsEnum.MoveNext())
                {
                    totalDrawings++;

                    var drawing = drawingsEnum.Current as Tekla.Structures.Drawing.Drawing;
                    if (drawing == null) continue;

                    string drawnBy = "";
                    drawing.GetUserProperty("DRAWN_BY", ref drawnBy);

                    if (string.IsNullOrWhiteSpace(drawnBy))
                    {
                        string drawingName = drawing.Name;
                        string type = drawing.GetType().Name;
                        missingDrawnByList.Add($"{drawingName} ({type})");
                    }
                }

                StringBuilder sb = new StringBuilder();
                sb.AppendLine($"🖊️ Drawing Metadata Check");
                sb.AppendLine($"• Total drawings: {totalDrawings}");
                sb.AppendLine($"• Missing 'Drawn By': {missingDrawnByList.Count}");

                if (missingDrawnByList.Count > 0)
                {
                    sb.AppendLine($"\n⚠️ Drawings missing 'Drawn By':");
                    foreach (var d in missingDrawnByList)
                    {
                        sb.AppendLine($"- {d}");
                    }
                }

                MessageBox.Show(sb.ToString(), "Drawing Check", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Optional: Export to file
                string tempPath = Path.Combine(Path.GetTempPath(), $"drawing_check_{DateTime.Now:yyyyMMdd_HHmmss}");
                Directory.CreateDirectory(tempPath);
                File.WriteAllLines(Path.Combine(tempPath, "missing_drawn_by.csv"), missingDrawnByList);
                File.WriteAllText(Path.Combine(tempPath, "summary.txt"), sb.ToString());

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }



    }
}

