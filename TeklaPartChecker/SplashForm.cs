using System;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Drawing.Text;


namespace TeklaPartChecker
{
    public partial class SplashForm : Form
    {
        private Timer progressTimer;
        private int progressValue = 0;
        private Panel progressContainer;
        private Panel progressFill;

        public SplashForm()
        {
            InitializeComponent();
            SetupUI();
        }

        private Font LoadCustomFont(string resourceName, float size)
        {
            PrivateFontCollection fonts = new PrivateFontCollection();
            var assembly = Assembly.GetExecutingAssembly();

            using (Stream fontStream = assembly.GetManifestResourceStream(resourceName))
            {
                if (fontStream == null)
                    throw new Exception("Font not found: " + resourceName);

                byte[] fontData = new byte[fontStream.Length];
                fontStream.Read(fontData, 0, fontData.Length);

                IntPtr fontPtr = Marshal.AllocCoTaskMem(fontData.Length);
                Marshal.Copy(fontData, 0, fontPtr, fontData.Length);
                fonts.AddMemoryFont(fontPtr, fontData.Length);
                Marshal.FreeCoTaskMem(fontPtr);
            }

            return new Font(fonts.Families[0], size, FontStyle.Regular);
        }

        private void SetupUI()
        {
            this.FormBorderStyle = FormBorderStyle.None;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.Black; // Pure black
            this.ClientSize = new Size(500, 250);

            Label title = new Label
            {
                Text = "🔧 Tekla Part Checker",
                Dock = DockStyle.Top,
                Height = 150,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = LoadCustomFont("TeklaPartChecker.Resources.Fonts.FiraMonoNerdFont-Regular.otf", 20f),
                ForeColor = Color.White
            };
            Controls.Add(title);

            // Outer progress bar container
            progressContainer = new Panel
            {
                Height = 20,
                Width = 400,
                Top = 180,
                Left = 50,
                BackColor = Color.FromArgb(50, 50, 50)
            };
            Controls.Add(progressContainer);

            // Orange fill panel (acts like progress bar)
            progressFill = new Panel
            {
                Height = 20,
                Width = 0,
                BackColor = Color.Orange
            };
            progressContainer.Controls.Add(progressFill);

            // Timer to simulate loading over 10s
            progressTimer = new Timer();
            progressTimer.Interval = 100; // ms
            progressTimer.Tick += ProgressTimer_Tick;
            progressTimer.Start();
        }

        private void ProgressTimer_Tick(object sender, EventArgs e)
        {
            progressValue++;

            int maxWidth = progressContainer.Width;
            progressFill.Width = (int)((progressValue / 100.0) * maxWidth);

            if (progressValue >= 100)
            {
                progressTimer.Stop();
                this.Close();
            }
        }

        private void SplashForm_Load(object sender, EventArgs e)
        {
            // Set the form to be topmost and transparent
            this.TopMost = true;
            this.TransparencyKey = this.BackColor; // Makes the background color transparent
            this.Opacity = 0.95; // Slightly transparent for a modern look
            // Center the form on the screen
            this.StartPosition = FormStartPosition.CenterScreen;
            // Load custom font if needed
            try
            {
                LoadCustomFont("TeklaPartChecker.Resources.Fonts.FiraMonoNerdFont-Regular.otf", 20f);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading custom font: " + ex.Message);
            }
        }
    }
}
