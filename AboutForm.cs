using System;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

namespace Wrok
{
    public partial class AboutForm : Form
    {
        public AboutForm()
        {
            InitializeComponent();
            InitContent();
        }

        private PictureBox picIcon = null!;
        private Label lblTitle = null!;
        private Label lblVersion = null!;
        private LinkLabel linkGitHub = null!;
        private Button btnClose = null!;
        private Button btnManual = null!;

        private void InitializeComponent()
        {
            // Icon picture box setup and resource fallback.
            picIcon = new PictureBox();
            picIcon.Size = new Size(48, 48);
            picIcon.Location = new Point(12, 12);
            picIcon.SizeMode = PictureBoxSizeMode.StretchImage;

            try
            {
                picIcon.Image = Properties.Resources.wrok_black.ToBitmap();
            }
            catch
            {
                // Fall back to a generic application icon if resource is missing.
                picIcon.Image = SystemIcons.Application.ToBitmap();
            }

            // Title label configuration.
            lblTitle = new Label();
            lblTitle.AutoSize = true;
            lblTitle.Font = new Font("Segoe UI", 11F, FontStyle.Bold);
            lblTitle.Location = new Point(72, 18);
            lblTitle.Text = "Wrok";

            // Version label configuration.
            lblVersion = new Label();
            lblVersion.AutoSize = true;
            lblVersion.Font = new Font("Segoe UI", 9F);
            lblVersion.Location = new Point(72, 48);
            lblVersion.Text = "Version";

            // GitHub link label configuration.
            linkGitHub = new LinkLabel();
            linkGitHub.AutoSize = true;
            linkGitHub.Font = new Font("Segoe UI", 9F);
            linkGitHub.Location = new Point(72, 78);
            linkGitHub.Text = "GitHub: manfred-mueller/Wrok";
            linkGitHub.TabStop = true;
            linkGitHub.LinkClicked += LinkGitHub_LinkClicked;

            // Manual button: opens a short manual explaining macros and cache clearing.
            btnManual = new Button();
            btnManual.Text = Wrok.Properties.Resources.Manual;
            btnManual.AutoSize = false;
            btnManual.Size = new Size(90, 28);
            btnManual.Location = new Point(138, 118);
            btnManual.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnManual.Click += BtnManual_Click;

            // Close button configuration — closes the about dialog.
            btnClose = new Button();
            btnClose.Text = Properties.Resources.Close;
            btnClose.DialogResult = DialogResult.OK;
            btnClose.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnClose.Location = new Point(240, 118);
            btnClose.Size = new Size(90, 28);

            // Finalize form.
            this.AcceptButton = btnClose;
            this.CancelButton = btnClose;
            this.ClientSize = new Size(342, 160);
            this.Controls.Add(picIcon);
            this.Controls.Add(lblTitle);
            this.Controls.Add(lblVersion);
            this.Controls.Add(linkGitHub);
            this.Controls.Add(btnManual);
            this.Controls.Add(btnClose);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Text = Properties.Resources.AboutWrok;
        }

        // Populate version and title from the executing assembly.
        private void InitContent()
        {
            try
            {
                var asm = Assembly.GetExecutingAssembly();
                var name = asm.GetName();
                var version = name.Version;

                lblTitle.Text = name.Name ?? "Wrok";

                if (version != null)
                {
                    lblVersion.Text = $"Version {version.Major}.{version.Minor}.{version.Build}";
                }
            }
            catch
            {
                lblVersion.Text = Properties.Resources.UnknownVersion;
            }
        }

        // Open project GitHub page using the default shell.
        private void LinkGitHub_LinkClicked(object? sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                var url = "https://github.com/manfred-mueller/Wrok";
                Process.Start(new ProcessStartInfo
                {
                    FileName = url,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                Trace.WriteLine(String.Format(Properties.Resources.CouldnTOpenLinkToGitHub0, ex));
                MessageBox.Show(
                    Properties.Resources.LinkCouldnTBeOepened,
                    "Wrok",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void BtnManual_Click(object? sender, EventArgs e)
        {
            try
            {
                ShowManualDialog();
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"ShowManualDialog failed: {ex}");
            }
        }

        // Show a modal dialog with a short manual describing macros and cache clearing behavior.
        private void ShowManualDialog()
        {
            using var dlg = new Form()
            {
                Text = Wrok.Properties.Resources.Manual,
                StartPosition = FormStartPosition.CenterParent,
                ClientSize = new Size(720, 430),
                FormBorderStyle = FormBorderStyle.SizableToolWindow,
                MinimizeBox = false,
                MaximizeBox = false
            };

            var tb = new TextBox()
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 9F),
                BackColor = SystemColors.Window,
                ForeColor = SystemColors.ControlText
            };

            tb.Text = GetManualText();

            // Create bottom panel that will contain the OK button.
            var panel = new Panel() { Dock = DockStyle.Bottom, Height = 44 };

            var btnOk = new Button()
            {
                Text = Properties.Resources.OK,
                DialogResult = DialogResult.OK,
                Size = new Size(90, 28),
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };

            // Position the button relative to the panel's client size and keep it right-aligned on resize.
            void PositionOkButton()
            {
                try
                {
                    btnOk.Location = new Point(Math.Max(8, panel.ClientSize.Width - btnOk.Width - 10), (panel.ClientSize.Height - btnOk.Height) / 2);
                }
                catch { /* defensive */ }
            }

            // Initial positioning and resize handler.
            panel.Resize += (s, e) => PositionOkButton();
            panel.Controls.Add(btnOk);

            // Add controls in correct z-order (text fills, panel at bottom).
            dlg.Controls.Add(tb);
            dlg.Controls.Add(panel);

            // Accept button and shown-focus behavior.
            dlg.AcceptButton = btnOk;
            dlg.Shown += (s, e) =>
            {
                try
                {
                    PositionOkButton();
                    btnOk.Focus();
                }
                catch { }
            };

            // Close dialog when OK clicked (DialogResult is already set).
            btnOk.Click += (s, e) => dlg.Close();

            dlg.ShowDialog(this);
        }

        private string GetManualText()
        {
            return Wrok.Properties.Resources.MacroHelp;
        }
    }
}
