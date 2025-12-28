using System.Diagnostics;
using System.Reflection;

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

        private void InitializeComponent()
        {
            // --- ICON ---
            picIcon = new PictureBox();
            picIcon.Size = new Size(48, 48);
            picIcon.Location = new Point(12, 12);
            picIcon.SizeMode = PictureBoxSizeMode.StretchImage;

            try
            {
                // Ressourcen-Icon laden (z. B. wrok_black)
                picIcon.Image = Properties.Resources.wrok_black.ToBitmap();
            }
            catch
            {
                // Falls das Icon fehlt
                picIcon.Image = SystemIcons.Application.ToBitmap();
            }

            // --- Überschrift ---
            lblTitle = new Label();
            lblTitle.AutoSize = true;
            lblTitle.Font = new Font("Segoe UI", 11F, FontStyle.Bold);
            lblTitle.Location = new Point(72, 18);
            lblTitle.Text = "Wrok";

            // --- Version ---
            lblVersion = new Label();
            lblVersion.AutoSize = true;
            lblVersion.Font = new Font("Segoe UI", 9F);
            lblVersion.Location = new Point(72, 48);
            lblVersion.Text = "Version";

            // --- Github-Link ---
            linkGitHub = new LinkLabel();
            linkGitHub.AutoSize = true;
            linkGitHub.Font = new Font("Segoe UI", 9F);
            linkGitHub.Location = new Point(72, 78);
            linkGitHub.Text = "GitHub: manfred-mueller/Wrok";
            linkGitHub.TabStop = true;
            linkGitHub.LinkClicked += LinkGitHub_LinkClicked;

            // --- Schließen-Button ---
            btnClose = new Button();
            btnClose.Text = Properties.Resources.Close;
            btnClose.DialogResult = DialogResult.OK;
            btnClose.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnClose.Location = new Point(240, 118);
            btnClose.Size = new Size(90, 28);

            // --- Fenster ---
            this.AcceptButton = btnClose;
            this.CancelButton = btnClose;
            this.ClientSize = new Size(342, 160);
            this.Controls.Add(picIcon);
            this.Controls.Add(lblTitle);
            this.Controls.Add(lblVersion);
            this.Controls.Add(linkGitHub);
            this.Controls.Add(btnClose);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Text = Properties.Resources.AboutWrok;
        }

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
    }
}
