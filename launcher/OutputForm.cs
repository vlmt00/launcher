using System;
using System.Drawing;
using System.Timers;
using System.Windows.Forms;

namespace launcher
{
    public partial class OutputForm : Form
    {
        private System.Timers.Timer closeTimer;
        private System.Timers.Timer fadeOutTimer;
        private float transparency;
        private int fadeOutInterval;
        private bool isError;

        public OutputForm(string command, string output, float transparency, int displayDuration, int fadeOutInterval, bool isError, int maxLines)
        {
            InitializeComponent();

            this.transparency = transparency;
            this.fadeOutInterval = fadeOutInterval;
            this.isError = isError;

            this.Text = command;
            this.Opacity = transparency;
            this.BackColor = isError ? Color.LightCoral : Color.White;
            this.FormBorderStyle = FormBorderStyle.FixedToolWindow;

            TextBox outputTextBox = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                BackColor = isError ? Color.LightCoral : Color.White
            };

            outputTextBox.Text = output;
            outputTextBox.SelectionStart = outputTextBox.Text.Length;
            outputTextBox.SelectionLength = 0;
            outputTextBox.ScrollToCaret();

            this.Controls.Add(outputTextBox);

            int lines = outputTextBox.GetLineFromCharIndex(outputTextBox.TextLength) + 1;
            this.ClientSize = new Size(this.ClientSize.Width, Math.Min(lines, maxLines) * outputTextBox.Font.Height + 10);

            // closeTimer の初期化
            closeTimer = new System.Timers.Timer(displayDuration);
            closeTimer.Elapsed += CloseTimerElapsed;
            fadeOutTimer = new System.Timers.Timer(fadeOutInterval);
            fadeOutTimer.Elapsed += FadeOutTimerElapsed;

            // fadeOutTimer の初期化
            fadeOutTimer = new System.Timers.Timer(fadeOutInterval);
            fadeOutTimer.Elapsed += FadeOutTimerElapsed;
        }

        public void StopFadeOut()
        {
            if (fadeOutTimer.Enabled)
            {
                fadeOutTimer.Stop();
            }
        }

        private void CloseTimerElapsed(object? sender, ElapsedEventArgs e)
        {
            fadeOutTimer.Start();
        }

        private void FadeOutTimerElapsed(object? sender, ElapsedEventArgs e)
        {
            if (this.Opacity > 0)
            {
                this.Invoke((MethodInvoker)delegate
                {
                    this.Opacity -= 0.05;
                });
            }
            else
            {
                fadeOutTimer.Stop();
                this.Invoke((MethodInvoker)delegate
                {
                    this.Close();
                });
            }
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            //
            // OutputForm
            //
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Name = "OutputForm";
            this.ResumeLayout(false);
        }
    }
}
