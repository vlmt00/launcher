using System;
using System.Diagnostics.Tracing;
using System.Drawing;
using System.Timers;
using System.Windows.Forms;

namespace launcher
{
    public partial class OutputForm : Form
    {

        private RichTextBox outputTextBox = new RichTextBox
        {
            Multiline = true,
            ReadOnly = true,
            ScrollBars = RichTextBoxScrollBars.Vertical,
            Dock = DockStyle.Fill,
            BackColor = Color.Black,
            WordWrap = false,
            HideSelection = false,
            Font = new Font("Arial", 10),
            ForeColor = Color.White
        };
        private System.Timers.Timer closeTimer = new System.Timers.Timer();
        private System.Timers.Timer fadeOutTimer = new System.Timers.Timer();
        private float transparency;
        private int fadeOutInterval;
        private bool isError;

        public OutputForm()
        {
            InitializeComponent();

            //this.FormBorderStyle = FormBorderStyle.FixedToolWindow;

            // this.BackColor = Color.Lime;  // ウィンドウ背景を特殊な色に設定
            // this.TransparencyKey = Color.Lime;

            this.Controls.Add(outputTextBox);
            this.Show();
        }

        public int AppendText(string text, Color textColor, Color bgColor)
        {
            // 現在のテキスト末尾にカーソルを移動
            outputTextBox.SelectionStart = outputTextBox.TextLength;
            outputTextBox.SelectionLength = 0;

            // テキストの色を設定
            outputTextBox.SelectionColor = textColor;
            outputTextBox.SelectionBackColor = bgColor;

            // テキストを追加
            outputTextBox.AppendText(text);

            // カーソルの色をデフォルトに戻す
            outputTextBox.SelectionColor = outputTextBox.ForeColor;
            outputTextBox.SelectionBackColor = outputTextBox.BackColor;

            // 現在のテキストを末尾にカーソルを移動
            outputTextBox.Select(outputTextBox.Text.Length, 0);
            outputTextBox.SelectionStart = outputTextBox.Text.Length;
            outputTextBox.SelectionLength = 0;
            outputTextBox.Focus();
            outputTextBox.ScrollToCaret();

            return outputTextBox.Lines.Length + 1;
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
            //this.SuspendLayout();
            //
            // OutputForm
            //
            this.ClientSize = new System.Drawing.Size(420, 261);
            this.Name = "OutputForm";
            //this.ResumeLayout(false);
            this.FormBorderStyle = FormBorderStyle.None;
        }
    }

}
