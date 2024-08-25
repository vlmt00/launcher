using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Data;
using System.Timers;
using System.Runtime.InteropServices; //HotKey
using Microsoft.Extensions.Configuration; //IniFile
// Require Extension Install : dotnet add package Microsoft.Extensions.Configuration.Ini
using xFunc.Maths;
//dotnet add package xFunc.Maths --version 4.4.1

namespace launcher
{
    public partial class MainForm : Form
    {
        private Dictionary<string, string> aliases = new Dictionary<string, string>();
        private float windowTransparency = 0.8f;
        private int maxLines = 10;
        private int errorDisplayDuration = 5000;
        private int displayDuration = 10000;
        private int fadeOutInterval = 50;
        private System.Timers.Timer MainTimer = new System.Timers.Timer();
        private TextBox commandTextBox = new TextBox();

        // ホットキーの識別子
        private const int HOTKEY_ID = 1;

        // Windows APIからRegisterHotKey関数をインポート
        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        // Windows APIからUnregisterHotKey関数をインポート
        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        // 修飾キー
        private const uint MOD_ALT = 0x0001;
        private const uint MOD_CONTROL = 0x0002;
        private const uint MOD_SHIFT = 0x0004;
        private const uint MOD_WIN = 0x0008;

        // 仮想キーコード
        private const uint VK_F1 = 0x70;
        private OutputForm outputForm = new OutputForm();

        public MainForm()
        {
            InitializeComponent();
            InitializeCustomComponents();
            LoadSettings();
        }

        private void InitializeCustomComponents()
        {
            commandTextBox.Location = new System.Drawing.Point(0, 0);
            commandTextBox.Width = 200;
            commandTextBox.KeyDown += new KeyEventHandler(OnKeyDownHandler);
            this.Controls.Add(commandTextBox);

            this.Text = "launcher";
            this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ClientSize = new System.Drawing.Size(200, 22);

            RegisterHotKey(this.Handle, HOTKEY_ID, MOD_CONTROL | MOD_ALT, VK_F1); // Ctrl + Alt + F1でホットキー登録

            this.MainTimer.Interval = 100;

            this.MainTimer.Elapsed += OnTimedEvent;
            this.MainTimer.AutoReset = true;
            this.MainTimer.Enabled = true;
            // イベントハンドラーの登録
            this.Activated += new EventHandler(MainForm_Activated);
            this.Deactivate += new EventHandler(MainForm_Deactivate);

        }

        // フォームがアクティブになった時の処理
        private void MainForm_Activated(object? sender, EventArgs e)
        {
            // アクティブ化時の処理
            this.Opacity = 1.0; // 透明度を100%にする
        }

        // フォームが非アクティブになった時の処理
        private void MainForm_Deactivate(object? sender, EventArgs e)
        {
            // 非アクティブ化時の処理
            this.Opacity = 0.8; // 透明度を80%にする
        }

        protected override void WndProc(ref Message m)
        {
            const int WM_HOTKEY = 0x0312;

            if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID)
            {
                this.Activate();
                this.commandTextBox.Select();
            }

            base.WndProc(ref m);
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            this.Activated -= MainForm_Activated;
            this.Deactivate -= MainForm_Deactivate;
            // フォームが閉じられたときにホットキーを解除
            UnregisterHotKey(this.Handle, HOTKEY_ID);
            base.OnFormClosed(e);
        }

        private void OnTimedEvent(Object? source, System.Timers.ElapsedEventArgs e)
        {
            DateTime dt = DateTime.Now;
            this.Text = dt.ToString("yyyy/MM/dd(ddd) HH:mm:ss");
        }

        private void OnKeyDownHandler(object? sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                // sender が TextBox であることを確認する
                if (sender is TextBox textBox)
                {
                    string command = textBox.Text;
                    LoadAliases(); // 毎回エイリアスを読み込む
                    ExecuteCommand(command);
                    textBox.Clear();
                }
            }
        }

        private void LoadSettings()
        {
            string iniFilePath = "launcher.ini";
            if (File.Exists(iniFilePath))
            {
                foreach (var line in File.ReadLines(iniFilePath))
                {
                    var match = Regex.Match(line, @"^\s*\(?(\d+(\.\d+)?|\(\d+(\.\d+)?\))\s*([-+*/]\s*\(?(\d+(\.\d+)?|\(\d+(\.\d+)?\))\s*\)?)*\s*$");//@"(\w+)=[""]?(.*?)[""]?$");
                    if (match.Success)
                    {
                        switch (match.Groups[1].Value)
                        {
                            case "transparency":
                                if (float.TryParse(match.Groups[2].Value, out float transparency))
                                    windowTransparency = transparency;
                                break;
                            case "maxLines":
                                if (int.TryParse(match.Groups[2].Value, out int lines))
                                    maxLines = lines;
                                break;
                            case "DisplayDuration":
                                if (int.TryParse(match.Groups[2].Value, out int duration))
                                    displayDuration = duration;
                                break;
                            case "errorDisplayDuration":
                                if (int.TryParse(match.Groups[2].Value, out int eduration))
                                    errorDisplayDuration = eduration;
                                break;
                            case "fadeOutInterval":
                                if (int.TryParse(match.Groups[2].Value, out int interval))
                                    fadeOutInterval = interval;
                                break;
                        }
                    }
                }
            }
        }

        private void LoadAliases()
        {
            aliases.Clear();
            string filename = "launcher.ini";
            string path = Directory.GetCurrentDirectory();
            string iniFilePath = path + "\\" +filename;
            bool inAliasesSection = false;
            if (!File.Exists(iniFilePath))
            {
                outputForm.AppendText("Ini File not found : " + iniFilePath, Color.White, Color.DarkRed);
            }else{
                IConfigurationBuilder builder = new ConfigurationBuilder()
                    .AddIniFile(iniFilePath);
                IConfiguration config = builder.Build();
                IConfigurationSection section;

                section = config.GetSection("Aliases");
                foreach (var child in section.GetChildren())
                {
                    if (null != child.Value)
                        aliases[child.Key] = child.Value;
                    //outputForm.AppendText($"{child.Key}: {child.Value}", Color.White, Color.Black);
                }
            }

        }

        private void ExecuteCommand(string command)
        {
            string output;
            string error;
            bool isOutput = false;
            bool isError = false;

            if (aliases.TryGetValue(command, out string? alias))
            {
                command = alias;
            }

            if (IsMathExpression(command))
            {
                isOutput = true;
                isError = false;
                try
                {
                    string result = EvaluateMathExpression(command);
                    output = result + "\r\n";
                }
                catch (Exception ex)
                {
                    output = "計算エラー:" + ex.Message;
                    isOutput = true;
                    isError = true;
                }
            }
            else if (Directory.Exists(command))
            {
                Process.Start("explorer.exe", command);
                return;
            }
            else
            {
                try
                {
                    using (Process process = new Process())
                    {
                        process.StartInfo.FileName = "cmd.exe";
                        process.StartInfo.Arguments = $"/C \"{command}\"";
                        process.StartInfo.RedirectStandardOutput = true;
                        process.StartInfo.RedirectStandardError = true;
                        process.StartInfo.UseShellExecute = false;
                        process.StartInfo.CreateNoWindow = true;

                        process.Start();
                        output = process.StandardOutput.ReadToEnd();
                        error = process.StandardError.ReadToEnd();
                        process.WaitForExit();

                        if (!string.IsNullOrWhiteSpace(error))
                        {
                            output =  "エラー: " + error;
                            isOutput = true;
                            isError = true;
                        }
                        else if (string.IsNullOrWhiteSpace(output))
                        {
                            isOutput = false;
                            //output = "(コマンドの実行には結果がありません)";
                        }else{
                            isOutput = true;
                            isError = false;
                        }
                    }
                }
                catch (Exception ex)
                {
                    output = $"エラー: {ex.Message}";
                    isOutput = true;
                    isError = true;
                }
            }
            if (isOutput)
            {
                //ShowOutput(command, output, displayDuration, isError);
                int lines = outputForm.AppendText(command + "\r\n", Color.LimeGreen, Color.Black);
                if(isError)
                {
                    lines = outputForm.AppendText(output, Color.White, Color.DarkRed);
                }else{
                    lines = outputForm.AppendText(output, Color.White, Color.Black);
                }

                commandTextBox.Text = lines.ToString();
            }
        }

        private bool IsMathExpression(string input)
        {
            return Regex.IsMatch(input, @"^\s*-?\(?(\d+(\.\d+)?|\(\d+(\.\d+)?\))\s*([+\-*/^]\s*-?\(?(\d+(\.\d+)?|\(\d+(\.\d+)?\))\s*\)?)*\s*$");
        }

        private string EvaluateMathExpression(string expression)
        {
            var processor = new Processor();
            var result = processor.Solve(expression);
            if (null != result)
            {
                return result.ToString();
            }else{
                return "solving error: " + result;
            }
            // DataTable table = new DataTable();
            // return Convert.ToDouble(table.Compute(expression, string.Empty));
        }

        private void ShowOutput(string command, string output, int displayDuration, bool isError)
        {
            // OutputForm outputForm = new OutputForm(
            //     command,
            //     output,
            //     windowTransparency,
            //     displayDuration,
            //     fadeOutInterval,
            //     isError,
            //     maxLines
            // );

            outputForm.Show();
            outputForm.TopMost = true;
            outputForm.Focus(); // ウインドウを最前面に表示

            // // フォーカスがない場合は80%の透明度に設定
            // if (!outputForm.Focused)
            // {
            //     outputForm.Opacity = windowTransparency / 100.0;
            // }

            // // ウインドウがアクティブになったときの処理
            // outputForm.Activated += (s, e) =>
            // {
            //     outputForm.Opacity = 1.0; // 透明度を100%に設定
            //     outputForm.StopFadeOut(); // フェードアウトを停止
            // };

            // 10秒後にウインドウが選択されていない場合は閉じる
            // var closeTimer = new System.Windows.Forms.Timer { Interval = displayDuration };
            // closeTimer.Tick += (s, e) =>
            // {
            //     if (!outputForm.Focused)
            //     {
            //         outputForm.Close();
            //         closeTimer.Stop();
            //         closeTimer.Dispose();
            //     }
            // };
            // closeTimer.Start();
        }
    }
}
