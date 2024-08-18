using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Data;
// using Microsoft.Extensions.Configuration;

namespace launcher
{
    public partial class Form1 : Form
    {
        private Dictionary<string, string> aliases = new Dictionary<string, string>();
        private float windowTransparency = 0.8f;
        private int maxLines = 10;
        private int errorDisplayDuration = 5000;
        private int displayDuration = 10000;
        private int fadeOutInterval = 50;

        public Form1()
        {
            InitializeComponent();
            InitializeCustomComponents();
            LoadSettings();
        }

        private void InitializeCustomComponents()
        {
            TextBox commandTextBox = new TextBox
            {
                Location = new System.Drawing.Point(10, 10),
                Width = 400
            };
            commandTextBox.KeyDown += new KeyEventHandler(OnKeyDownHandler);
            this.Controls.Add(commandTextBox);

            this.Text = "launcher";
            this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ClientSize = new System.Drawing.Size(420, 40);
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
                    var match = Regex.Match(line, @"(\w+)=[""]?(.*?)[""]?$");
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
            string iniFilePath = "launcher.ini";
            bool inAliasesSection = false;

            if (File.Exists(iniFilePath))
            {
                foreach (var line in File.ReadLines(iniFilePath))
                {
                    if (line.Trim().Equals("[Aliases]", StringComparison.OrdinalIgnoreCase))
                    {
                        inAliasesSection = true;
                        continue;
                    }

                    if (inAliasesSection)
                    {
                        if (line.StartsWith("[") && line.EndsWith("]"))
                        {
                            break; // End of [Aliases] section
                        }

                        var match = Regex.Match(line, @"(\w+)=[""]?(.*?)[""]?$");
                        if (match.Success)
                        {
                            aliases[match.Groups[1].Value] = match.Groups[2].Value;
                        }
                        else
                        {
                            ShowOutput("エイリアス読み込みエラー", $"エイリアスの形式が不正です: {line}", displayDuration:5000, isError: true);
                        }
                    }
                }
            }
        }

        private void ExecuteCommand(string command)
        {
            string output;
            string error;
            bool isError = false;

            if (aliases.TryGetValue(command, out string? alias))
            {
                command = alias;
            }

            if (IsMathExpression(command))
            {
                try
                {
                    double result = EvaluateMathExpression(command);
                    output = result.ToString();
                }
                catch (Exception ex)
                {
                    output = $"計算エラー: {ex.Message}";
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
                            output = error;
                            isError = true;
                        }
                        else if (string.IsNullOrWhiteSpace(output))
                        {
                            output = "(コマンドの実行には結果がありません)";
                        }
                    }
                }
                catch (Exception ex)
                {
                    output = $"エラー: {ex.Message}";
                    isError = true;
                }
            }

            ShowOutput(command, output, displayDuration, isError);
        }

        private bool IsMathExpression(string input)
        {
            return Regex.IsMatch(input, @"^\s*[-+]?(\d+(\.\d+)?|\.\d+)([-+*/]\s*[-+]?(\d+(\.\d+)?|\.\d+))*\s*$");
        }

        private double EvaluateMathExpression(string expression)
        {
            DataTable table = new DataTable();
            return Convert.ToDouble(table.Compute(expression, string.Empty));
        }

        private void ShowOutput(string command, string output, int displayDuration, bool isError)
        {
            OutputForm outputForm = new OutputForm(
                command,
                output,
                windowTransparency,
                displayDuration,
                fadeOutInterval,
                isError,
                maxLines
            );

            outputForm.Show();
            outputForm.TopMost = true;
            outputForm.Focus(); // ウインドウを最前面に表示

            // フォーカスがない場合は80%の透明度に設定
            if (!outputForm.Focused)
            {
                outputForm.Opacity = windowTransparency / 100.0;
            }

            // ウインドウがアクティブになったときの処理
            outputForm.Activated += (s, e) =>
            {
                outputForm.Opacity = 1.0; // 透明度を100%に設定
                outputForm.StopFadeOut(); // フェードアウトを停止
            };

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
