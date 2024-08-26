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
using System.Security;
//dotnet add package xFunc.Maths --version 4.4.1

namespace launcher
{
    public partial class MainForm : Form
    {
        private Dictionary<string, string> aliases = new Dictionary<string, string>();
        private float windowTransparency = 0.8f;
        private int maxLines = 5;
        private int errorDisplayDuration = 5000;
        private int displayDuration = 10000;
        private int fadeOutInterval = 50;
        private System.Timers.Timer MainTimer = new System.Timers.Timer();
        private TextBox commandTextBox = new TextBox();

        // ホットキーの識別子
        private const int HOTKEY_ID = 1;
        private const uint MOD_NOREPEAT = 0x4000;

        // Windows APIからRegisterHotKey関数をインポート
        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        // Windows APIからUnregisterHotKey関数をインポート
        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

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
            // Timer
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
            string filename = "launcher.ini";
            string path = Directory.GetCurrentDirectory();
            string iniFilePath = path + "\\" +filename;

            if (!File.Exists(iniFilePath))
            {
                outputForm.AppendText("Ini File not found : " + iniFilePath, Color.White, Color.DarkRed);
            }else{
                IConfigurationBuilder builder = new ConfigurationBuilder()
                    .AddIniFile(iniFilePath);
                IConfiguration config = builder.Build();
                IConfigurationSection section;

                section = config.GetSection("Config");
                foreach (var child in section.GetChildren())
                {
                    if(null != child.Key && null != child.Value)
                    {
                        switch (child.Key.ToString())
                        {
                            case "MaxLines":
                                if(int.TryParse(child.Value, out int lines))
                                    this.maxLines = lines;
                                break;
                            case "HotKey":
                                uint mod_code = MOD_NOREPEAT;
                                uint key_code = 0x00;
                                string[] words = child.Value.Split(' ', '　', ',', '|');
                                foreach (var word in words)
                                {
                                    (uint mod, uint vk) = conv_key_char(word);
                                    mod_code |= mod;
                                    key_code |= vk;
                                }
                                // HotKey
                                if (!RegisterHotKey(this.Handle, HOTKEY_ID, mod_code, key_code))
                                {
                                    //Error
                                }
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
        private (uint, uint) conv_key_char(string chr)
        {
            var mod_list = new List<(string str, uint code, string desc)>
            {
                ("ALT",0x01,"MOD_ALT"),
                ("CTRL",0x02,"MOD_CONTROL"),
                ("CONTROL",0x02,"MOD_CONTROL"),
                ("SHIFT",0x04,"MOD_SHIFT"),
                ("WIN",0x08,"MOD_WIN"),
            };
            var vk_list = new List<(string str, uint code, string desc)>
            {
                ("LBUTTON",0x01,"マウスの左ボタン"),
                ("RBUTTON",0x02,"マウスの右ボタン"),
                ("CANCEL",0x03,"制御中断処理"),
                ("MBUTTON",0x04,"マウスの中央ボタン"),
                ("XBUTTON1",0x05,"X1 マウス ボタン"),
                ("XBUTTON2",0x06,"X2 マウス ボタン"),
                ("BACK",0x08,"Backspace キー"),
                ("TAB",0x09,"Tab キー"),
                ("CLEAR",0x0C,"Clear キー"),
                ("RETURN",0x0D,"Enter キー"),
                // ("SHIFT",0x10,"Shift キー"),
                // ("CONTROL",0x11,"Ctrl キー"),
                // ("MENU",0x12,"ALT キー"),
                ("PAUSE",0x13,"Pause キー"),
                ("CAPITAL",0x14,"CAPS LOCK キー"),
                ("KANA",0x15,"IME かなモード"),
                ("HANGUL",0x15,"IME ハングル モード"),
                ("IME_ON",0x16,"IME オン"),
                ("JUNJA",0x17,"IME Junja モード"),
                ("FINAL",0x18,"IME Final モード"),
                ("HANJA",0x19,"IME Hanja モード"),
                ("KANJI",0x19,"IME 漢字モード"),
                ("IME_OFF",0x1A,"IME オフ"),
                ("ESCAPE",0x1B,"Esc キー"),
                ("CONVERT",0x1C,"IME 変換"),
                ("NONCONVERT",0x1D,"IME 無変換"),
                ("ACCEPT",0x1E,"IME 使用可能"),
                ("MODECHANGE",0x1F,"IME モード変更要求"),
                ("SPACE",0x20,"Space キー"),
                ("PRIOR",0x21,"PageUp キー"),
                ("NEXT",0x22,"PageDown キー"),
                ("END",0x23,"End キー"),
                ("HOME",0x24,"Home キー"),
                ("LEFT",0x25,"左方向キー"),
                ("UP",0x26,"上方向キー"),
                ("RIGHT",0x27,"右方向キー"),
                ("DOWN",0x28,"下方向キー"),
                ("SELECT",0x29,"Select キー"),
                ("PRINT",0x2A,"Print キー"),
                ("EXECUTE",0x2B,"Execute キー"),
                ("SNAPSHOT",0x2C,"Print Screen キー"),
                ("INSERT",0x2D,"Ins キー"),
                ("DELETE",0x2E,"DEL キー"),
                ("HELP",0x2F,"Help キー"),
                ("0",0x30,"0 キー"),
                ("1",0x31,"1 キー"),
                ("2",0x32,"2 キー"),
                ("3",0x33,"3 キー"),
                ("4",0x34,"4 キー"),
                ("5",0x35,"5 キー"),
                ("6",0x36,"6 キー"),
                ("7",0x37,"7 キー"),
                ("8",0x38,"8 キー"),
                ("9",0x39,"9 キー"),
                ("A",0x41,"A キー"),
                ("B",0x42,"B キー"),
                ("C",0x43,"C キー"),
                ("D",0x44,"D キー"),
                ("E",0x45,"E キー"),
                ("F",0x46,"F キー"),
                ("G",0x47,"G キー"),
                ("H",0x48,"H キー"),
                ("I",0x49,"I キー"),
                ("J",0x4A,"J キー"),
                ("K",0x4B,"K キー"),
                ("L",0x4C,"L キー"),
                ("M",0x4D,"M キー"),
                ("N",0x4E,"N キー"),
                ("O",0x4F,"O キー"),
                ("P",0x50,"P キー"),
                ("Q",0x51,"Q キー"),
                ("R",0x52,"R キー"),
                ("S",0x53,"S キー"),
                ("T",0x54,"T キー"),
                ("U",0x55,"U キー"),
                ("V",0x56,"V キー"),
                ("W",0x57,"W キー"),
                ("X",0x58,"X キー"),
                ("Y",0x59,"Y キー"),
                ("Z",0x5A,"Z キー"),
                // ("LWIN",0x5B,"Windows の左キー"),
                // ("RWIN",0x5C,"右の Windows キー"),
                ("APPS",0x5D,"アプリケーション キー"),
                ("SLEEP",0x5F,"コンピューターのスリープ キー"),
                ("NUMPAD0",0x60,"テンキーの 0 キー"),
                ("NUMPAD1",0x61,"テンキーの 1 キー"),
                ("NUMPAD2",0x62,"テンキーの 2 キー"),
                ("NUMPAD3",0x63,"テンキーの 3 キー"),
                ("NUMPAD4",0x64,"テンキーの 4 キー"),
                ("NUMPAD5",0x65,"テンキーの 5 キー"),
                ("NUMPAD6",0x66,"テンキーの 6 キー"),
                ("NUMPAD7",0x67,"テンキーの 7 キー"),
                ("NUMPAD8",0x68,"テンキーの 8 キー"),
                ("NUMPAD9",0x69,"テンキーの 9 キー"),
                ("MULTIPLY",0x6A,"乗算キー"),
                ("ADD",0x6B,"キーの追加"),
                ("SEPARATOR",0x6C,"区切り記号キー"),
                ("SUBTRACT",0x6D,"減算キー"),
                ("DECIMAL",0x6E,"10 進キー"),
                ("DIVIDE",0x6F,"除算キー"),
                ("F1",0x70,"F1 キー"),
                ("F2",0x71,"F2 キー"),
                ("F3",0x72,"F3 キー"),
                ("F4",0x73,"F4 キー"),
                ("F5",0x74,"F5 キー"),
                ("F6",0x75,"F6 キー"),
                ("F7",0x76,"F7 キー"),
                ("F8",0x77,"F8 キー"),
                ("F9",0x78,"F9 キー"),
                ("F10",0x79,"F10 キー"),
                ("F11",0x7A,"F11 キー"),
                ("F12",0x7B,"F12 キー"),
                ("F13",0x7C,"F13 キー"),
                ("F14",0x7D,"F14 キー"),
                ("F15",0x7E,"F15 キー"),
                ("F16",0x7F,"F16 キー"),
                ("F17",0x80,"F17 キー"),
                ("F18",0x81,"F18 キー"),
                ("F19",0x82,"F19 キー"),
                ("F20",0x83,"F20 キー"),
                ("F21",0x84,"F21 キー"),
                ("F22",0x85,"F22 キー"),
                ("F23",0x86,"F23 キー"),
                ("F24",0x87,"F24 キー"),
                ("NUMLOCK",0x90,"NUM LOCK キー"),
                ("SCROLL",0x91,"ScrollLock キー"),
                // ("LSHIFT",0xA0,"左 Shift キー"),
                // ("RSHIFT",0xA1,"右 Shift キー"),
                // ("LCONTROL",0xA2,"左 Ctrl キー"),
                // ("RCONTROL",0xA3,"右 Ctrl キー"),
                // ("LMENU",0xA4,"左 Alt キー"),
                // ("RMENU",0xA5,"右 Alt キー"),
                ("BROWSER_BACK",0xA6,"ブラウザーの戻るキー"),
                ("BROWSER_FORWARD",0xA7,"ブラウザーの進むキー"),
                ("BROWSER_REFRESH",0xA8,"ブラウザーの更新キー"),
                ("BROWSER_STOP",0xA9,"ブラウザーの停止キー"),
                ("BROWSER_SEARCH",0xAA,"ブラウザーの検索キー"),
                ("BROWSER_FAVORITES",0xAB,"ブラウザーのお気に入りキー"),
                ("BROWSER_HOME",0xAC,"ブラウザーのスタートとホーム キー"),
                ("VOLUME_MUTE",0xAD,"音量ミュート キー"),
                ("VOLUME_DOWN",0xAE,"音量下げるキー"),
                ("VOLUME_UP",0xAF,"音量上げるキー"),
                ("MEDIA_NEXT_TRACK",0xB0,"次のトラックキー"),
                ("MEDIA_PREV_TRACK",0xB1,"前のトラック"),
                ("MEDIA_STOP",0xB2,"メディアの停止キー"),
                ("MEDIA_PLAY_PAUSE",0xB3,"メディアの再生/一時停止キー"),
                ("LAUNCH_MAIL",0xB4,"メール開始キー"),
                ("LAUNCH_MEDIA_SELECT",0xB5,"メディアの選択キー"),
                ("LAUNCH_APP1",0xB6,"アプリケーション 1 の起動キー"),
                ("LAUNCH_APP2",0xB7,"アプリケーション 2 の起動キー"),
                ("OEM_1",0xBA,"その他の文字に使用されます。キーボードによって異なる場合があります。 米国標準キーボードの場合は、 ;: キー"),
                ("OEM_PLUS",0xBB,"どの国/地域の場合でも + 、キー"),
                ("OEM_COMMA",0xBC,"どの国/地域の場合でも , 、キー"),
                ("OEM_MINUS",0xBD,"どの国/地域の場合でも - 、キー"),
                ("OEM_PERIOD",0xBE,"どの国/地域の場合でも . 、キー"),
                ("OEM_2",0xBF,"その他の文字に使用されます。キーボードによって異なる場合があります。 米国標準キーボードの場合は、 /? キー"),
                ("OEM_3",0xC0,"その他の文字に使用されます。キーボードによって異なる場合があります。 米国標準キーボードの場合は、 `~ キー"),
                ("OEM_4",0xDB,"その他の文字に使用されます。キーボードによって異なる場合があります。 米国標準キーボードの場合は、 [{ キー"),
                ("OEM_5",0xDC,"その他の文字に使用されます。キーボードによって異なる場合があります。 米国標準キーボードの場合は、 \\| キー"),
                ("OEM_6",0xDD,"その他の文字に使用されます。キーボードによって異なる場合があります。 米国標準キーボードの場合は、 ]} キー"),
                ("OEM_7",0xDE,"その他の文字に使用されます。キーボードによって異なる場合があります。 米国標準キーボードの場合は、 '\" キー"),
                ("OEM_8",0xDF,"その他の文字に使用されます。キーボードによって異なる場合があります。"),
                ("OEM_102",0xE2,"標準的な US キーボードの <> キー、US 以外の 102 キー キーボードの \\| キー"),
                ("PROCESSKEY",0xE5,"IME PROCESS キー"),
                ("PACKET",0xE7,"Unicode 文字がキーストロークであるかのように渡されます。 VK_PACKET キー値は、キーボード以外の入力手段に使用される 32 ビット仮想キー値の下位ワードです。 詳細については、KEYBDINPUT、SendInput、WM_KEYDOWN、WM_KEYUP の注釈を参照してください"),
                ("ATTN",0xF6,"Attn キー"),
                ("CRSEL",0xF7,"CrSel キー"),
                ("EXSEL",0xF8,"ExSel キー"),
                ("EREOF",0xF9,"EOF 消去キー"),
                ("PLAY",0xFA,"再生キー"),
                ("ZOOM",0xFB,"ズーム キー"),
                ("NONAME",0xFC,"予約済み"),
                ("PA1",0xFD,"PA1 キー"),
                ("OEM_CLEAR",0xFE,"クリア キー")
            };
            uint code = 0x00;
            uint mod_code = 0x00;
            foreach (var mod in mod_list) {
                if (mod.str == chr)
                {
                    mod_code = mod.code;
                }
            }
            foreach (var vk in vk_list) {
                if (vk.str == chr)
                {
                    code = vk.code;
                }
            }
            return (mod_code, code);
        }
    }
}
