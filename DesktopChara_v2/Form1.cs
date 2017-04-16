using System;
using System.Drawing;
using System.Windows.Forms;
using System.Diagnostics;
using SpeechLib;
using CoreTweet;
using System.Data;
using System.Runtime.InteropServices;
using System.Collections;
using System.Threading.Tasks;

namespace DesktopChara
{
    public partial class Form1 : Form
    {
        public enum ExitWindows : uint {
            EWX_LOGOFF = 0x00,
            EWX_SHUTDOWN = 0x01,
            EWX_REBOOT = 0x02,
            EWX_POWEROFF = 0x08,
            EWX_RESTARTAPPS = 0x40,
            EWX_FORCE = 0x04,
            EWX_FORCEIFHUNG = 0x10,
        }

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool ExitWindowsEx(ExitWindows uFlags,int dwReason);
        [DllImport("user32")]
        public static extern void LockWorkStation();

        private Random rnd;
        private CsvFile filelist = new CsvFile("filelist.csv");
        private CsvFile list = new CsvFile("program.csv");
        private CsvFile commandlist = new CsvFile("command.csv");
        private Point lastMousePosition;
        private bool mouseCapture;
        private string standtype = "start";
        private string lasttype;
        private int lastno  = 0;
        private string mode = "clock";
        private Timer timer;
        private Tokens twitter;
        private int starttime;

        //モード切替用(Dictation実装見送り)
        //音声認識オブジェクト
        private SpInProcRecoContext AlwaysRule = null;
        //音声認識のための言語モデル
        private ISpeechRecoGrammar AlwaysGrammarRule = null;
        //音声認識のための言語モデルのルール
        private ISpeechGrammarRule AlwaysGrammarRuleGrammarRule = null;

        //操作用
        //音声認識オブジェクト
        private SpInProcRecoContext ControlRule = null;
        //音声認識のための言語モデル
        private ISpeechRecoGrammar ControlGrammarRule = null;
        //音声認識のための言語モデルのルール
        private ISpeechGrammarRule ControlGrammarRuleGrammarRule = null;

        //プログラム起動用
        //音声認識オブジェクト
        private SpInProcRecoContext ProgramRule = null;
        //音声認識のための言語モデル
        private ISpeechRecoGrammar ProgramGrammarRule = null;
        //音声認識のための言語モデルのルール
        private ISpeechGrammarRule ProgramGrammarRuleGrammarRule = null;

        //音声合成
        //合成音声ライブラリの読み込み
        private SpVoice VoiceSpeeach = null;

        //コンストラクタ
        public Form1()
        {
            InitializeComponent();
            timer = new Timer();
            timer.Tick += new EventHandler(UpdateTime);
            timer.Interval = 1000;
            timer.Enabled = true;
            this.TopMost = true;
            rnd = new Random();
            AlwaysSpeechInit();
            ControlSpeechInit();
            ProgramSpeechInit();
            VoiceInit();
        }

        //ツイッター初期化
        private void InitTwitter()
        {
            if(Program.setting.GetValue("twitter","APIKey") != "")
            {
                try
                {
                    twitter = CoreTweet.Tokens.Create(Program.setting.GetValue("twitter", "APIKey")
                    , Program.setting.GetValue("twitter","APISecret")
                    , Program.setting.GetValue("twitter","AccessToken")
                    , Program.setting.GetValue("twitter","AccessSecret"));
                    Program.setting.SetValue("twitter", "use", "true");
                    //起動情報ツイート
                    DateTime dtNow = DateTime.Now;
                    //twitter.Statuses.Update(new { status = "有栖ちゃんを起動させました\n" + dtNow.ToString("HH:mm:ss") });
                }
                catch
                {
                    MessageBox.Show("Twitterへの接続に失敗しました", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    twitter = null;
                    Program.setting.SetValue("twitter", "use", "false");
                }
            }
        }

        //音声認識初期化
        private void AlwaysSpeechInit()
        {
            //ルール認識 音声認識オブジェクトの生成
            this.AlwaysRule = new SpInProcRecoContext();
            bool hit = false;
            foreach (SpObjectToken recoperson in this.AlwaysRule.Recognizer.GetRecognizers()) //'Go through the SR enumeration
            {
                string language = recoperson.GetAttribute("Language");
                if (language == "411")
                {//日本語を聴き取れる人だ
                    this.AlwaysRule.Recognizer.Recognizer = recoperson; //君に聞いていて欲しい
                    hit = true;
                    break;
                }
            }
            if (!hit)
            {
                MessageBox.Show("日本語認識が利用できません。\r\n日本語音声認識 MSSpeech_SR_ja-JP_TELE をインストールしてください。\r\n");
                Application.Exit();
            }

            //マイクから拾ってね。
            this.AlwaysRule.Recognizer.AudioInput = this.CreateMicrofon();

            //音声認識イベントで、デリゲートによるコールバックを受ける.
            //認識完了
            this.AlwaysRule.Recognition +=
                delegate (int streamNumber, object streamPosition, SpeechLib.SpeechRecognitionType srt, SpeechLib.ISpeechRecoResult isrr)
                {
                    //音声認識終了
                    this.AlwaysGrammarRule.CmdSetRuleState("AlwaysRule", SpeechRuleState.SGDSInactive);
                    //ウィンドウをアクティブにする
                    this.Activate();
                    //聞き取り開始時間取得
                    starttime = Environment.TickCount & int.MaxValue;
                    //聞き取り開始
                    label1_MouseUp(null, null);
                };
            //言語モデルの作成
            this.AlwaysGrammarRule = this.AlwaysRule.CreateGrammar(0);

            this.AlwaysGrammarRule.Reset(0);
            //言語モデルのルールのトップレベルを作成する.
            this.AlwaysGrammarRuleGrammarRule = this.AlwaysGrammarRule.Rules.Add("AlwaysRule",
                SpeechRuleAttributes.SRATopLevel | SpeechRuleAttributes.SRADynamic);
            //認証用文字列の追加.
            this.AlwaysGrammarRuleGrammarRule.InitialState.AddWordTransition(null, Program.skinini.GetValue("skininfo","response","テスト"));
            
            //ルールを反映させる。
            this.AlwaysGrammarRule.Rules.Commit();
        }
        private void ControlSpeechInit()
        {
            //ルール認識 音声認識オブジェクトの生成
            this.ControlRule = new SpInProcRecoContext();
            bool hit = false;
            foreach (SpObjectToken recoperson in this.ControlRule.Recognizer.GetRecognizers()) //'Go through the SR enumeration
            {
                string language = recoperson.GetAttribute("Language");
                if (language == "411")
                {//日本語を聴き取れる人だ
                    this.ControlRule.Recognizer.Recognizer = recoperson; //君に聞いていて欲しい
                    hit = true;
                    break;
                }
            }
            if (!hit)
            {
                MessageBox.Show("日本語認識が利用できません。\r\n日本語音声認識 MSSpeech_SR_ja-JP_TELE をインストールしてください。\r\n");
                Application.Exit();
            }

            //マイクから拾ってね。
            this.ControlRule.Recognizer.AudioInput = this.CreateMicrofon();

            //音声認識イベントで、デリゲートによるコールバックを受ける.
            //認識完了
            this.ControlRule.Recognition +=
                delegate (int streamNumber, object streamPosition, SpeechLib.SpeechRecognitionType srt, SpeechLib.ISpeechRecoResult isrr)
                {
                    //聞き始めから1分経っていたら強制的に終了
                    if ((Environment.TickCount & int.MaxValue) - starttime >= 60000) {
                        label1_MouseUp(null, null);
                        return;
                    }
                    string strText = isrr.PhraseInfo.GetText(0, -1, true);
                    SpeechTextBranch(strText);
                };
            //認識失敗
            this.ControlRule.FalseRecognition +=
                delegate (int streamNumber, object streamPosition, SpeechLib.ISpeechRecoResult isrr)
                {
                    label1.Text = "？？";
                    show("what", 0);
                    //聞き始めから1分経っていたら強制的に終了
                    if ((Environment.TickCount & int.MaxValue) - starttime >= 60000) {
                        label1_MouseUp(null, null);
                    }
                };

            //言語モデルの作成
            this.ControlGrammarRule = this.ControlRule.CreateGrammar(0);

            this.ControlGrammarRule.Reset(0);
            //言語モデルのルールのトップレベルを作成する.
            this.ControlGrammarRuleGrammarRule = this.ControlGrammarRule.Rules.Add("ControlRule",
                SpeechRuleAttributes.SRATopLevel | SpeechRuleAttributes.SRADynamic);
            //認証用文字列の追加.
            ArrayList voicelist = commandlist.GetProgramCommand();
            foreach (string voice in voicelist) {
                this.ControlGrammarRuleGrammarRule.InitialState.AddWordTransition(null, voice);
            }
            this.ControlGrammarRuleGrammarRule.InitialState.AddWordTransition(null, Program.skinini.GetValue("skininfo", "clock", "時計に戻して"));

            //ルールを反映させる。
            this.ControlGrammarRule.Rules.Commit();
        }
        private void ProgramSpeechInit()
        {
            //ルール認識 音声認識オブジェクトの生成
            this.ProgramRule = new SpInProcRecoContext();
            bool hit = false;
            foreach (SpObjectToken recoperson in this.ProgramRule.Recognizer.GetRecognizers()) //'Go through the SR enumeration
            {
                string language = recoperson.GetAttribute("Language");
                if (language == "411")
                {//日本語を聴き取れる人だ
                    this.ProgramRule.Recognizer.Recognizer = recoperson; //君に聞いていて欲しい
                    hit = true;
                    break;
                }
            }
            if (!hit)
            {
                MessageBox.Show("日本語認識が利用できません。\r\n日本語音声認識 MSSpeech_SR_ja-JP_TELE をインストールしてください。\r\n");
                Application.Exit();
            }

            //マイクから拾ってね。
            this.ProgramRule.Recognizer.AudioInput = this.CreateMicrofon();

            //音声認識イベントで、デリゲートによるコールバックを受ける.
            //認識完了
            this.ProgramRule.Recognition +=
                delegate (int streamNumber, object streamPosition, SpeechLib.SpeechRecognitionType srt, SpeechLib.ISpeechRecoResult isrr)
                {
                    string strText = isrr.PhraseInfo.GetText(0, -1, true);
                    ProgramRun(strText);
                };
            //認識失敗
            this.ProgramRule.FalseRecognition +=
                delegate (int streamNumber, object streamPosition, SpeechLib.ISpeechRecoResult isrr)
                {
                    label1.Text = "？？";
                    show("what", 0);
                };

            //言語モデルの作成
            this.ProgramGrammarRule = this.ProgramRule.CreateGrammar(0);

            this.ProgramGrammarRule.Reset(0);
            //言語モデルのルールのトップレベルを作成する.
            this.ProgramGrammarRuleGrammarRule = this.ProgramGrammarRule.Rules.Add("ProgramRule",
                SpeechRuleAttributes.SRATopLevel | SpeechRuleAttributes.SRADynamic);
            //認証用文字列の追加.
            ArrayList voicelist = list.GetProgramCommand();
            foreach(string voice in voicelist)
            {
                this.ProgramGrammarRuleGrammarRule.InitialState.AddWordTransition(null, voice);
            }
            this.ProgramGrammarRuleGrammarRule.InitialState.AddWordTransition(null, Program.skinini.GetValue("skininfo", "clock", "時計に戻して"));

            //ルールを反映させる。
            this.ProgramGrammarRule.Rules.Commit();
        }
        //合成音声初期化
        private void VoiceInit() {
            this.VoiceSpeeach = new SpeechLib.SpVoice();
            //合成音声エンジンで日本語を話す人を探す。(やらなくても動作はするけど、念のため)
            bool hit = false;
            foreach (SpObjectToken voiceperson in this.VoiceSpeeach.GetVoices()) {
                string language = voiceperson.GetAttribute("Language");
                if (language == "411") {//日本語を話す人だ!
                    this.VoiceSpeeach.Voice = voiceperson; //君に読みあげて欲しい
                    hit = true;
                    break;
                }
            }
            if (!hit) {
                MessageBox.Show("日本語合成音声が利用できません。\r\n日本語合成音声 MSSpeech_TTS_ja-JP_Haruka をインストールしてください。\r\n");
            }
        }
        //マイクから読み取るため、マイク用のデバイスを指定する
        private SpObjectToken CreateMicrofon()
        {
            SpeechLib.SpObjectTokenCategory objAudioTokenCategory = new SpeechLib.SpObjectTokenCategory();
            objAudioTokenCategory.SetId(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech Server\v11.0\AudioInput", false);
            SpObjectToken objAudioToken = new SpObjectToken();
            objAudioToken.SetId(objAudioTokenCategory.Default, @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech Server\v11.0\AudioInput", false);
            //return null;
            return objAudioToken;
        }

        //表示前処理
        private void Form1_Load(object sender, EventArgs e)
        {
            IniLoad();
            textBox1.Visible = false;
            button1.Visible = false;
            label1.Text = "";
            show("start", 0);
            if (Program.setting.GetValue("twitter","use","false") == "true") InitTwitter();
            if (Program.UseSpeech) this.AlwaysGrammarRule.CmdSetRuleState("AlwaysRule", SpeechRuleState.SGDSActive);
        }

        //マウス動作取得
        private void Form1_MouseDown(object sender, MouseEventArgs e)
        {
            switch (e.Button)
            {
                case MouseButtons.Right:
                    //マウスの位置を所得
                    this.lastMousePosition = Control.MousePosition;
                    this.mouseCapture = true;

                    //この間だけ画像変更
                    show("surprise",0);
                    break;
            }
        }
        private void Form1_MouseMove(object sender, MouseEventArgs e)
        {
            if (this.mouseCapture == false)
            {
                return;
            }

            // 最新のマウスの位置を取得
            Point mp = Control.MousePosition;

            // 差分を取得
            int offsetX = mp.X - this.lastMousePosition.X;
            int offsetY = mp.Y - this.lastMousePosition.Y;

            // コントロールを移動
            this.Location = new Point(this.Left + offsetX, this.Top + offsetY);

            this.lastMousePosition = mp;
        }
        private void Form1_MouseUp(object sender, MouseEventArgs e)
        {
            switch (e.Button)
            {
                case MouseButtons.Right:
                    this.mouseCapture = false;
                    show(lasttype,lastno);
                    break;
                case MouseButtons.Left:
                    if (this.mouseCapture == true) break;
                    switch (standtype) {
                        case "general":
                            int no = rnd.Next(7);
                            show("random", no);
                            break;
                        case "random":
                        case "change":
                            show("general", 0);
                            break;
                        case "start":
                            show("change", 0);
                            break;
                        default:
                            if (timer.Enabled == true) show("start", 0);
                            break;
                    }
                    break;
            }
        }
        private void Form1_MouseCaptureChanged(object sender, EventArgs e)
        {
            if (this.mouseCapture == true && this.Capture == false)
            {
                this.mouseCapture = false;
            }
        }

        //キー取得
        private void Form1_KeyUp(object sender, KeyEventArgs e)
        {
            switch(e.KeyCode)
            {
                case Keys.F2:
                    if (mode == "clock")
                    {
                        //音声認識ストップ
                        this.AlwaysGrammarRule.CmdSetRuleState("AlwaysRule", SpeechRuleState.SGDSInactive);
                        mode = "setting";
                        //現在の位置を一旦保存
                        IniSave();
                        show("kusonemi", 0);
                        Form2 setting = new Form2();
                        setting.ShowDialog(this);
                        //変更された情報を読み込む
                        IniLoad();
                        InitTwitter();
                        show(lasttype, lastno);
                        mode = "clock";
                        //音声認識再開
                        if(Program.UseSpeech)this.AlwaysGrammarRule.CmdSetRuleState("AlwaysRule", SpeechRuleState.SGDSActive);
                    }
                    break;
                case Keys.Escape:
                    DialogResult result = MessageBox.Show("終了しますか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                    {
                        IniSave();
                        Application.Exit();
                    }
                    break;
                case Keys.ControlKey:
                    this.mouseCapture = false;
                    show(lasttype, lastno);
                    break;
            }
        }

        //画像を表示する
        private void show(string type,int no)
        {
            //キャラ設置
            if (pictureBox2.Image != null) pictureBox2.Image.Dispose();
            pictureBox2.Image = Image.FromFile(filelist.GetPath(type, no));
            if (type == "start" || type == "change" || type == "general" || type == "random") {
                lasttype = type;
                lastno = no;
            } else if (type == "surprise" && standtype != "surprise") {
                lasttype = standtype;
            }
            standtype = type;
            //吹き出し設置
            if (pictureBox1.Image != null) pictureBox1.Image.Dispose();
            pictureBox1.Image = Image.FromFile(filelist.GetPath("ballon", 0));
            //ウィンドウ透過
            this.TransparencyKey = BackColor; //透過pngならこれでいける？
        }

        //時計の描画
        private void UpdateTime(object sender, EventArgs e)
        {
            DateTime dtNow = DateTime.Now;
            label1.Text = dtNow.ToString("yyyy/MM/dd(ddd)") + "\n" + dtNow.ToString("HH:mm:ss");
        }

        //設定ファイルへの書き込み
        private void IniSave()
        {
            Program.setting.SetValue("general","posx", this.Location.X.ToString());
            Program.setting.SetValue("general","posy", this.Location.Y.ToString());
            Program.setting.SetValue("skin","folder", Program.skinfolder);
            Program.setting.SetValue("general", "usespeech", Program.UseSpeech.ToString());
            Program.setting.SetValue("general", "usevoice", Program.UseVoice.ToString());
            Program.setting.SetValue("general", "showintaskbar", Program.ShowInTaskbar.ToString());
        }

        //設定ファイルの読み込み
        private void IniLoad()
        {
            this.Location = new Point(int.Parse(Program.setting.GetValue("general","posx","100")), int.Parse(Program.setting.GetValue("general","posy","100")));
            if(this.Location.X >= Program.DispSize.X || 
                this.Location.Y >= Program.DispSize.Y || this.Location.X < 0 || this.Location.Y < 0)
            {
                this.Location = new Point(100, 100);
            }
            Program.UseSpeech = bool.Parse(Program.setting.GetValue("general","usespeech","False"));
            Program.UseVoice = bool.Parse(Program.setting.GetValue("general", "usevoice", "False"));
            Program.ShowInTaskbar = bool.Parse(Program.setting.GetValue("general", "showintaskbar", "False"));
            this.label1.Font = new Font(label1.Font.FontFamily, int.Parse(Program.setting.GetValue("general", "fontsize", "14")));
            this.ShowInTaskbar = Program.ShowInTaskbar;
        }

        //長押しキー取得(移動用Ctrlキー)
        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.ControlKey:
                    //マウスがウィンドウに乗ってるときのみ実行
                    if (Control.MousePosition.X < this.Location.X || Control.MousePosition.X > this.Location.X + this.Size.Width ||
                        Control.MousePosition.Y < this.Location.Y || Control.MousePosition.Y > this.Location.Y + this.Size.Height) break;
                    this.lastMousePosition = Control.MousePosition;
                    this.mouseCapture = true;
                    //この間だけ画像変更
                    show("surprise", 0);
                    break;
            }
        }

        //実行ボタン
        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "")
            {

                if (mode != "file" && !Program.UseSpeech)
                {
                    SpeechTextBranch(textBox1.Text);
                    textBox1.Text = "";
                    return;
                }
                if(ProgramRun(textBox1.Text) == -1) return;
                textBox1.Text = "";
                label1_MouseUp(null, null);
            }
        }
        
        //音声認識分岐
        private async void SpeechTextBranch(string speechtext)
        {
            //音声認識終了
            this.ControlGrammarRule.CmdSetRuleState("ControlRule", SpeechRuleState.SGDSInactive);
            if(speechtext == Program.skinini.GetValue("skininfo", "clock", "時計に戻して")) {
                label1_MouseUp(null, null);
                return;
            }
            DataRow[] rows = commandlist.dt.Select("voice = '" + speechtext + "'");
            switch ((string)rows[0][1])
            {
                case "runprogram":
                    show("search", 0);
                    label1.Text = "何を実行する？";
                    if (!Program.UseSpeech)
                    {
                        textBox1.Location = new Point(12, 43);
                        button1.Location = new Point(136, 40);
                        textBox1.Visible = true;
                        button1.Visible = true;
                    }
                    mode = "file";
                    if (Program.UseSpeech) this.ProgramGrammarRule.CmdSetRuleState("ProgramRule", SpeechRuleState.SGDSActive);
                    break;
                case "tweet":
                    show("chair", 0);
                    label1.Text = "なんてつぶやく？";
                    if (!Program.UseSpeech)
                    {
                        textBox1.Visible = false;
                        button1.Visible = false;
                        textBox1.Text = "";
                    }
                    mode = "twitter";
                    if(twitter == null) {
                        MessageBox.Show("TwitterAPIが指定されていません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        label1_MouseUp(null, null);
                    }
                    IniSave();
                    Form3 tweet = new Form3();
                    tweet.ShowDialog(this);
                    if (Program.tweetdata != "" && twitter != null)
                    {
                        label1.Text = "ツイートを送信中";
                        label1.Refresh();
                        try
                        {
                            twitter.Statuses.Update(new { status = Program.tweetdata });
                        }
                        catch
                        {
                            MessageBox.Show("ツイートに失敗しました", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    label1_MouseUp(null, null);
                    break;
                case "search":
                    show("search", 0);
                    label1.Text = "何を検索する？";
                    if (!Program.UseSpeech)
                    {
                        textBox1.Visible = false;
                        button1.Visible = false;
                        textBox1.Text = "";
                    }
                    mode = "search";
                    IniSave();
                    Form4 search = new Form4();
                    search.ShowDialog(this);
                    label1_MouseUp(null, null);
                    break;
                case "lockpc":
                    mode = "lockpc";
                    try {
                        DateTime dtNow = DateTime.Now;
                        twitter.Statuses.Update(new { status = "ご主人が席を外されました【有栖ちゃん】\n" + dtNow.ToString("HH:mm:ss") });
                    }
                    catch {
                    }
                    label1_MouseUp(null, null);
                    LockWorkStation();
                    break;
                case "shutdownpc":
                    mode = "shutdownpc";
                    try {
                        DateTime dtNow = DateTime.Now;
                        twitter.Statuses.Update(new { status = "ご主人が外出されました【有栖ちゃん】\n" + dtNow.ToString("HH:mm:ss") });
                    }
                    catch {
                    }
                    show("smile", 0);
                    label1.Text = "シャットダウンするね";
                    await Task.Delay(5 * 1000);
                    label1_MouseUp(null, null);
                    //強制的にシャットダウンする
                    Program.AdjustToken();
                    ExitWindowsEx(ExitWindows.EWX_POWEROFF, 0);
                    break;
                case "battery":
                    show("find", 0);
                    PowerLineStatus pls = SystemInformation.PowerStatus.PowerLineStatus;
                    if (pls == PowerLineStatus.Online)
                    {
                        label1.Text = "AC電源駆動だよ";
                    }
                    else
                    {
                        float blp = SystemInformation.PowerStatus.BatteryLifePercent;
                        label1.Text = "バッテリー残量は\n" + blp * 100 + "% だよ";
                    }
                    mode = "battery";
                    if (Program.UseSpeech) this.ControlGrammarRule.CmdSetRuleState("ControlRule", SpeechRuleState.SGDSActive);
                    break;
                case "listreload":
                    label1.Text = "読み直してるよ";
                    show("sleep", 3);
                    mode = "refresh";
                    //ルール削除
                    this.ProgramGrammarRule.Reset(0);
                    //言語モデルのルールのトップレベルを作成する.
                    this.ProgramGrammarRuleGrammarRule = this.ProgramGrammarRule.Rules.Add("ProgramRule",
                        SpeechRuleAttributes.SRATopLevel | SpeechRuleAttributes.SRADynamic);
                    //認証用文字列の追加.
                    ArrayList voicelist = list.GetProgramCommand();
                    foreach (string voice in voicelist) {
                        this.ProgramGrammarRuleGrammarRule.InitialState.AddWordTransition(null, voice);
                    }
                    this.ProgramGrammarRuleGrammarRule.InitialState.AddWordTransition(null, Program.skinini.GetValue("skininfo", "clock", "時計に戻して"));
                    //ルールを反映させる。
                    this.ProgramGrammarRule.Rules.Commit();
                    MessageBox.Show("更新完了");
                    GC.Collect();
                    label1_MouseUp(null, null);
                    break;
                case "opensetting":
                    label1_MouseUp(null, null);
                    //音声認識ストップ
                    this.AlwaysGrammarRule.CmdSetRuleState("AlwaysRule", SpeechRuleState.SGDSInactive);
                    mode = "setting";
                    //現在の位置を一旦保存
                    IniSave();
                    show("kusonemi", 0);
                    Form2 setting = new Form2();
                    setting.ShowDialog(this);
                    //変更された情報を読み込む
                    IniLoad();
                    InitTwitter();
                    show(lasttype, lastno);
                    //音声認識再開
                    if (Program.UseSpeech) this.AlwaysGrammarRule.CmdSetRuleState("AlwaysRule", SpeechRuleState.SGDSActive);
                    break;
                case "tired":
                    show("tere", 0);
                    if (!Program.UseSpeech) {
                        textBox1.Visible = false;
                        button1.Visible = false;
                        textBox1.Text = "";
                    }
                    label1.Text = "大丈夫？\nおっぱい揉む？";
                    mode = "tired";
                    if (Program.UseSpeech) this.ControlGrammarRule.CmdSetRuleState("ControlRule", SpeechRuleState.SGDSActive);
                    break;
                case "exit":
                    label1.Text = "終了しちゃうの？";
                    show("naki", 0);
                    DialogResult result = MessageBox.Show("終了しますか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                    {
                        IniSave();
                        Application.Exit();
                    }
                    mode = "exit";
                    label1_MouseUp(null, null);
                    break;
            }
        }

        //プログラム実行音声認識
        private int ProgramRun(string command)
        {
            this.ProgramGrammarRule.CmdSetRuleState("ProgramRule", SpeechRuleState.SGDSInactive);
            if (command == Program.skinini.GetValue("skininfo", "clock", "時計に戻して")) {
                label1_MouseUp(null, null);
                return 0;
            }
            while (true) {
                if (!Program.UseSpeech) {
                    try {
                        Process.Start(command);
                        break;
                    }
                    catch {
                        goto listcheck;
                    }
                }
listcheck:
                DataRow[] rows = list.dt.Select("voice = '" + command + "'");
                try {
                    Process.Start((string)rows[0][1]);
                    break;
                }
                catch {
                    MessageBox.Show("実行出来なかったよ\n" + command, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return -1;
                }
            }
            if (Program.UseSpeech) label1_MouseUp(null, null);
            return 0;
        }

        //テキストクリック(音声認識時もここに飛ぶ)
        private void label1_MouseUp(object sender, MouseEventArgs e)
        {
            if (e != null && e.Button == MouseButtons.Right) return;
            if (mode == "clock")
            {
                label1.Text = "どうしたの？";
                //コレジャナイ感が半端ないので一旦実装停止
                //this.VoiceSpeeach.Speak(label1.Text, SpeechVoiceSpeakFlags.SVSFlagsAsync);
                timer.Enabled = false;
                show("voice", 0);
                mode = "voice";
                if (!Program.UseSpeech)
                {
                    //デバッグ用入力ボックス
                    textBox1.Visible = true;
                    button1.Visible = true;
                    textBox1.Focus();
                }
                //音声認識開始
                if (Program.UseSpeech) this.ControlGrammarRule.CmdSetRuleState("ControlRule", SpeechRuleState.SGDSActive);
            }
            else
            {
                //音声認識終了
                this.ControlGrammarRule.CmdSetRuleState("ControlRule", SpeechRuleState.SGDSInactive);
                textBox1.Text = "";
                textBox1.Visible = false;
                button1.Visible = false;
                UpdateTime(null, null);
                timer.Enabled = true;
                show(lasttype, lastno);
                this.Focus();
                mode = "clock";
                //音声認識開始
                if (Program.UseSpeech) this.AlwaysGrammarRule.CmdSetRuleState("AlwaysRule", SpeechRuleState.SGDSActive);
            }
        }

        //通知領域アイコンクリック
        private void notifyIcon1_MouseClick(object sender, MouseEventArgs e)
        {
            if (this.Visible && mode == "clock")
            {
                this.Visible = false;
                if (this.ShowInTaskbar) this.ShowInTaskbar = false;
                this.WindowState = FormWindowState.Minimized;
                this.AlwaysGrammarRule.CmdSetRuleState("AlwaysRule", SpeechRuleState.SGDSInactive);
                timer.Enabled = false;
            }
            else
            {
                this.Visible = true;
                this.WindowState = FormWindowState.Normal;
                this.TopMost = true;
                this.ShowInTaskbar = Program.ShowInTaskbar;
                if (Program.UseSpeech) this.AlwaysGrammarRule.CmdSetRuleState("AlwaysRule", SpeechRuleState.SGDSActive);
                UpdateTime(null, null);
                timer.Enabled = true;
            }
        }
    }
}
