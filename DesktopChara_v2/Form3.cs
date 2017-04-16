using SpeechLib;
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace DesktopChara
{
    public partial class Form3 : Form
    {
        [DllImport("SpeechDialog.dll")]
        public extern static bool SpeechDlg(IntPtr Handle, [MarshalAs(UnmanagedType.LPArray)] byte[] res);

        //ダイアログ一般コマンド
        //音声認識オブジェクト
        private SpInProcRecoContext DialogRule = null;
        //音声認識のための言語モデル
        private ISpeechRecoGrammar DialogGrammarRule = null;
        //音声認識のための言語モデルのルール
        private ISpeechGrammarRule DialogGrammarRuleGrammarRule = null;

        public Form3()
        {
            InitializeComponent();
            VoiceCommandInit();
            this.Activate();
        }

        //音声認識初期化
        private void VoiceCommandInit() {
            //ルール認識 音声認識オブジェクトの生成
            this.DialogRule = new SpInProcRecoContext();
            bool hit = false;
            foreach (SpObjectToken recoperson in this.DialogRule.Recognizer.GetRecognizers()) //'Go through the SR enumeration
            {
                string language = recoperson.GetAttribute("Language");
                if (language == "411") {//日本語を聴き取れる人だ
                    this.DialogRule.Recognizer.Recognizer = recoperson; //君に聞いていて欲しい
                    hit = true;
                    break;
                }
            }
            if (!hit) {
                MessageBox.Show("日本語認識が利用できません。\r\n日本語音声認識 MSSpeech_SR_ja-JP_TELE をインストールしてください。\r\n");
                Application.Exit();
            }

            //マイクから拾ってね。
            this.DialogRule.Recognizer.AudioInput = this.CreateMicrofon();

            //音声認識イベントで、デリゲートによるコールバックを受ける.
            //認識完了
            this.DialogRule.Recognition +=
                delegate (int streamNumber, object streamPosition, SpeechLib.SpeechRecognitionType srt, SpeechLib.ISpeechRecoResult isrr) {
                    string strText = isrr.PhraseInfo.GetText(0, -1, true);
                    this.SpeechTextBranch(strText);
                };

            //言語モデルの作成
            this.DialogGrammarRule = this.DialogRule.CreateGrammar(0);

            this.DialogGrammarRule.Reset(0);
            //言語モデルのルールのトップレベルを作成する.
            this.DialogGrammarRuleGrammarRule = this.DialogGrammarRule.Rules.Add("DialogRule",
                SpeechRuleAttributes.SRATopLevel | SpeechRuleAttributes.SRADynamic);
            //認証用文字列の追加.
            this.DialogGrammarRuleGrammarRule.InitialState.AddWordTransition(null, "音声入力");
            this.DialogGrammarRuleGrammarRule.InitialState.AddWordTransition(null, "オーケー");
            this.DialogGrammarRuleGrammarRule.InitialState.AddWordTransition(null, "キャンセル");

            //ルールを反映させる。
            this.DialogGrammarRule.Rules.Commit();
        }

        //音声認証分岐
        private void SpeechTextBranch(String text) {
            this.DialogGrammarRule.CmdSetRuleState("DialogRule", SpeechRuleState.SGDSInactive);
            switch (text) {
                case "音声入力":
                    button3_Click(null, null);
                    break;
                case "オーケー":
                    button1_Click(null, null);
                    break;
                case "キャンセル":
                    button2_Click(null, null);
                    break;
            }
        }

        //マイクから読み取るため、マイク用のデバイスを指定する
        private SpObjectToken CreateMicrofon() {
            SpeechLib.SpObjectTokenCategory objAudioTokenCategory = new SpeechLib.SpObjectTokenCategory();
            objAudioTokenCategory.SetId(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech Server\v11.0\AudioInput", false);
            SpObjectToken objAudioToken = new SpObjectToken();
            objAudioToken.SetId(objAudioTokenCategory.Default, @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech Server\v11.0\AudioInput", false);
            //return null;
            return objAudioToken;
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            //レジストリの読み込み
            this.Location = new Point(int.Parse(Program.setting.GetValue("general","posx")) - this.Size.Width, int.Parse(Program.setting.GetValue("general","posy")) - this.Size.Height);
            //音声認識開始
            this.DialogGrammarRule.CmdSetRuleState("DialogRule", SpeechRuleState.SGDSActive);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            label1.Text = "" + (140 - textBox1.TextLength) + "";
        }

        //OK
        private void button1_Click(object sender, EventArgs e)
        {
            Program.tweetdata = textBox1.Text;
            textBox1.Text = "";
            this.Dispose();
        }

        //キャンセル
        private void button2_Click(object sender, EventArgs e)
        {
            Program.tweetdata = "";
            textBox1.Text = "";
            this.Dispose();
        }

        //音声認識
        private void button3_Click(object sender, EventArgs e)
        {
            bool res;
            byte[] res_byte = new byte[4096];
            res = SpeechDlg(IntPtr.Zero, res_byte);
            if (res)
            {
                textBox1.Text = System.Text.Encoding.GetEncoding("shift_jis").GetString(res_byte);           
            }
            this.DialogGrammarRule.CmdSetRuleState("DialogRule", SpeechRuleState.SGDSActive);
        }
    }
}

