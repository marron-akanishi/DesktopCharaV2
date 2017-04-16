using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace DesktopChara
{
    public partial class Form4 : Form
    {
        [DllImport("SpeechDialog.dll")]
        public extern static bool SpeechDlg(IntPtr Handle, [MarshalAs(UnmanagedType.LPArray)] byte[] res);
        [DllImport("GoogleSearch.dll")]
        public extern static void Search(int mode, string in_str);

        public Form4()
        {
            InitializeComponent();
            this.Activate();
        }

        private void Form4_Load(object sender, EventArgs e)
        {
            //レジストリの読み込み
            this.Location = new Point(int.Parse(Program.setting.GetValue("general","posx")) - this.Size.Width, int.Parse(Program.setting.GetValue("general","posy")) - this.Size.Height);
        }

        //音声入力
        private void button1_Click(object sender, EventArgs e)
        {
            bool res;
            byte[] res_byte = new byte[4096];
            res = SpeechDlg(IntPtr.Zero, res_byte);
            if (res)
            {
                textBox1.Text = System.Text.Encoding.GetEncoding("shift_jis").GetString(res_byte);
            }
        }

        //キャンセル
        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            this.Dispose();
        }
        
        //OK
        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "") button2_Click(null, null);

            //ここはもっとスマートにできそう
            int selectid = 0;
            if (radioButton1.Checked) selectid = 0;
            if (radioButton6.Checked) selectid = 5;
            if (radioButton7.Checked) selectid = 6;

            //これ、いる？
            Encoding sjisEnc = Encoding.GetEncoding("Shift_JIS");
            byte[] bytes = sjisEnc.GetBytes(textBox1.Text);
            //MemoryStreamオブジェクトの作成
            System.IO.MemoryStream ms = new System.IO.MemoryStream();
            //バイト型配列をMemoryStreamに追加する
            ms.Write(bytes, 0, bytes.Length);
            //さらに追加する
            byte[] bs = new byte[] { 0 };
            ms.Write(bs, 0, bs.Length);
            //MemoryStreamの大きさを変更する
            ms.SetLength(bytes.Length + bs.Length);
            //MemoryStreamの内容をバイト型配列に変換する
            bytes = ms.ToArray();
            //閉じる
            ms.Close();

            Search(selectid, textBox1.Text);
            this.Dispose();
        }
    }
}
