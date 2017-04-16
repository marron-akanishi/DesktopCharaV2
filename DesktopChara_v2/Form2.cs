using System;
using System.Drawing;
using System.Windows.Forms;

namespace DesktopChara
{
    public partial class Form2 : Form
    {
        private string skinname;

        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            //設定読み込み
            this.Location = new Point(int.Parse(Program.setting.GetValue("general","posx")) - this.Size.Width, int.Parse(Program.setting.GetValue("general","posy")) - this.Size.Height);
            textBox1.Text = Program.setting.GetValue("twitter","APIKey");
            textBox2.Text = Program.setting.GetValue("twitter","APISecret");
            textBox3.Text = Program.setting.GetValue("twitter","AccessToken");
            textBox4.Text = Program.setting.GetValue("twitter","AccessSecret");
            checkBox1.Checked = bool.Parse(Program.setting.GetValue("general", "usespeech", "False"));
            checkBox2.Checked = bool.Parse(Program.setting.GetValue("general", "usevoice", "False"));
            checkBox3.Checked = bool.Parse(Program.setting.GetValue("general", "showintaskbar", "False"));
            //スキンデータの読み込み
            skinname = Program.skinini.GetValue("skininfo","name","スキン名");
            string skinicon = Program.skinini.GetValue("skininfo", "icon", "icon.png");
            label2.Text = skinname;
            pictureBox1.Image = Image.FromFile(Program.skinpath+skinicon);
            //コマンドリスト表示
            CsvFile commandlist = new CsvFile("command.csv");
            foreach(string command in commandlist.GetProgramCommand()){
                richTextBox1.SelectedText = command + "\n";
            }
    }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            Program.UseSpeech = checkBox1.Checked;
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e) {
            Program.UseVoice = checkBox2.Checked;
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e) {
            Program.ShowInTaskbar = checkBox3.Checked;
        }

        //OK
        private void button1_Click(object sender, EventArgs e)
        {
            IniSave();
            this.Dispose();
        }

        //キャンセル
        private void button2_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        //設定保存
        private void IniSave()
        {
            Program.setting.SetValue("twitter","APIKey", textBox1.Text);
            Program.setting.SetValue("twitter","APISecret", textBox2.Text);
            Program.setting.SetValue("twitter","AccessToken", textBox3.Text);
            Program.setting.SetValue("twitter","AccessSecret", textBox4.Text);
            Program.setting.SetValue("general","usespeech", Program.UseSpeech.ToString());
            Program.setting.SetValue("general","usevoice", Program.UseVoice.ToString());
            Program.setting.SetValue("general", "showintaskbar", Program.ShowInTaskbar.ToString());
        }

        private void button4_Click(object sender, EventArgs e) {
            System.Diagnostics.Process.Start(Program.basepath);
        }
    }
}
