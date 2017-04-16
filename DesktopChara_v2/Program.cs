using System;
using System.Collections;
using System.Windows.Forms;
using System.Data;
using System.Runtime.InteropServices;
using System.Text;
using System.Drawing;

namespace DesktopChara
{
    static class Program
    {
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GetCurrentProcess();
        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool OpenProcessToken(IntPtr ProcessHandle,
            uint DesiredAccess,
            out IntPtr TokenHandle);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool CloseHandle(IntPtr hObject);
        [DllImport("advapi32.dll", SetLastError = true,
            CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        private static extern bool LookupPrivilegeValue(string lpSystemName,
            string lpName,
            out long lpLuid);
        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool AdjustTokenPrivileges(IntPtr TokenHandle,
                    bool DisableAllPrivileges,
                    ref TOKEN_PRIVILEGES NewState,
                    int BufferLength,
                    IntPtr PreviousState,
                    IntPtr ReturnLength);
        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        private struct TOKEN_PRIVILEGES {
            public int PrivilegeCount;
            public long Luid;
            public int Attributes;
        }

        public static string basepath = Application.StartupPath;
        public static string skinpath = "";
        public static IniFile setting = new IniFile(basepath + @"\setting.ini");
        public static IniFile skinini;
        public static string skinfolder = "default";
        public static string tweetdata = "";
        public static bool UseSpeech = false;
        public static bool UseVoice = false;
        public static bool ShowInTaskbar = false;
        public static Point DispSize;

        //シャットダウンするためのセキュリティ特権を有効にする
        public static void AdjustToken() {
            const uint TOKEN_ADJUST_PRIVILEGES = 0x20;
            const uint TOKEN_QUERY = 0x8;
            const int SE_PRIVILEGE_ENABLED = 0x2;
            const string SE_SHUTDOWN_NAME = "SeShutdownPrivilege";

            if (Environment.OSVersion.Platform != PlatformID.Win32NT)
                return;

            IntPtr procHandle = GetCurrentProcess();

            //トークンを取得する
            IntPtr tokenHandle;
            OpenProcessToken(procHandle,
                TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, out tokenHandle);
            //LUIDを取得する
            TOKEN_PRIVILEGES tp = new TOKEN_PRIVILEGES();
            tp.Attributes = SE_PRIVILEGE_ENABLED;
            tp.PrivilegeCount = 1;
            LookupPrivilegeValue(null, SE_SHUTDOWN_NAME, out tp.Luid);
            //特権を有効にする
            AdjustTokenPrivileges(
                tokenHandle, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero);

            //閉じる
            CloseHandle(tokenHandle);
        }

        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //前回設定したスキンを読み込む
            skinfolder = setting.GetValue("skin", "folder", skinfolder);
            skinpath = basepath + @"\skins\" + skinfolder + @"\";
            skinini = new IniFile(skinpath + "skin.ini");
            //全ディスプレイを合わせたサイズを取得する
            DispSize.X = 0;
            DispSize.Y = 0;
            foreach (System.Windows.Forms.Screen s in System.Windows.Forms.Screen.AllScreens) {
                DispSize.X += s.Bounds.Width;
                DispSize.Y += s.Bounds.Height;
            }
            Application.Run(new Form1());
        }
    }

    /// <summary>
    /// スキン用csv読み込みクラス
    /// </summary>
    public class CsvFile {
        public DataTable dt = new DataTable();

        /// <summary>
        /// csvファイルを開きます。
        /// </summary>
        /// <param name="filename">開くファイル名を指定します</param>
        public CsvFile(string filename) {
            //接続文字列
            string conString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
                + Program.skinpath + ";Extended Properties=\"text;HDR=Yes;FMT=Delimited\"";
            System.Data.OleDb.OleDbConnection con =
                new System.Data.OleDb.OleDbConnection(conString);

            string commText = "SELECT * FROM [" + filename + "]";
            System.Data.OleDb.OleDbDataAdapter da =
                new System.Data.OleDb.OleDbDataAdapter(commText, con);

            //DataTableに格納する
            da.Fill(dt);
        }

        /// <summary>
        /// 画像リスト内から検索を行います。
        /// </summary>
        /// <returns>画像のパスを返します。</returns>
        public string GetPath(string type, int no) {
            string path = "";
            try {
                DataRow[] rows = dt.Select("type = '" + type + "'");
                path = Program.skinpath + (string)rows[no][0];
            }
            catch (Exception) {
                MessageBox.Show("ファイルリスト参照中にエラーが発生しました\n初期画像に戻します\ntype=" + type + ",no=" + no, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                DataRow[] rows = dt.Select("type = 'start'");
                path = Program.skinpath + (string)rows[0][0];
            }
            return path;
        }

        /// <summary>
        /// プログラム実行コマンドを読み込みます。
        /// </summary>
        /// <returns>リストを返します</returns>
        public ArrayList GetProgramCommand() {
            ArrayList list = new ArrayList();
            //DataTable内のVoiceデータをArrayに収納
            for (int i = 0; i < dt.Rows.Count; i++) {
                DataRow[] rows = dt.Select();
                list.Add((string)rows[i][0]);
            }
            return list;
        }
    }

    /// <summary>
    /// INIファイルを読み書きするクラス
    /// </summary>
    public class IniFile
    {
        [DllImport("kernel32.dll")]
        private static extern int GetPrivateProfileString(
            string lpApplicationName,
            string lpKeyName,
            string lpDefault,
            StringBuilder lpReturnedstring,
            int nSize,
            string lpFileName);

        [DllImport("kernel32.dll")]
        private static extern int WritePrivateProfileString(
            string lpApplicationName,
            string lpKeyName,
            string lpstring,
            string lpFileName);

        string filePath;

        /// <summary>
        /// ファイル名を指定して初期化します。
        /// ファイルが存在しない場合は初回書き込み時に作成されます。
        /// </summary>
        public IniFile(string filePath)
        {
            this.filePath = filePath;
        }

        /// <summary>
        /// sectionとkeyからiniファイルの設定値を取得、設定します。 
        /// </summary>
        /// <returns>指定したsectionとkeyの組合せが無い場合は""が返ります。</returns>
        public string this[string section, string key]
        {
            set
            {
                WritePrivateProfileString(section, key, value, filePath);
            }
            get
            {
                StringBuilder sb = new StringBuilder(256);
                GetPrivateProfileString(section, key, string.Empty, sb, sb.Capacity, filePath);
                return sb.ToString();
            }
        }

        /// <summary>
        /// sectionとkeyからiniファイルの設定値を取得します。
        /// 指定したsectionとkeyの組合せが無い場合はdefaultvalueで指定した値が返ります。
        /// </summary>
        /// <returns>
        /// 指定したsectionとkeyの組合せが無い場合はdefaultvalueで指定した値が返ります。
        /// </returns>
        public string GetValue(string section, string key, string defaultvalue="")
        {
            StringBuilder sb = new StringBuilder(256);
            GetPrivateProfileString(section, key, defaultvalue, sb, sb.Capacity, filePath);
            return sb.ToString();
        }

        /// <summary>
        /// 指定されたsectionとkeyにデータを書き込みます。
        /// 指定したsectionとkeyの組合せが無い場合は生成されます。
        /// </summary>
        public void SetValue(string section, string key, string value) {
            WritePrivateProfileString(section, key, value, filePath);
            return;
        }
    }
}
