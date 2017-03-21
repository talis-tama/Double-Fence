using System;
using System.Drawing;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ClosedXML.Excel;
using System.Collections;

namespace 八重垣MK2
{
    static class Program
    {
        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }

    public partial class Form1 : Form
    {
        ComboBox combobox1, combobox2;
        TextBox textbox1, textbox2;
        Button button1;
        int id = 0;
        int grade, cls;
        int sid = 1;
        bool ck = true;
        public static string excelfile, fdirectory;
        public Form1()
        {
            fdirectory="C:\\Users\\"+Environment.UserName+"\\Desktop\\";
            Width = 320;
            Height = 300;
            Text = "生徒名簿作成ツール";
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;

            MenuStrip strip1 = new MenuStrip();
            SuspendLayout();
            strip1.SuspendLayout();
            ToolStripMenuItem file = new ToolStripMenuItem();
            file.Text = "ファイル(&F)";
            strip1.Items.Add(file);
            ToolStripMenuItem open = new ToolStripMenuItem();
            open.Text = "名簿ファイル読み込み(結合)(&O)";
            open.Click += new EventHandler(open_Click);
            file.DropDownItems.Add(open);
            Controls.Add(strip1);
            MainMenuStrip = strip1;
            strip1.ResumeLayout(false);
            strip1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();

            textbox2 = new TextBox();
            textbox2.Location = new Point(150, 35);
            textbox2.Size = new Size(140, 165);
            textbox2.ReadOnly = true;
            textbox2.Multiline = true;
            textbox2.ScrollBars = ScrollBars.Vertical;
            Controls.Add(textbox2);

            combobox1 = new ComboBox();
            combobox1.Location = new Point(15, 35);
            combobox1.Size = new Size(80, 20);
            combobox1.Text = "学年を選択";
            combobox1.Items.Add("中学1年");
            combobox1.Items.Add("中学2年");
            combobox1.Items.Add("中学3年");
            combobox1.Items.Add("高校1年");
            combobox1.Items.Add("高校2年");
            combobox1.Items.Add("高校3年");
            combobox1.SelectedIndexChanged += new EventHandler(combobox1_Changed);
            Controls.Add(combobox1);

            combobox2 = new ComboBox();
            combobox2.Location = new Point(15, 60);
            combobox2.Size = new Size(80, 20);
            combobox2.Text = "クラスを選択";
            combobox2.SelectedIndexChanged += new EventHandler(combobox2_Changed);
            combobox2.Enabled = false;
            Controls.Add(combobox2);

            textbox1 = new TextBox();
            textbox1.Location = new Point(15, 85);
            textbox1.Size = new Size(120, 20);
            textbox1.Text = "名前を入力 " + sid.ToString() + "番";
            Controls.Add(textbox1);

            button1 = new Button();
            button1.Location = new Point(15, 110);
            button1.Size = new Size(120, 40);
            button1.Enabled = false;
            button1.Text = "登録";
            button1.Click += new EventHandler(button1_Click);
            Controls.Add(button1);

            Button button2 = new Button();
            button2.Location = new Point(15, 160);
            button2.Size = new Size(120, 40);
            button2.Text = "クラス変更";
            button2.Click += new EventHandler(button2_Click);
            Controls.Add(button2);

            Button button3 = new Button();
            button3.Location = new Point(15, 210);
            button3.Size = new Size(120, 40);
            button3.Text = "終了";
            button3.Click += new EventHandler(button3_Click);
            Controls.Add(button3);

            Button button4 = new Button();
            button4.Location = new Point(150, 210);
            button4.Size = new Size(140, 40);
            button4.Text = "Excelシート適用";
            button4.Click += new EventHandler(button4_Click);
            Controls.Add(button4);
        }

        private void open_Click(object sender, EventArgs e)
        {
            string name;
            OpenFileDialog op = new OpenFileDialog();
            op.FileName = "生徒名簿.txt";
            op.InitialDirectory = fdirectory;
            op.Filter = "テキストファイル(*.txt)|*.txt";
            op.Title = "読み込むファイルを選択";
            op.RestoreDirectory = true;
            op.CheckFileExists = true;
            if (op.ShowDialog() == DialogResult.OK)
            {
                name = op.FileName;
                StreamReader opens = new StreamReader(name, Encoding.GetEncoding("shift_jis"));
                StreamWriter write = new StreamWriter(fdirectory+"生徒名簿.txt", true, Encoding.GetEncoding("shift_jis"));
                ck = false;
                string data;
                while (opens.Peek() > -1)
                {
                    data = opens.ReadLine();
                    write.Write(data + write.NewLine);
                }
                write.Close();
                opens.Close();
                StreamReader input = new StreamReader(fdirectory+"生徒名簿.txt", Encoding.GetEncoding("shift_jis"));
                string datas;
                while (input.Peek() > -1)
                {
                    datas = input.ReadToEnd();
                    textbox2.Text = datas;
                }
                input.Close();
            }
        }
        void combobox1_Changed(object sender, EventArgs e)
        {
            combobox2.Enabled = true;
            if (combobox1.SelectedIndex == 0 || combobox1.SelectedIndex == 1 || combobox1.SelectedIndex == 2)
            {
                combobox2.Items.Clear();
                combobox2.Items.Add("1組");
                combobox2.Items.Add("2組");
                combobox2.Items.Add("3組");
            }
            else if (combobox1.SelectedIndex == 3 || combobox1.SelectedIndex == 4 || combobox1.SelectedIndex == 5)
            {
                combobox2.Items.Clear();
                combobox2.Items.Add("A組");
                combobox2.Items.Add("B組");
                combobox2.Items.Add("C組");
                combobox2.Items.Add("D組");
                combobox2.Items.Add("E組");
                combobox2.Items.Add("F組");
                combobox2.Items.Add("G組");
                combobox2.Items.Add("H組");
            }
        }
        void combobox2_Changed(object sender, EventArgs e) { button1.Enabled = true; }
        void button1_Click(object sender, EventArgs e)
        {
            if (textbox1.Text.Length == 0) { MessageBox.Show("名前を入力してください"); }
            else
            {
                int all;
                string sall,name;
                if (combobox1.SelectedIndex == 0) { grade = 1000; }
                else if (combobox1.SelectedIndex == 1) { grade = 2000; }
                else if (combobox1.SelectedIndex == 2) { grade = 3000; }
                else if (combobox1.SelectedIndex == 3) { grade = 4000; }
                else if (combobox1.SelectedIndex == 4) { grade = 5000; }
                else if (combobox1.SelectedIndex == 5) { grade = 6000; }
                if (combobox2.SelectedIndex == 0) { cls = 100; }
                else if (combobox2.SelectedIndex == 1) { cls = 200; }
                else if (combobox2.SelectedIndex == 2) { cls = 300; }
                else if (combobox2.SelectedIndex == 3) { cls = 400; }
                else if (combobox2.SelectedIndex == 4) { cls = 500; }
                else if (combobox2.SelectedIndex == 5) { cls = 600; }
                else if (combobox2.SelectedIndex == 6) { cls = 700; }
                else if (combobox2.SelectedIndex == 7) { cls = 800; }
                sid += 1;
                id += 1;
                name = textbox1.Text;
                combobox1.Enabled = false;
                combobox2.Enabled = false;
                all = grade + cls + id;
                sall = all.ToString();
                textbox1.Text = "名前を入力 " + sid.ToString() + "番";
                if (ck == true)
                {
                    StreamWriter output = new StreamWriter(fdirectory+"生徒名簿.txt", false, System.Text.Encoding.GetEncoding("shift_jis"));
                    output.Write("99" + sall + " " + name + output.NewLine);
                    output.Close();
                    ck = false;
                }
                else
                {

                    StreamWriter output = new StreamWriter(fdirectory+"生徒名簿.txt", true, System.Text.Encoding.GetEncoding("shift_jis"));
                    output.Write("99" + sall + " " + name + output.NewLine);
                    output.Close();
                }
                StreamReader input = new StreamReader(fdirectory+"生徒名簿.txt", System.Text.Encoding.GetEncoding("shift_jis"));
                string data;
                while (input.Peek() > -1)
                {
                    data = input.ReadToEnd();
                    textbox2.Text = data;
                }
                input.Close();
            }
        }
        void button2_Click(object sender, EventArgs e)
        {
            id = 0;
            sid = 1;
            textbox1.Text = "名前を入力 " + sid.ToString() + "番";
            combobox1.Enabled = true;
        }
        void button3_Click(object sender, EventArgs e) { Application.Exit(); }
        void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.FileName = "";
            dialog.InitialDirectory = fdirectory;
            dialog.Filter = "Excelシート(*.xlsx)|*.xlsx";
            dialog.Title = "対象Excelファイルを選択";
            dialog.RestoreDirectory = true;
            dialog.CheckFileExists = true;
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                excelfile = dialog.FileName;
                form2 f2 = new form2();
                f2.ShowDialog();
            }
        }
    }

    public partial class form2 : Form
    {
        public static TextBox textbox1,textbox2,textbox3;
        public form2()
        {
            Width = 200;
            Height = 240;
            Text = "必要事項入力";
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;

            Label label1 = new Label();
            label1.Location = new Point(15, 15);
            label1.Size = new Size(150, 15);
            label1.Text = "操作するエクセルシートの番号";
            Controls.Add(label1);

            textbox1 = new TextBox();
            textbox1.Location = new Point(15, 35);
            textbox1.Size = new Size(60, 20);
            textbox1.TextChanged += new EventHandler(text_Changed);
            Controls.Add(textbox1);

            Label label2 = new Label();
            label2.Location = new Point(15, 60);
            label2.Size = new Size(150, 15);
            label2.Text = "計測する横のセルの数";
            Controls.Add(label2);

            textbox2 = new TextBox();
            textbox2.Location = new Point(15, 80);
            textbox2.Size = new Size(60, 20);
            textbox2.TextChanged += new EventHandler(text_Changed);
            Controls.Add(textbox2);

            Label label3 = new Label();
            label3.Location = new Point(15, 105);
            label3.Size = new Size(150, 15);
            label3.Text = "計測する縦のセルの数";
            Controls.Add(label3);

            textbox3 = new TextBox();
            textbox3.Location = new Point(15, 125);
            textbox3.Size = new Size(60, 20);
            textbox3.TextChanged += new EventHandler(text_Changed);
            Controls.Add(textbox3);

            Button button1 = new Button();
            button1.Location = new Point(15, 150);
            button1.Size = new Size(150, 40);
            button1.Text = "開始";
            button1.Click += new EventHandler(button1_Click);
            Controls.Add(button1);
        }

        void button1_Click(object sender, EventArgs e)
        {
            if (textbox1.Text.Length != 0) {
                if (textbox2.Text.Length != 0) {
                    if (textbox3.Text.Length != 0)
                    {
                        BASIC_information BASIC = new BASIC_information();
                        int sheet_snum = int.Parse(textbox1.Text);
                        BASIC.horizontal_number = int.Parse(textbox2.Text);
                        BASIC.vertical_number = int.Parse(textbox3.Text);
                        string file_point = Form1.fdirectory + "生徒名簿.txt";
                        int vertical_sell = 0;
                        string horizontal_sell;
                        string spot;
                        var wb = new XLWorkbook(Form1.excelfile);
                        for(int n = 1; n <= BASIC.horizontal_number; n++)
                        {
                            horizontal_sell = BASIC.SELL_WORD(n);
                            for(int i = 1; i <= BASIC.vertical_number; i++)
                            {
                                vertical_sell = i;
                                spot = BASIC.Sell_SPOT(vertical_sell, horizontal_sell);

                            }
                        }
                    } else { MessageBox.Show("数値を入力してください", "", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                } else { MessageBox.Show("数値を入力してください", "", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            } else { MessageBox.Show("数値を入力してください", "", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }
        private void text_Changed(object sender, EventArgs e)
        {
            TextBox tb = (TextBox)sender;
            Regex reg = new Regex(@"[^0-9]");
            Match m = reg.Match(tb.Text);
            if (m.Success)
            {
                MessageBox.Show("半角数字以外は入力しないでください。");
                tb.Text = reg.Replace(tb.Text, "");
            }
            if (tb.MaxLength == tb.Text.Length)
            {
                int i = tb.TabIndex + 1;
                foreach (Control c in Controls) { if (c.TabIndex == i) { c.Focus(); } }
            }
        }
    }
    public class BASIC_information
    {
        public char[] SELL_SPOT_WORD;
        public int horizontal_number, vertical_number;
        public string SELL_WORD(int a)
        {
            int b = a / 27;
            int c = a % 27;
            if (c == 0) { c++; }
            return string.Concat(SELL_SPOT_WORD[b], SELL_SPOT_WORD[c]);
        }
        public string Sell_SPOT(int a, string b) { return string.Concat(b, a.ToString()); }
        public int include_check(string a)
        {
            int b, c = 0;
            if (a.IndexOf("99") >= 0)
            {
                b = a.IndexOf("99");
                c = int.Parse(a.Substring(b, b + 6));
            }
            return c;
        }
        public object Txt_read(int a,string b)
        {
            int c = a / 100;
            int d = a % 100;
            string e = b + c.ToString() + ".txt";
            StreamReader txtR = new StreamReader(e, System.Text.Encoding.GetEncoding(932)/*文字コードを指定*/);//テキストファイルのオープン
            string Line = "";
            ArrayList arText = new ArrayList();
            int i = 1;
            while (Line != null)
            {
                Line = txtR.ReadLine();
                if (i == d) { break; }
                i++;
            }
            txtR.Close();
            object x = Line;
            Console.WriteLine(Line);
            return x;
        }

        public BASIC_information()
        {
            SELL_SPOT_WORD = new char[27] { ' ', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z' };
        }
    };
}