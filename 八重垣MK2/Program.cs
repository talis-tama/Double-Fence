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
        internal static string filename;
        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            DateTime tim = new DateTime();
            tim = DateTime.Now;
            filename = @"C:\Users\" + Environment.UserName + @"\Desktop\output" + tim.ToString("yyyyMMddHHmmss") + ".txt";
            Application.Run(new Form1());
        }
    }
    public partial class Form1 : Form
    {
        ComboBox combobox1, combobox2;
        Label label1;
        TextBox textbox1, textbox2;
        Button button1, button2;
        NumericUpDown numericupdown;
        ToolStripMenuItem read, combine;
        public Form1()
        {
            Width = 360;
            Height = 340;
            Text = "八重垣";
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;

            MenuStrip menu = new MenuStrip();
            SuspendLayout();
            menu.SuspendLayout();
            ToolStripMenuItem file = new ToolStripMenuItem();
            read = new ToolStripMenuItem();
            combine = new ToolStripMenuItem();
            file.Text = "ファイル";
            read.Text = "開く";
            combine.Text = "ファイル結合";
            combine.Enabled = false;
            menu.Items.Add(file);
            file.DropDownItems.Add(read);
            file.DropDownItems.Add(combine);
            read.Click += new EventHandler(read_click);
            combine.Click += new EventHandler(combine_click);
            Controls.Add(menu);
            MainMenuStrip = menu;
            menu.ResumeLayout(false);
            menu.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
            
            combobox1 = new ComboBox();
            combobox1.Location = new Point(15, 35);
            combobox1.Size = new Size(95, 20);
            combobox1.Text = "学年を選択";
            combobox1.Items.Add("中学1年");
            combobox1.Items.Add("中学2年");
            combobox1.Items.Add("中学3年");
            combobox1.Items.Add("高校1年");
            combobox1.Items.Add("高校2年");
            combobox1.Items.Add("高校3年");
            combobox1.SelectedIndexChanged += new EventHandler(combobox1_changed);
            Controls.Add(combobox1);

            combobox2 = new ComboBox();
            combobox2.Location = new Point(15, 70);
            combobox2.Size = new Size(95, 20);
            combobox2.Text = "クラスを選択";
            combobox2.Enabled = false;
            combobox2.SelectedIndexChanged += new EventHandler(combobox2_changed);
            Controls.Add(combobox2);

            Label label2 = new Label();
            label2.Location = new Point(15, 105);
            label2.Size = new Size(55, 20);
            label2.Text = "出席番号";
            Controls.Add(label2);

            numericupdown = new NumericUpDown();
            numericupdown.Location = new Point(70, 105);
            numericupdown.Size = new Size(40, 20);
            numericupdown.KeyPress += new KeyPressEventHandler(numericupdown_keypress);
            numericupdown.Maximum = 99;
            numericupdown.Minimum = 1;
            numericupdown.ImeMode = ImeMode.Disable;
            numericupdown.ValueChanged += new EventHandler(numericupdown_changed);
            Controls.Add(numericupdown);

            label1 = new Label();
            label1.Location = new Point(15, 140);
            label1.Size = new Size(95, 20);
            Controls.Add(label1);

            textbox1 = new TextBox();
            textbox1.Location = new Point(15, 175);
            textbox1.Size = new Size(95, 20);
            textbox1.Text = "名前を入力";
            textbox1.Click += new EventHandler(textbox1_click);
            textbox1.TextChanged += new EventHandler(textbox1_changed);
            Controls.Add(textbox1);

            button1 = new Button();
            button1.Location = new Point(15, 210);
            button1.Size = new Size(95, 30);
            button1.Enabled = false;
            button1.Text = "書き込み";
            button1.Click += new EventHandler(button1_click);
            Controls.Add(button1);

            button2 = new Button();
            button2.Location = new Point(15, 255);
            button2.Size = new Size(95, 30);
            button2.Text = "Excelシート適用";
            button2.Click += new EventHandler(button2_click);
            button2.Enabled = false;
            Controls.Add(button2);

            textbox2 = new TextBox();
            textbox2.Location = new Point(125, 35);
            textbox2.Size = new Size(200, 250);
            textbox2.ReadOnly = true;
            textbox2.Multiline = true;
            textbox2.ScrollBars = ScrollBars.Vertical;
            Controls.Add(textbox2);
        }
        void textbox2_refresh()
        {
            StreamReader inputrefresh = new StreamReader(Program.filename, Encoding.GetEncoding("shift_jis"));
            textbox2.Text = inputrefresh.ReadToEnd();
            inputrefresh.Close();
        }
        void classinfo() { label1.Text = combobox1.Text + combobox2.Text + numericupdown.Value + "番"; }
        void read_click(object sender,EventArgs e)
        {
            OpenFileDialog readOFD = new OpenFileDialog();
            readOFD.FileName = "";
            readOFD.InitialDirectory = @"C:\Users\" + Environment.UserName + @"\Desktop\";
            readOFD.Filter = "テキストファイル(*.txt)|*.txt";
            readOFD.Title = "開く";
            readOFD.RestoreDirectory = true;
            readOFD.CheckFileExists = true;
            if (readOFD.ShowDialog() == DialogResult.OK)
            {
                Program.filename = readOFD.FileName;
                textbox2_refresh();
                combine.Enabled = true;
                button2.Enabled = true;
                read.Enabled = false;
            }
        }
        void combine_click(object sender,EventArgs e)
        {
            OpenFileDialog combineOFD = new OpenFileDialog();
            combineOFD.FileName = "";
            combineOFD.InitialDirectory = @"C:\Users\" + Environment.UserName + @"\Desktop\";
            combineOFD.Filter = "テキストファイル(*.txt)|*.txt";
            combineOFD.Title = "開く";
            combineOFD.RestoreDirectory = true;
            combineOFD.CheckFileExists = true;
            if (combineOFD.ShowDialog() == DialogResult.OK)
            {
                StreamWriter outputcombine = new StreamWriter(Program.filename, true, Encoding.GetEncoding("shift_jis"));
                StreamReader inputcombine = new StreamReader(combineOFD.FileName, Encoding.GetEncoding("shift_jis"));
                while (inputcombine.Peek() > -1) { outputcombine.Write(inputcombine.ReadLine() + Environment.NewLine); }
                outputcombine.Close();
                inputcombine.Close();
                textbox2_refresh();
                MessageBox.Show("完了しました", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        void combobox1_changed(object sender, EventArgs e)
        {
            combobox2.Enabled = true;
            combobox2.Items.Clear();
            if (combobox1.SelectedIndex >= 0 && combobox1.SelectedIndex <= 2)
            {
                combobox2.Items.Add("1組");
                combobox2.Items.Add("2組");
                combobox2.Items.Add("3組");
            }
            else
            {
                combobox2.Items.Add("A組");
                combobox2.Items.Add("B組");
                combobox2.Items.Add("C組");
                combobox2.Items.Add("D組");
                combobox2.Items.Add("E組");
                combobox2.Items.Add("F組");
                combobox2.Items.Add("G組");
                combobox2.Items.Add("H組");
            }
            combobox2.SelectedIndex = 0;
            classinfo();
        }
        void combobox2_changed(object sender,EventArgs e) { classinfo(); }
        void numericupdown_keypress(object sender,KeyPressEventArgs e) { if ((e.KeyChar < '0' || '9' < e.KeyChar) && e.KeyChar != '\b') { e.Handled = true; } }
        void numericupdown_changed(object sender,EventArgs e) { classinfo(); }
        void textbox1_click(object sender,EventArgs e) { textbox1.Text = ""; }
        void textbox1_changed(object sender,EventArgs e)
        {
            if (textbox1.Text.Length == 0) { button1.Enabled = false; }
            else { button1.Enabled = true; }
        }
        void button1_click(object sender,EventArgs e)
        {
            int dat, grade, clas;
            grade = combobox1.SelectedIndex + 1;
            grade = grade * 1000;
            clas = combobox2.SelectedIndex + 1;
            clas = clas * 100;
            dat = 990000 + grade + clas + int.Parse(numericupdown.Value.ToString());
            StreamWriter outputmain = new StreamWriter(Program.filename, true, Encoding.GetEncoding("shift_jis"));
            outputmain.Write(dat.ToString() + textbox1.Text + Environment.NewLine);
            outputmain.Close();
            textbox2_refresh();
            button2.Enabled = true;
            textbox1.Text = "";
            numericupdown.Value = numericupdown.Value + 1;
            read.Enabled = false;
            combine.Enabled = true;
        }
        void button2_click(object sender,EventArgs e)
        {
            OpenFileDialog excel = new OpenFileDialog();
            excel.FileName = "";
            excel.InitialDirectory = @"C:\Users\" + Environment.UserName + @"\Desktop\";
            excel.Filter = "Excelシート(*.xlsx)|*.xlsx";
            excel.Title = "対象Excelファイル選択";
            excel.RestoreDirectory = true;
            excel.CheckFileExists = true;
            if (excel.ShowDialog() == DialogResult.OK)
            {
                form2.filename = excel.FileName;
                form2 f2 = new form2();
                f2.ShowDialog();
            }
        }
    }
    public partial class form2 : Form
    {
        public static string filename;
        NumericUpDown numericupdown2, numericupdown3, numericupdown4;
        Button button3;
        public form2()
        {
            Width = 195;
            Height = 240;
            Text = "必要事項入力";
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;

            Label label1 = new Label();
            label1.Location = new Point(15, 15);
            label1.Size = new Size(150, 15);
            label1.Text = "操作するエクセルシートの番号";
            Controls.Add(label1);

            numericupdown2 = new NumericUpDown();
            numericupdown2.Location = new Point(15, 35);
            numericupdown2.Size = new Size(60, 20);
            numericupdown2.Minimum = 1;
            numericupdown2.Maximum = 999;
            numericupdown2.ImeMode = ImeMode.Disable;
            numericupdown2.KeyPress += new KeyPressEventHandler(numericupdownkeypressevent);
            Controls.Add(numericupdown2);

            Label label2 = new Label();
            label2.Location = new Point(15, 60);
            label2.Size = new Size(150, 15);
            label2.Text = "計測する横のセルの数";
            Controls.Add(label2);

            numericupdown3 = new NumericUpDown();
            numericupdown3.Location = new Point(15, 80);
            numericupdown3.Size = new Size(60, 20);
            numericupdown3.Minimum = 1;
            numericupdown3.Maximum = 999;
            numericupdown3.ImeMode = ImeMode.Disable;
            numericupdown3.KeyPress += new KeyPressEventHandler(numericupdownkeypressevent);
            Controls.Add(numericupdown3);

            Label label3 = new Label();
            label3.Location = new Point(15, 105);
            label3.Size = new Size(150, 15);
            label3.Text = "計測する縦のセルの数";
            Controls.Add(label3);

            numericupdown4 = new NumericUpDown();
            numericupdown4.Location = new Point(15, 125);
            numericupdown4.Size = new Size(60, 20);
            numericupdown4.Minimum = 1;
            numericupdown4.Maximum = 999;
            numericupdown4.ImeMode = ImeMode.Disable;
            numericupdown4.KeyPress += new KeyPressEventHandler(numericupdownkeypressevent);
            Controls.Add(numericupdown4);

            button3 = new Button();
            button3.Location = new Point(15, 150);
            button3.Size = new Size(150, 40);
            button3.Text = "開始";
            button3.Click += new EventHandler(button3_click);
            Controls.Add(button3);
        }
        void numericupdownkeypressevent(object sender,KeyPressEventArgs e) { if ((e.KeyChar < '0' || '9' < e.KeyChar) && e.KeyChar != '\b') { e.Handled = true; } }
        void button3_click(object sender,EventArgs e)
        {
            int sheet = int.Parse(numericupdown2.Value.ToString());
            int horizontal = int.Parse(numericupdown3.Value.ToString());
            int vertical = int.Parse(numericupdown4.Value.ToString());
            using(var book=new XLWorkbook(filename, XLEventTracking.Disabled))
            {
                StreamReader infile = new StreamReader(Program.filename, Encoding.GetEncoding("shift_jis"));
                var readsheet = book.Worksheet(sheet);
                int count = 0;
                string[] indat = null;
                for(int a = 1; a <= horizontal; a++)
                {
                    for(int b = 1; b <= vertical; b++)
                    {
                        Array.Resize(ref indat, count + 1);
                        var cell = readsheet.Cell(b, a);
                        indat[count] = cell.GetString();
                        count++;
                    }
                }
                string buff;
                int buffint, datint;
                string name;
                while (infile.Peek() > -1)
                {
                    buff = infile.ReadLine();
                    name = buff;
                    buff = buff.Remove(6);
                    name = name.Remove(0, 6);
                    buffint = int.Parse(buff);
                    for(int counta=0; counta < count; counta++)
                    {
                        if (indat[counta].IndexOf("99") >= 0)
                        {
                            datint = int.Parse(indat[counta]);
                            if (buffint == datint) { indat[counta] = name; }
                        }
                    }
                }
                int countb = 0;
                for(int c = 1; c <= horizontal; c++)
                {
                    for(int d = 1; d <= vertical; d++)
                    {
                        var cell1 = readsheet.Cell(d, c);
                        cell1.Value = indat[countb];
                        countb++;
                    }
                }
                book.SaveAs(filename);
                MessageBox.Show("完了しました", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}