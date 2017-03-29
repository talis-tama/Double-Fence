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
        internal static string time;
        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
            DateTime tim = new DateTime();
            tim = DateTime.Now;
            time = tim.ToLongTimeString();
        }
    }
    public partial class Form1 : Form
    {
        ComboBox combobox1, combobox2;
        Label label1;
        TextBox textbox1, textbox2;
        Button button1;
        NumericUpDown numericupdown;
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
            ToolStripMenuItem read = new ToolStripMenuItem();
            ToolStripMenuItem combine = new ToolStripMenuItem();
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
            Controls.Add(textbox1);

            button1 = new Button();
            button1.Location = new Point(15, 210);
            button1.Size = new Size(95, 30);
            button1.Text = "書き込み";
            button1.Click += new EventHandler(button1_click);
            Controls.Add(button1);

            Button button2 = new Button();
            button2.Location = new Point(15, 255);
            button2.Size = new Size(95, 30);
            button2.Text = "Excelシート適用";
            button2.Click += new EventHandler(button2_click);
            Controls.Add(button2);

            textbox2 = new TextBox();
            textbox2.Location = new Point(125, 35);
            textbox2.Size = new Size(200, 250);
            textbox2.ReadOnly = true;
            textbox2.Multiline = true;
            textbox2.ScrollBars = ScrollBars.Vertical;
            Controls.Add(textbox2);
        }
        void read_click(object sender,EventArgs e) { }
        void combine_click(object sender,EventArgs e) { }
        void combobox1_changed(object sender,EventArgs e) { }
        void combobox2_changed(object sender,EventArgs e) { }
        void numericupdown_keypress(object sender,KeyPressEventArgs e) { if ((e.KeyChar < '0' || '9' < e.KeyChar) && e.KeyChar != '\b') { e.Handled = true; } }
        void textbox1_click(object sender,EventArgs e) { textbox1.Text = ""; }
        void button1_click(object sender,EventArgs e) { }
        void button2_click(object sender,EventArgs e) { }
    }
}