using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Linq;

namespace _6_лаба
{
    public partial class Form1 : Form
    {
        String[] font = { "Segoe UI", "Arial", "Times New Roman", "Segoe Script", "Tahoma" };

        Font lastFont = new Font("Segoe UI", 9);
        Color lastColor = Color.Black;


        List<CustomFont> listfont = new List<CustomFont> { };
        int combobox2;

        struct TextBD
        {
           public int id;
           public String text;

            public TextBD(int v1, String v2) : this()
            {
                this.id = v1;
                this.text = v2;
            }
        }
        List<TextBD> listtext = new List<TextBD> { };  
        int combotext;
        

        //string jsonfile = @"listfont2.json";
        //string textfile = @"text";

        string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=DatabaseFont.accdb;";



        public Form1()
        {

            InitializeComponent();
            comboBox1.Items.AddRange(font);
            comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox1.SelectedIndex = 0;
            comboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboText.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            // comboBox2.SelectedIndex = 0;
            richTextBox1.SelectionChanged += richTextBox1_SelectionChanged;
            //SaveFileDialog saveFileDialog = new SaveFileDialog();
            // saveFileDialog.Filter = 
            // string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;";
            SelectFont();
            SelectText();
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (comboBox1.Items == null) comboBox1.SelectedIndex = 0;
            // string a = comboBox1.Items[comboBox1.SelectedIndex].ToString();
            string a = font[comboBox1.SelectedIndex].ToString();
            var start = richTextBox1.SelectionStart;
            var startlen = richTextBox1.SelectionLength;
            if (richTextBox1.SelectionFont != null)
            {
                richTextBox1.SelectionFont = new Font(a, richTextBox1.SelectionFont.Size, richTextBox1.SelectionFont.Style);
            }
            else
            {
                for (int k = 0; k < startlen; k++)
                {
                    richTextBox1.Select(start + k, 1);
                    richTextBox1.SelectionFont = new Font(a, richTextBox1.SelectionFont.Size, richTextBox1.SelectionFont.Style);
                }
                richTextBox1.SelectionStart = start;
                richTextBox1.SelectionLength = startlen;
            }
            richTextBox1.Focus();
            if (richTextBox1.Text.Length != 0)
            {
                lastFont = richTextBox1.SelectionFont;
                // ChangedFont();
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            if (richTextBox1.SelectionFont != null)
            {
                richTextBox1.SelectionFont = new Font(richTextBox1.SelectionFont.FontFamily, richTextBox1.SelectionFont.Size,
                        Use_BoldStyle());
            }
            else
            {
                var start = richTextBox1.SelectionStart;
                var startlen = richTextBox1.SelectionLength;
                for (int k = 0; k < startlen; k++)
                {
                    richTextBox1.Select(start + k, 1);
                    richTextBox1.SelectionFont = new Font(richTextBox1.SelectionFont.FontFamily, richTextBox1.SelectionFont.Size,
                        Use_BoldStyle());
                }
                richTextBox1.SelectionStart = start;
                richTextBox1.SelectionLength = startlen;
            }
            richTextBox1.Focus();
            if (richTextBox1.Text.Length != 0)
            {
                lastFont = richTextBox1.SelectionFont;
                //  ChangedFont();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (richTextBox1.SelectionFont != null)
            {
                richTextBox1.SelectionFont = new Font(richTextBox1.SelectionFont.FontFamily, richTextBox1.SelectionFont.Size,
                Use_ItalicStyle());
            }
            else
            {
                var start = richTextBox1.SelectionStart;
                var startlen = richTextBox1.SelectionLength;
                for (int k = 0; k < startlen; k++)
                {
                    richTextBox1.Select(start + k, 1);
                    richTextBox1.SelectionFont = new Font(richTextBox1.SelectionFont.FontFamily, richTextBox1.SelectionFont.Size,
                    Use_ItalicStyle());
                }
                richTextBox1.SelectionStart = start;
                richTextBox1.SelectionLength = startlen;
            }
            richTextBox1.Focus();
            if (richTextBox1.Text.Length != 0)
            {
                lastFont = richTextBox1.SelectionFont;
                // ChangedFont();
            }
        }

        private FontStyle Use_BoldStyle()
        {
            if (richTextBox1.SelectionFont.Italic == false)
            {
                if (richTextBox1.SelectionFont.Bold == false)
                {
                    button1.BackColor = System.Drawing.Color.DarkGray;
                    return FontStyle.Bold;
                }
                else
                {
                    button1.BackColor = System.Drawing.Color.White;
                    return FontStyle.Regular;
                }
            }
            else
            {
                if (richTextBox1.SelectionFont.Bold == false)
                {
                    button1.BackColor = System.Drawing.Color.DarkGray;
                    return FontStyle.Bold | FontStyle.Italic;
                }
                else
                {
                    button1.BackColor = System.Drawing.Color.White;
                    return FontStyle.Italic;
                }
            }
        }

        private FontStyle Use_ItalicStyle()
        {
            if (richTextBox1.SelectionFont.Bold == false)
            {
                if (richTextBox1.SelectionFont.Italic == false)
                {
                    button2.BackColor = System.Drawing.Color.DarkGray;
                    return FontStyle.Italic;
                }
                else
                {
                    button2.BackColor = System.Drawing.Color.White;
                    return FontStyle.Regular;
                }
            }
            else
            {
                if (richTextBox1.SelectionFont.Italic == false)
                {
                    button2.BackColor = System.Drawing.Color.DarkGray;
                    return FontStyle.Bold | FontStyle.Italic;
                }
                else
                {
                    button2.BackColor = System.Drawing.Color.White;
                    return FontStyle.Bold;
                }
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            ColorDialog MyDialog = new ColorDialog();
            MyDialog.AllowFullOpen = false;
            MyDialog.ShowHelp = true;
            MyDialog.Color = richTextBox1.SelectionColor;

            if (MyDialog.ShowDialog() == DialogResult.OK)
            {
                richTextBox1.SelectionColor = MyDialog.Color;
                if (MyDialog.Color == Color.Black)
                {
                    button3.BackColor = Color.Black;
                    button3.ForeColor = Color.White;
                }
                else
                {
                    button3.BackColor = MyDialog.Color;
                    button3.ForeColor = Color.Black;
                }
            }
            // lastColor = richTextBox1.SelectionColor;
            richTextBox1.Focus();
            // lastFont = richTextBox1.SelectionFont;
            if (richTextBox1.Text.Length != 0)
            {
                lastColor = richTextBox1.SelectionColor;
                // ChangedFont();
            }
        }


        private void button4_Click(object sender, EventArgs e)
        {
            ToolTip tool = new ToolTip();
            int.TryParse(textBox1.Text.ToString(), out int b);
            if (b < 9)
            {
                MessageBox.Show(textBox1, "Допустимые значение размера шрифта 9-45");
                b = ((int)richTextBox1.SelectionFont.Size);
            }
            if (b > 45)
            {
                MessageBox.Show(textBox1, "Допустимые значение размера шрифта 9-45");
                b = ((int)richTextBox1.SelectionFont.Size);
            }
            if (richTextBox1.SelectionFont != null)
            {
                richTextBox1.SelectionFont = new Font(richTextBox1.SelectionFont.FontFamily, b, richTextBox1.SelectionFont.Style);
            }
            else
            {
                var start = richTextBox1.SelectionStart;
                var startlen = richTextBox1.SelectionLength;
                for (int k = 0; k < startlen; k++)
                {
                    richTextBox1.Select(start + k, 1);
                    richTextBox1.SelectionFont = new Font(richTextBox1.SelectionFont.FontFamily, b, richTextBox1.SelectionFont.Style);
                }
                richTextBox1.SelectionStart = start;
                richTextBox1.SelectionLength = startlen;
            }
            richTextBox1.Focus();
            if (richTextBox1.Text.Length != 0)
            {
                lastFont = richTextBox1.SelectionFont;
                //  ChangedFont();
            }
        }
        private void richTextBox1_SelectionChanged(object sender, EventArgs e)
        {
            if (richTextBox1.SelectionFont != null)
            {
                textBox1.Text = richTextBox1.SelectionFont.Size.ToString();

                if (richTextBox1.SelectionColor == Color.Black)
                {
                    button3.BackColor = Color.Black;
                    button3.ForeColor = Color.White;
                }
                else
                {
                    button3.BackColor = richTextBox1.SelectionColor;
                    button3.ForeColor = Color.Black;
                }
                if (Array.IndexOf(font, richTextBox1.SelectionFont.Name.ToString()) != -1)
                {
                    comboBox1.SelectedIndex = Array.IndexOf(font, richTextBox1.SelectionFont.Name.ToString());
                }
                if (richTextBox1.SelectionFont.Bold == true) button1.BackColor = System.Drawing.Color.DarkGray;
                else button1.BackColor = System.Drawing.Color.White;
                if (richTextBox1.SelectionFont.Italic == true) button2.BackColor = System.Drawing.Color.DarkGray;
                else button2.BackColor = System.Drawing.Color.White;

                if (richTextBox1.Text.Length != 0)
                    lastColor = richTextBox1.SelectionColor;
                // lastFont = richTextBox1.SelectionFont;

                CustomFont custom = new CustomFont(richTextBox1.SelectionFont.FontFamily.Name, richTextBox1.SelectionFont.Size, richTextBox1.SelectionFont.Style, richTextBox1.SelectionColor);
                for (int i = 0; i < listfont.Count; i++)
                    if (listfont[i].EqualsCustom(custom) == true) { 
                        comboBox2.SelectedIndex = i; 
                        return; }
                comboBox2.SelectedIndex = -1;
            }

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

            if (richTextBox1.Text.Length == 0)
            {
                //  if (comboBox2.Items.Count != 0) lastColor = listfont.Last().color;
              //  listfont.Clear();
               // combobox2 = 0;
               // comboBox2.Items.Clear();
                ChangedParam(lastFont, lastColor);
                richTextBox1.SelectionFont = lastFont;
                richTextBox1.SelectionColor = lastColor;
            }

        }

        private void ChangedParam(Font changeFont, Color color)
        {
            if (changeFont != null)
            {
                textBox1.Text = changeFont.Size.ToString();
                richTextBox1.SelectionFont = new Font(richTextBox1.SelectionFont.FontFamily, changeFont.Size);
                if (color == Color.Black)
                {
                    button3.BackColor = Color.Black;
                    button3.ForeColor = Color.White;
                }
                else
                {
                    button3.BackColor = color;
                    button3.ForeColor = Color.Black;
                }
                if (Array.IndexOf(font, changeFont.Name.ToString()) != -1)
                {
                    comboBox1.SelectedIndex = Array.IndexOf(font, changeFont.Name.ToString());
                }
                if (changeFont.Bold == true) button1.BackColor = System.Drawing.Color.DarkGray;
                else button1.BackColor = System.Drawing.Color.White;
                if (changeFont.Italic == true) button2.BackColor = System.Drawing.Color.DarkGray;
                else button2.BackColor = System.Drawing.Color.White;
            }
        }

        private void ChangedParamCustom(CustomFont changeFont)
        {
            if (changeFont != null)
            {
                textBox1.Text = changeFont.size.ToString();
                richTextBox1.SelectionFont = new Font(changeFont.fontFamily, changeFont.size);
                if (changeFont.color == Color.Black)
                {
                    button3.BackColor = Color.Black;
                    button3.ForeColor = Color.White;
                }
                else
                {
                    button3.BackColor = changeFont.color;
                    button3.ForeColor = Color.Black;
                }
                if (Array.IndexOf(font, changeFont.fontFamily.ToString()) != -1)
                {
                    comboBox1.SelectedIndex = Array.IndexOf(font, changeFont.fontFamily.ToString());
                }
                if (changeFont.fontStyle == FontStyle.Bold) button1.BackColor = System.Drawing.Color.DarkGray;
                else button1.BackColor = System.Drawing.Color.White;
                if (changeFont.fontStyle == FontStyle.Italic) button2.BackColor = System.Drawing.Color.DarkGray;
                else button2.BackColor = System.Drawing.Color.White;
                if (changeFont.fontStyle == (FontStyle.Bold | FontStyle.Italic))
                {
                    button1.BackColor = System.Drawing.Color.DarkGray;
                    button2.BackColor = System.Drawing.Color.DarkGray;
                }
                //   else button1.BackColor = System.Drawing.Color.White;
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            int a = comboBox2.SelectedIndex;
            if (a == -1) return;
            ChangedParamCustom(listfont[a]);
            richTextBox1.SelectionFont = new Font(listfont[a].fontFamily, listfont[a].size, listfont[a].fontStyle);
            richTextBox1.SelectionColor = listfont[a].color;
            richTextBox1.Focus();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                if (richTextBox1.Text.Length != 0)
                {
                    if (richTextBox1.SelectionFont != null)
                    {

                        CustomFont custom = new CustomFont(richTextBox1.SelectionFont.FontFamily.Name, richTextBox1.SelectionFont.Size, richTextBox1.SelectionFont.Style, richTextBox1.SelectionColor);
                        for (int i = 0; i < listfont.Count; i++)
                            if (listfont[i].EqualsCustom(custom) == true)
                            {
                                MessageBox.Show("Такой стиль уже существует");
                                richTextBox1.Focus();
                                return;
                            }
                        using (OleDbConnection connection = new OleDbConnection(connectionString))
                        {
                            connection.Open();
                            OleDbCommand myOleDbCommand = connection.CreateCommand();
                            myOleDbCommand.CommandText =
                                    @" INSERT INTO [Font] ( [FontFamily], [Size], [FontStyle], [Color] ) 
                                        VALUES ( @param1, @param2, @param3, @param4 ); " ;
                            myOleDbCommand.Parameters.AddWithValue("@param1", richTextBox1.SelectionFont.FontFamily.Name);
                            myOleDbCommand.Parameters.AddWithValue("@param2", richTextBox1.SelectionFont.Size);
                            myOleDbCommand.Parameters.AddWithValue("@param3", richTextBox1.SelectionFont.Style);
                            int c = ColorTranslator.ToWin32(richTextBox1.SelectionColor);
                            myOleDbCommand.Parameters.AddWithValue("@param4", c.ToString());
                            myOleDbCommand.ExecuteNonQuery();
                            connection.Close();
                        }
                        listfont.Add(custom);
                        combobox2++;
                        comboBox2.Items.Add("Стиль " + combobox2);
                    }
                    else
                    {
                        MessageBox.Show("Выбрано несколько шрифтов");
                        return;
                    }
                    //listfont.Add(richTextBox1.SelectionFont);
                    //  listcolor.Add(richTextBox1.SelectionColor);
                }
            }
            catch (FileNotFoundException f)
            {
                MessageBox.Show("Ошибка " + f.Message);
            }
            catch (UnauthorizedAccessException ex)
            {
                MessageBox.Show("Ошибка " + ex.Message);
            }
            catch (Exception exp)
            {
                MessageBox.Show("Error " + exp.Message);
            }
            richTextBox1.Focus();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                if (richTextBox1.Text.Length == 0)
                {
                    MessageBox.Show("Текст пуст");
                    return;
                }
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    OleDbCommand myOleDbCommand = connection.CreateCommand();
                    myOleDbCommand.CommandText =
                            @" INSERT INTO [Texts] ([Filetext]) 
                                        VALUES ( @param ); ";
                    myOleDbCommand.Parameters.AddWithValue("@param1", richTextBox1.Text);
                    myOleDbCommand.ExecuteNonQuery();
                    connection.Close();
                }
             //   listfont.Add(custom);
                combotext++;
                comboText.Items.Add("Текст " + combotext);
                MessageBox.Show("Текст сохранен");
            }
            catch (FileNotFoundException f)
            {
                MessageBox.Show("Ошибка " + f.Message);
            }
            catch (UnauthorizedAccessException ex)
            {
                MessageBox.Show("Ошибка " + ex.Message);
            }
            catch (Exception exp)
            {
                Console.WriteLine("Error " + exp.Message);
            }
            richTextBox1.Focus();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            try
            {
                if (richTextBox1.Text.Length != 0)
                {
                    DialogResult res = MessageBox.Show("Текущий текст будет утерян. Продолжить?", "Получение текста",
                         MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                    if (res == DialogResult.No)
                    {
                        return;
                    }
                }
                if (comboText.SelectedIndex != -1)
                richTextBox1.Text = listtext[comboText.SelectedIndex].text.ToString();
            }
            catch (FileNotFoundException f)
            {
                MessageBox.Show("Ошибка " + f.Message);
            }
            catch (UnauthorizedAccessException ex)
            {
                MessageBox.Show("Ошибка " + ex.Message);
            }
            catch (Exception exp)
            {
                MessageBox.Show("Error " + exp.Message);
            }
            richTextBox1.Focus();
        }

        private void ChangedFont()
        {
            if (richTextBox1.Text.Length != 0)
            {
                CustomFont custom = new CustomFont(richTextBox1.SelectionFont.FontFamily.Name, richTextBox1.SelectionFont.Size, richTextBox1.SelectionFont.Style, richTextBox1.SelectionColor);
                //listfont.Add(richTextBox1.SelectionFont);
                //  listcolor.Add(richTextBox1.SelectionColor);
                listfont.Add(custom);
                combobox2++;
                comboBox2.Items.Add("Стиль " + combobox2);
            }
            richTextBox1.Focus();
        }

        private void SelectFont()
        {
            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    OleDbCommand myOleDbCommand = connection.CreateCommand();
                    myOleDbCommand.CommandText =
                        "SELECT [id],[FontFamily], [Size], [FontStyle], [Color] " +
                        "FROM Font";
                    DataTable dateFont  = createDataTable("Font", myOleDbCommand);
                    connection.Close();
                    foreach (DataRow dr in dateFont.Rows)
                    {
                        CustomFont customFont = new CustomFont();
                        customFont.fontFamily = dr["FontFamily"].ToString();
                        customFont.size = Convert.ToInt32(dr["Size"]);
                       // FontConverter fontConverter = new FontConverter();
                        customFont.fontStyle = (FontStyle)Enum.Parse(typeof(FontStyle), ( dr["FontStyle"].ToString()) );

                        // ColorConverter converter = new ColorConverter();
                        // customFont.color =(Color) converter.ConvertFromString( dr["Color"].ToString
                        customFont.color = ColorTranslator.FromWin32(Convert.ToInt32(dr["Color"]));
                        customFont.id = Convert.ToInt32( dr["id"]);
                        listfont.Add(customFont);
                        combobox2++;
                        comboBox2.Items.Add("Стиль " + combobox2);
                    }
                }
            }
            catch (Exception exp)
            {
                MessageBox.Show("Error " + exp.Message);
            }

        }

        private void SelectText()
        {
            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    OleDbCommand myOleDbCommand = connection.CreateCommand();
                    myOleDbCommand.CommandText =
                        "SELECT [id],[Filetext] " +
                        "FROM Texts";
                    DataTable dateFont = createDataTable("Texts", myOleDbCommand);
                    connection.Close();
                    foreach (DataRow dr in dateFont.Rows)
                    {
                        TextBD ctr = new TextBD(Convert.ToInt32(dr["id"]), dr["Filetext"].ToString());
                        listtext.Add(ctr);
                        combotext++;
                        comboText.Items.Add("Текст " + combotext);
                    }
                }
            }
            catch (Exception exp)
            {
                MessageBox.Show("Error " + exp.Message);
            }

        }

        static DataTable createDataTable(string tableName, OleDbCommand myOleDbCommand)
        {
            OleDbDataAdapter adapter = new OleDbDataAdapter();
            adapter.SelectCommand = myOleDbCommand;
            adapter.MissingSchemaAction = MissingSchemaAction.AddWithKey;
            DataSet myDataset = new DataSet();
            adapter.Fill(myDataset, tableName);
            return myDataset.Tables[tableName];
        }

        private void deletefontButton_Click(object sender, EventArgs e)
        {
            try
            {
                int id = listfont[comboBox2.SelectedIndex].id;
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    OleDbCommand myOleDbCommand = connection.CreateCommand();
                    myOleDbCommand.CommandText = "DELETE FROM Font WHERE id=" + id;
                    myOleDbCommand.ExecuteNonQuery();
                    connection.Close();
                }
                listfont.RemoveAt(comboBox2.SelectedIndex);
                comboBox2.Items.RemoveAt(comboBox2.SelectedIndex);
                MessageBox.Show("Стиль удален");
                richTextBox1.Focus();
            }
            catch (Exception exp)
            {
                MessageBox.Show("Error " + exp.Message);
            }
        }

        private void deleteTextButton_Click(object sender, EventArgs e)
        {
            try
            {
                int id = listfont[comboBox2.SelectedIndex].id;
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    OleDbCommand myOleDbCommand = connection.CreateCommand();
                    myOleDbCommand.CommandText = "DELETE FROM Font WHERE id=" + id;
                    myOleDbCommand.ExecuteNonQuery();
                    connection.Close();
                }
                listfont.RemoveAt(comboBox2.SelectedIndex);
                comboBox2.Items.RemoveAt(comboBox2.SelectedIndex);
                richTextBox1.Text = "";
                MessageBox.Show("Текст удален");
                richTextBox1.Focus();
            }
            catch (Exception exp)
            {
                MessageBox.Show("Error " + exp.Message);
            }
        }
    }
}
