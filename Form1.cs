using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
//using Spire.Pdf;
using System.Text.RegularExpressions;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;



namespace parse_PDF
{
    public partial class Form1 : Form
    {

        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
       

        public Form1()
        {
            InitializeComponent();
            richTextBox1.Text = "Прывітанне, гультай!";
        }

           

        private void BrowseFolderButton_Click(object sender, EventArgs e)
        {
            DialogResult result = this.folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                textBox1.Text = folderBrowserDialog1.SelectedPath;

            }

            richTextBox1.Text += Environment.NewLine + "- папка: " + textBox1.Text;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            /// PartReport
            if (radioButton1.Checked == true)
            {
                
                string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                long ellapledTicks = DateTime.Now.Ticks;
                richTextBox1.Text += Environment.NewLine + " ------- ";
                string[] part_null = new string[] { "Part Name", "Quantity", "Creation Date", "Material",
                                                "Thickness", "Cuttings",
                                                "Weight(kg)","Dimensions","Area","Length", "File", "Date creation file" };

                var csv = new StringBuilder();

                //Suggestion made by KyleMit
                var newLine = string.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10};{11}", part_null[0], part_null[1],
                                                part_null[2], part_null[3], part_null[4], part_null[5], part_null[6],
                                                part_null[7], part_null[8], part_null[9], part_null[10], part_null[11]);
                csv.AppendLine(newLine);

                File.WriteAllText(desktop + "\\Report.csv", csv.ToString());
                richTextBox1.Text += Environment.NewLine + " Створаны: " + desktop + "\\Report.csv";


                if (subfolder.Checked == true)
                {
                    string[] subdirectoryEntries = Directory.GetDirectories(textBox1.Text);


                    foreach (string subdirectory in subdirectoryEntries)
                    {
                        string[] subsubdir = Directory.GetDirectories(subdirectory);

                        foreach (string ss in subsubdir)
                        {

                            string[] subsubsubdir = Directory.GetDirectories(ss);
                            foreach (string sss in subsubsubdir)
                            {
                                richTextBox1.Text += Environment.NewLine + " Знойдзена тэчка: " + sss;
                                parsers(sss, desktop);
                                
                            }

                            richTextBox1.Text += Environment.NewLine + " Знойдзена тэчка: " + ss;
                            parsers(ss, desktop);
                        }

                        richTextBox1.Text += Environment.NewLine + " Знойдзена тэчка: " + subdirectory;
                        parsers(subdirectory, desktop);

                    }
                }

                richTextBox1.Text += Environment.NewLine + " Знойдзена тэчка: " + textBox1.Text;
                parsers(textBox1.Text, desktop);

                ellapledTicks = DateTime.Now.Ticks - ellapledTicks;
                richTextBox1.Text += Environment.NewLine + " Выкарыстана: " + Math.Round(ellapledTicks / 10000000.0, 3) + "c";

            }
            
            /// NestReport
            else if (radioButton2.Checked == true)
            {
                string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                long ellapledTicks = DateTime.Now.Ticks;
                richTextBox1.Text += Environment.NewLine + " ------- ";
                string[] part_null = new string[] { "File", "Date creation file", "JOB", "Material", "Thickness",
                                                    "Weight(kg)", "X Dimension", "Y Dimension", "Usage(kg)", "Usage(%)", "Date Nest"};

                var csv = new StringBuilder();

                var newLine = string.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10}", part_null[0], part_null[1],
                                                part_null[2], part_null[3], part_null[4], part_null[5], part_null[6],
                                                part_null[7], part_null[8], part_null[9], part_null[10]);
                csv.AppendLine(newLine);

                File.WriteAllText(desktop + "\\NestReport.csv", csv.ToString());
                richTextBox1.Text += Environment.NewLine + " Створаны: " + desktop + "\\NestReport.csv";
              /// поиск по п
                if (subfolder.Checked == true)
                {
                    string[] subdirectoryEntries = Directory.GetDirectories(textBox1.Text);


                    foreach (string subdirectory in subdirectoryEntries)
                    {
                        string[] subsubdir = Directory.GetDirectories(subdirectory);

                        foreach (string ss in subsubdir)
                        {

                            string[] subsubsubdir = Directory.GetDirectories(ss);
                            foreach (string sss in subsubsubdir)
                            {
                                richTextBox1.Text += Environment.NewLine + " Знойдзена тэчка: " + sss;
                                nest_parsers(sss, desktop);
                            }

                            richTextBox1.Text += Environment.NewLine + " Знойдзена тэчка: " + ss;
                            nest_parsers(ss, desktop);
                        }

                        richTextBox1.Text += Environment.NewLine + " Знойдзена тэчка: " + subdirectory;
                        nest_parsers(subdirectory, desktop);

                    }
                }

                richTextBox1.Text += Environment.NewLine + " Знойдзена тэчка: " + textBox1.Text;
                nest_parsers(textBox1.Text, desktop);

                ellapledTicks = DateTime.Now.Ticks - ellapledTicks;
                richTextBox1.Text += Environment.NewLine + " Выкарыстана: " + Math.Round(ellapledTicks / 10000000.0, 3) + "c";

            }
        }


        string get_text_to_pdf(string path)
        {
            ///richTextBox1.Text += Environment.NewLine + " - " + path;
            try
            {
                //Console.WriteLine("- " + path);
                PdfReader reader = new PdfReader(path);
                MemoryStream ms = new MemoryStream();
                PdfReader.unethicalreading = true;
                PdfStamper stamper = new PdfStamper(reader, ms);
                stamper.Close();

                byte[] result = ms.ToArray();
                PdfReader reader2 = new PdfReader(result);

                StringBuilder text = new StringBuilder();

                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    text.Append(PdfTextExtractor.GetTextFromPage(reader2, i));
                }

                //Console.WriteLine("-++---" + text.ToString());

                return text.ToString();

            }
            catch
            {
                //код, который нужно выполнить, если случилась ошибка в блоке try
                richTextBox1.Text += Environment.NewLine + " Памылка: " + path;

            }
            finally
            {
                
                //сюда можно добавить любой код который должен выполнится независимо от того, случилась ли ошибка
            }
            return "";
        }

        string check_data(string data)
        {
            try 
            {
                int ret = Int32.Parse(data);
                string resu = data.Substring(0, 2);
                string resu2 = data.Substring(2, 2);

                return "01." +resu + ".20" + resu2;
            }
            catch 
            {
                //код, который нужно выполнить, если случилась ошибка в блоке try
                Console.WriteLine("error");
                return "null";
            }
            finally
            {
                //сюда можно добавить любой код который должен выполнится независимо от того, случилась ли ошибка
            }
        }


        public void nest_parsers(string patch_dir, string desk)
        {
            string[] second = Directory.GetFiles(patch_dir); // путь к папке
            int cont = 0;
            int part_num = 0;

            for (int i = 0; i < second.Length; i++)
            {

                /// тут пишем для парсинга ***Repot.pdf
                if (second[i].ToString().Contains("Part") == false & second[i].ToString().EndsWith("Report.pdf"))
                {

                    string[] dlina = second[i].Replace('_', ';').Split(';');
                    int ddl = dlina.Length;
                    Console.WriteLine(" --- " + ddl + " / " + check_data(dlina[ddl-2].ToString()));

                    /// проверить является ли это датой
                    if (check_data(dlina[ddl - 2].ToString()) != "null")
                    {
                        //cont = cont + 1;

                        string buffer2 = get_text_to_pdf(second[i]);
                        buffer2 = System.Text.RegularExpressions.Regex.Replace(buffer2, @"\s+", " ");

                        //Console.WriteLine( "---" + buffer2);

                        ///терерь парсить содержимое
                        string[] stringSeparators = new string[] { "Sheets" };
                        string[] words = buffer2.Split(stringSeparators, StringSplitOptions.None);


                        string[] output = new string[11];
                        /// Теперь заполняем карточку для каждого Sheet
                        for (int j = 1; j < words.Length; j++)
                        {

                            ///Console.WriteLine(j + "-aaa-" + words[j]);
                            ///
                            char[] separators = new char[] { ' ', '\n', '\r' };
                            string[] card_part = words[j].Split(separators, StringSplitOptions.RemoveEmptyEntries);
                            for (int a = 0; a < card_part.Length; a++)
                            {
                                ///Console.WriteLine(a + " - " + card_part[a]);

                                ///заполнить output
                                ///0 - "File", 
                                output[0] = second[i];
                                ///1 - "Date creation file", 
                                output[1] = File.GetCreationTime(second[i]).ToString().Split(' ')[0];
                                ///2 - "JOB", 
                                if (card_part[a] == "Machine")
                                {
                                    output[2] = card_part[a + 1];
                                }
                                ///3 - "Material",
                                else if (card_part[a] == "Material") { output[3] = card_part[a + 1]; }
                                ///4 - "Thickness", 
                                else if (card_part[a] == "Thickness") { output[4] = card_part[a + 1].Replace(',', '.'); }
                                ///5 - "Weight(kg)",
                                else if (card_part[a] == "Weight(kg)") { output[5] = card_part[a + 1].Replace(',', '.'); }
                                ///6 - "X Dimension", 
                                else if (card_part[a] == "Dimension")
                                {
                                    if (card_part[a - 1] == "X")
                                    {
                                        output[6] = card_part[a + 1].Split('(')[0].Replace(',', '.');
                                    }
                                    ///7 - "Y Dimension",
                                    else if (card_part[a - 1] == "Y")
                                    {
                                        output[7] = card_part[a + 1].Split('(')[0].Replace(',', '.');
                                    }
                                }
                                ///8 - "Usage(kg)",
                                else if (card_part[a] == "Usage(kg)") { output[8] = card_part[a + 1].Replace(',', '.'); }
                                ///9 - "Usage(%)"
                                else if (card_part[a] == "Usage(%)") { output[9] = card_part[a + 1].Replace(',', '.'); }
                                ///10 - "Date Nest"
                                output[10] = check_data(dlina[ddl - 2].ToString());

                            }

                            if (checkBox1.Checked == false)
                            {
                                /// если выбран диапазон дат
                                if (dateTimePicker1.Value <= File.GetCreationTime(second[i]) & dateTimePicker2.Value >= File.GetCreationTime(second[i]))
                                {
                                    var newLine2 = string.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10}\n", output[0], output[1],
                                                            output[2], output[3], output[4], output[5], output[6], output[7],
                                                            output[8], output[9], output[10]);

                                    File.AppendAllText(desk + "\\NestReport.csv", newLine2);
                                    //cont += cont;
                                }
                            }
                            else /// если за всё время
                            {
                                var newLine2 = string.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10}\n", output[0], output[1],
                                                        output[2], output[3], output[4], output[5], output[6], output[7],
                                                        output[8], output[9], output[10]);

                                File.AppendAllText(desk + "\\NestReport.csv", newLine2);

                            }

                            //cont += cont;
                        }

                    }
                    /// если нет записываем в лог исключение
                    else
                    {
                        richTextBox1.Text += Environment.NewLine+ "Імя файла не адпавядае ўмове: " + second[i].ToString() ;
                    }

                }

                
               
            }
            ///richTextBox1.Text += Environment.NewLine + "- Файлаў: " + cont;

        }

        public void parsers(string patch_dir, string desk)
        {
            string[] second = Directory.GetFiles(patch_dir); // путь к папке
            int cont = 0;
            int part_num = 0;

            //// PartRepot or Report

            for (int i = 0; i < second.Length; i++)
            {
                //Console.WriteLine(" - +" + second[i]);
                cont += 1;
                /// тут пишем для парсинга ***PartRepot.pdf
                if (second[i].ToString().EndsWith("PartsReport.pdf"))
                    {

                        

                        string buffer2 = get_text_to_pdf(second[i]);
                        buffer2 = System.Text.RegularExpressions.Regex.Replace(buffer2, @"\s+", " ");

                        ///терерь парсить содержимое
                        string[] stringSeparators = new string[] { "Code: " };
                        string[] words = buffer2.Split(stringSeparators, StringSplitOptions.None);


                        string[] output = new string[12];
                    /// Теперь заполняем карточку для каждого Part
                    for (int j = 1; j < words.Length; j++)
                    {
                        ///Console.WriteLine(j + "---" + words);

                        string[] card_part = words[j].Split(' ', '\n', '\r');
                        for (int a = 0; a < card_part.Length; a++)
                        {
                            Console.WriteLine(a + " - " + card_part[a]);
                            if (card_part[a] == "Machine")
                            {
                                if (a == 2) { output[0] = card_part[a - 1] + card_part[a + 1]; }
                                else if (a == 3) { output[0] = card_part[a - 2]+ card_part[a - 1]; }

                            }
                            if (card_part[a] == "Quantity") {output[1] = card_part[a - 3];}
                            if (card_part[a] == "Creation") { output[2] = card_part[a - 3]; };
                            if (card_part[a] == "Material") { output[3] = card_part[a - 2]; };
                            if (card_part[a] == "Thickness")
                            {
                                double number = Convert.ToDouble(card_part[a - 2], System.Globalization.CultureInfo.InvariantCulture);
                                string result = String.Format("{0:f3}", number);
                                output[4] = result.Replace('.', ',');
                            }
                            if (card_part[a] == "Cuttings") { output[5] = card_part[a - 2].Replace('.', ','); };
                            if (card_part[a] == "Weight(kg)") { output[6] = card_part[a - 2].Replace('.', ','); };
                            if (card_part[a] == "Dimensions") { output[7] = card_part[a - 4] + card_part[a - 3] + card_part[a - 2]; };
                            if (card_part[a] == "Area") { output[8] = card_part[a - 2].Replace('.', ','); };
                            if (card_part[a] == "Length") { output[9] = card_part[a - 1].Replace('.', ','); };
                            output[10] = second[i];
                            output[11] = File.GetCreationTime(second[i]).ToString().Split(' ')[0];
                        }

                        if (checkBox1.Checked == false)
                        {
                            /// если выбран диапазон дат
                            if (dateTimePicker1.Value <= File.GetCreationTime(second[i]) & dateTimePicker2.Value >= File.GetCreationTime(second[i]))
                            {

                                var newLine2 = string.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10};{11}\n", output[0], output[1], output[2],
                                                    output[3], output[4], output[5], output[6], output[7], output[8],
                                                    output[9], output[10], output[11]);

                                Console.WriteLine("--/-" + newLine2);

                                File.AppendAllText(desk + "\\Report.csv", newLine2);
                                
                                

                            }
                        }
                        else
                        {
                            var newLine2 = string.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10};{11}\n", output[0], output[1], output[2],
                                                    output[3], output[4], output[5], output[6], output[7], output[8],
                                                    output[9], output[10], output[11]);

                            File.AppendAllText(desk + "\\Report.csv", newLine2);
                           
                            
                        }

                        
                    }

                    }

            }
            
            richTextBox1.Text += Environment.NewLine + "- Файлаў: " + cont;
            //richTextBox1.Text += Environment.NewLine + "- Дэталяў: " + part_num;

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                groupBox2.Enabled = false;                
            }
            else
            {
                groupBox2.Enabled = true;
            }
        }

        
    }
}
