using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;


using System.IO;
using System.Xml;

using System.IO.Ports;
using System.Management;
using OfficeOpenXml;

namespace AlatUkurSuhu
{
    public partial class Form1 : Form
    {
        string[] ports;
        int portscount = 0;

        bool receiveing = false, closing = false;

        //Add items in the listview
        string[] arr = new string[17];
        ListViewItem itm;

        public Form1()
        {
            InitializeComponent();
        }
        string CommText="Arduino";
        private void Form1_Load(object sender, EventArgs e)
        {
            serialPort1.ReadTimeout = 1000;
            serialPort1.WriteTimeout = 1000;

            timer1.Enabled = true;
            toolStripStatusLabel2.Text = "Alat tidak terhubung";

            comboBox2.SelectedIndex = 0;

            listView1.Columns.Add("Waktu", 50);
            for (int s = 1; s < 16; s++)
            {
                listView1.Columns.Add("S " + s.ToString(),41);
            }


            string path = System.IO.Directory.GetCurrentDirectory() + @"\koneksi.txt";
            if (File.Exists(path))
            {
                try
                {
                    string[] FileContents = System.IO.File.ReadAllLines(path);
                    CommText = FileContents[0];
                }
                catch { }
            }

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            ports = SerialPort.GetPortNames();

            if (ports.Length != portscount)
            {
                portscount = ports.Length;
                comboBox1.Items.Clear();
                comboBox1.Text = "";

                using (var searcher = new ManagementObjectSearcher("SELECT * FROM WIN32_SerialPort"))
                {
                    string[] portnames = SerialPort.GetPortNames();
                    var portss = searcher.Get().Cast<ManagementBaseObject>().ToList();
                    var tList = (from n in portnames join p in portss on n equals p["DeviceID"].ToString() select n + " - " + p["Caption"]).ToList();

                    foreach (string s in tList)
                    {
                        if (s.Contains(CommText))
                        {
                            String[] portName = s.Split(' ');
                            comboBox1.Items.Add(portName[0]);
                            comboBox1.SelectedIndex = 0;
                        }

                    }
                }
                if (comboBox1.Items.Count > 0) {
                    serialPort1.PortName = comboBox1.Text;
                    try
                    {
                        if (serialPort1.IsOpen == false) //if not open, open the port
                        {
                            serialPort1.Open();
                            toolStripStatusLabel2.Text = "Alat telah terhubung";
                            button2.Enabled = true;
                            timer2.Interval = 1000;
                            timer2.Start();
                        }
                    }
                    catch (Exception err) { }
                }
            }
        }

        String SerialBuffer = "";
        int dataCount = -1;
        private void serialPort1_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            receiveing = true;
            SerialPort sp = (SerialPort)sender;
            string indata = sp.ReadExisting();
            SetText(indata);

            dataCount++; if (dataCount > 65535) dataCount = 0;

            if (closing) { serialPort1.Close();  }
        }

        delegate void SetTextCallback(string text);
        private void SetText(string text)
        {
            if (this.textBox1.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(SetText);
                try
                {
                    this.Invoke(d, new object[] { text });
                }
                catch {
                    return;
                }
            }
            else
            {
                if (text.Contains('$'))
                {
                    String tampilkan = SerialBuffer + text;
                    int index = tampilkan.IndexOf("#");
                    tampilkan = tampilkan.Substring(index);
                    this.textBox1.Text = tampilkan;//.Trim('$');
                    SerialBuffer = text;
                }
                else
                    SerialBuffer += text;
                
                textBox1.SelectionStart = textBox1.Text.Length;
                textBox1.ScrollToCaret();

            }
            receiveing = false;

            String[] x = textBox1.Text.Split('$');


       }
        int lastDataCount = 0;
        private void timer2_Tick(object sender, EventArgs e)
        {

            if (closing) { timer2.Enabled = false; this.Close(); return; }
            if (lastDataCount != dataCount)
            {
                lastDataCount = dataCount;

            } else {
                timer2.Enabled = false;
                toolStripStatusLabel2.Text = "Alat tidak terhubung";

                button2.Enabled = false;
                timer4.Enabled = false;

                MessageBox.Show("Sambungan ke alat terputus. Jalankan program kembali untuk pengambilan data!", "Perhatian", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            closing = true;
            if (receiveing) { e.Cancel = true; return; }
            timer1.Enabled = false; timer2.Enabled = false;
            try
            {
                serialPort1.Close();
                serialPort1.Dispose();
            }
            catch { }
        }

        private void button1_Click(object sender, EventArgs e)
        {

            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Excel 2007 |*.xlsx";
            saveFileDialog1.Title = "Save to Excel File";
            saveFileDialog1.ShowDialog();

            if (saveFileDialog1.FileName != "")
            {
                switch (saveFileDialog1.FilterIndex)
                {
                    case 1:
                        FileInfo newFile = new FileInfo(saveFileDialog1.FileName);
                        if (newFile.Exists)
                        {
                            try
                            {
                                newFile.Delete();  // ensures we create a new workbook
                                newFile = new FileInfo(saveFileDialog1.FileName);
                            }
                            catch {
                                MessageBox.Show("Pembuatan file gagal. Pastikan file tidak digunakan");
                                return;
                            }
                        }

                        using (ExcelPackage xlPackage = new ExcelPackage(newFile))
                        {
                            xlPackage.DebugMode = true;

                            ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets.Add("Tinned Goods");

                            //Title
                            worksheet.Cell(1, 1).Value = "Waktu";
                            for (int s = 1; s < 16; s++)
                            {
                                worksheet.Cell(1, s+1).Value = "S "+s.ToString();
                                worksheet.Column(s+1).Width = 5;
                            }

                            xxx = new string[listView1.Items.Count, 16];
                            for (int y = 0; y < listView1.Items.Count; y++)
                            {
                                for (int x = 0; x < 16; x++)
                                {
                                    //xxx[y, x] = listView1.Items[y].SubItems[x].Text;

                                    //y = 1, title
                                    worksheet.Cell(y+2, x+1).Value = listView1.Items[y].SubItems[x].Text;
                                }
                            }


                            // lets set the header text 
                            worksheet.HeaderFooter.oddHeader.CenteredText = "Rakapitulasi Suhu Beton";
 
                            // add the page number to the footer plus the total number of pages
                            worksheet.HeaderFooter.oddFooter.RightAlignedText =
                                string.Format("Halaman {0} dari {1}", ExcelHeaderFooter.PageNumber, ExcelHeaderFooter.NumberOfPages);
                            
                            // add the sheet name to the footer
                            worksheet.HeaderFooter.oddFooter.CenteredText = ExcelHeaderFooter.SheetName;
                            
                            // add the file path to the footer
                            worksheet.HeaderFooter.oddFooter.LeftAlignedText = ExcelHeaderFooter.FilePath + ExcelHeaderFooter.FileName;

                            // change the sheet view to show it in page layout mode
                            worksheet.View.PageLayoutView = true;

                            // set some core property values
                            xlPackage.Workbook.Properties.Title = "Rakapitulasi Suhu Beton";
                            xlPackage.Workbook.Properties.Author = "Ristiana Dyah";
                            xlPackage.Workbook.Properties.Subject = "Lembar Rakapitulasi Suhu Beton";
                            xlPackage.Workbook.Properties.Keywords = "Suhu Beton";
                            xlPackage.Workbook.Properties.Category = "Laporan Rekapitulasi";
                            xlPackage.Workbook.Properties.Comments = "Rakapitulasi Suhu Beton dari pengukuran alat pengukur suhu beton";

                            // set some extended property values
                            xlPackage.Workbook.Properties.Company = "Universitas Muhammadiyah Purwokerto";
                            xlPackage.Workbook.Properties.HyperlinkBase = new Uri("http://www.ump.ac.id");

                            // set some custom property values
                            xlPackage.Workbook.Properties.SetCustomPropertyValue("Checked by", "Ristiana Dyah");
                            xlPackage.Workbook.Properties.SetCustomPropertyValue("EmployeeID", "-");
                            xlPackage.Workbook.Properties.SetCustomPropertyValue("AssemblyName", "Excel Package");

                            // save our new workbook and we are done!
                            xlPackage.Save();

                        }
                        break;

                }

            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            timer3.Interval = int.Parse(comboBox2.Text) * 1000;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (button2.Text == "Mulai")
            {
                timer3.Enabled = true;
                button2.Text = "Berhenti";
                timer4.Enabled = true;
            }
            else {
                timer3.Enabled = false;
                button2.Text = "Mulai";
                timer4.Enabled = false;

            }
        }
        string[,] xxx;
        private void timer3_Tick(object sender, EventArgs e)
        {
            String jam = DateTime.Now.Hour.ToString();
            String mnt = DateTime.Now.Minute.ToString();
            String dtk = DateTime.Now.Second.ToString();
            
            String parsing = textBox1.Text;
            String[] parsingList = parsing.Split('\n');

            for (int x = 0; x < parsingList.Length; x++) {
                if (x >= 16) continue;
                if (parsingList[x].Length > 2)
                {
                    int sharpParsing = int.Parse(parsingList[x].Substring(1, 2));

                    arr[sharpParsing] = parsingList[x].Substring(4,5);//parsingList[x];
                }
            }

            for (int x = 0; x < arr.Length; x++)
            {
                if (arr[x] == null) arr[x] = "0";
            }

            //Add first item
            arr[0] = jam + ":" + mnt + ":"+dtk;
            itm = new ListViewItem(arr);
            listView1.Items.Add(itm);

            //Auto scroll
            if (checkBox1.Checked) { listView1.EnsureVisible(listView1.Items.Count - 1); }

            // usage

            
            

       }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Yakin akan menghapus semua data pada tabel rekapitulasi?", "Perhatian", MessageBoxButtons.YesNo,MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes)
            {
                //do something
                listView1.Items.Clear();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        private void timer4_Tick(object sender, EventArgs e)
        {
            if (toolStripProgressBar1.Value < 10)
                toolStripProgressBar1.Value++;
            else
                toolStripProgressBar1.Value = 0;
        }

        
    }
}
