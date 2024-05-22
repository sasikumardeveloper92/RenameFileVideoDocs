using AxAcroPDFLib;
using AxAXVLC;
using AxWMPLib;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace RenameFile
{
    public partial class Form2 : Form
    {
        public Dictionary<int, string> SongsList = new Dictionary<int, string>();
        public int CurrentFileIndex = 0;
        public int PreviousFileIndex; 
            static readonly string[] SizeSuffixes = {
      "bytes",
          "KB",
          "MB",
          "GB",
          "TB",
          "PB",
          "EB",
          "ZB",
          "YB"
    };
        public Form2()
        {
            InitializeComponent();
            LoadControlsData();
        }
        private void LoadControlsData()
        {
            List<Category> catogeries = new List<Category>();
            foreach (string key in ConfigurationManager.AppSettings)
            {
                if (!key.Contains("SupportedFormat"))
                {
                    if (Radbtn_Docs.Checked)
                    {
                        if (key.StartsWith("d"))
                        {
                            string value = ConfigurationManager.AppSettings[key];
                            string[] parts = value.Split(',');
                            if (parts.Length == 2)
                            {
                                catogeries.Add(new Category
                                {
                                    CatogeryName = parts[0],
                                    Id = int.Parse(parts[1])
                                });

                            }

                        }
                    }

                    if (Radbtn_Video.Checked)
                    {
                        if (key.StartsWith("v"))
                        {
                            string value = ConfigurationManager.AppSettings[key];
                            string[] parts = value.Split(',');
                            if (parts.Length == 2)
                            {
                                catogeries.Add(new Category
                                {
                                    CatogeryName = parts[0],
                                    Id = int.Parse(parts[1])
                                });

                            }

                        }
                        axWindowsMediaPlayer1.Visible = true;
                        axAcroPDF1.Visible = false;
                    }

                    
                }

            }
            var Sorted = catogeries.OrderBy(x => x.Id);
            listBox1.Items.Clear();
            foreach (var item in Sorted)
            {
                listBox1.Items.Add(item.CatogeryName);
            }

        }



        

        private string GetRenameText()
        {
            string resultext = "";
            foreach (var item in listBox1.SelectedItems)
            {
                if (resultext == "" )
                {
                    resultext = item.ToString();
                }
                else
                {
                    resultext = resultext + "_" + item.ToString();
                }
                if (listBox1.SelectedItems.Count > 0)
                {
                    resultext = resultext + "_";
                }
                
            }
            return resultext;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            using (var folderBrowsedialog = new FolderBrowserDialog())
            {
                if (folderBrowsedialog.ShowDialog() == DialogResult.OK)
                {
                    string directionpath = folderBrowsedialog.SelectedPath;
                    textBox1.Text = directionpath;
                    PopulateFile(false);
                }
            }
        }
        private void ClearValue()
        {
            dataGridView1.Rows.Clear();
            CurrentFileIndex = 0;
            PreviousFileIndex = 0;
            //webBrowser1.Navigate ("about:blank");
            string url = "";
            axWindowsMediaPlayer1.URL = url;
            //webBrowser1.Refresh();
            dataGridView1.DataSource = null;
        }
        private void PopulateFile(bool filtered)
        {
            string path = textBox1.Text;
            string[] ListFiles = Directory.GetFiles(path);
            dataGridView1.Rows.Clear();
            if (filtered)
            {
                List<string> selectedCategories = listBox1.SelectedItems.Cast<string>().ToList();
                List<string> filteredFiles = new List<string>();
                if (selectedCategories.Any())
                {
                    string[] files = Directory.GetFiles(textBox1.Text);
                    foreach (string file in files)
                    {
                        foreach (var Category in selectedCategories)
                        {
                            if (Path.GetFileName(file).Contains(Category))
                            {
                                filteredFiles.Add(file);
                                break;
                            }
                        }
                    }
                }
                dataGridView1.Rows.Clear();
                ListFiles = filteredFiles.ToArray();
            }
            int i = 0;
            string Videovalue = ConfigurationManager.AppSettings["VideoSupportFormat"];
            string[] Videoextenstion = Videovalue.Split(',');
            string Docsvalue = ConfigurationManager.AppSettings["DocsSupportedFormat"];
            string[] Docsextenstion = Docsvalue.Split(',');
            if (!dataGridView1.Columns.Cast<DataGridViewColumn>().Any(col => col.Name == "Name"))
            {
                dataGridView1.Columns.Add("ID", "ID");
                dataGridView1.Columns.Add("Name", "FileName");
                dataGridView1.Columns.Add("Type", "Extension");
                dataGridView1.Columns.Add("Last DateModified", "DateModified");
                dataGridView1.Columns.Add("Size", "Size");
            }
            try
            {
                foreach (var file in ListFiles)
                {
                    if (Radbtn_Docs.Checked)
                    {
                        if (Array.Exists(Docsextenstion, ext => ext.Equals(Path.GetExtension(file), StringComparison.OrdinalIgnoreCase)))
                        {
                            i++;
                            FileInfo fileInfo = new FileInfo(file);
                            string filesize = SizeSuffix(fileInfo.Length);
                            dataGridView1.Rows.Add(i, fileInfo.Name, fileInfo.Extension, fileInfo.LastWriteTime, filesize);
                        }
                    }
                    if (Radbtn_Video.Checked)
                    {
                        if (Array.Exists(Videoextenstion, ext => ext.Equals(Path.GetExtension(file), StringComparison.OrdinalIgnoreCase)))
                        {
                            i++;
                            FileInfo fileInfo = new FileInfo(file);
                            string filesize = SizeSuffix(fileInfo.Length);
                            dataGridView1.Rows.Add(i, fileInfo.Name, fileInfo.Extension, fileInfo.LastWriteTime, filesize);
                        }
                    }
                }
                dataGridView1.Columns["ID"].Visible = false;
                var parentpath = textBox1.Text;
                CurrentFileIndex = 0;
                if (dataGridView1.Rows.Count > 0)
                {
                    dataGridView1.Rows[0].Selected = true;
                    var Name = dataGridView1.Rows[CurrentFileIndex].Cells[1].Value;
                    if (Radbtn_Video.Checked)
                    {
                        string url = parentpath + "\\" + Convert.ToString(Name);
                        axWindowsMediaPlayer1.URL = url;
                    }
                    if (Radbtn_Docs.Checked)
                    {
                        axAcroPDF1.src = parentpath + "/" + Convert.ToString(Name);
                    }
                }
                else
                {
                    MessageBox.Show("No Files!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void PlaySong()
        {
            var parentpath = textBox1.Text;
            if (dataGridView1.CurrentRow != null)
            {
                CurrentFileIndex = dataGridView1.CurrentRow.Index;
            }
            dataGridView1.ClearSelection();
            if (CurrentFileIndex < dataGridView1.RowCount - 1)
            {
                int NextIndex = CurrentFileIndex + 1;
                dataGridView1.Rows[NextIndex].Selected = true;
                PreviousFileIndex = CurrentFileIndex;
                CurrentFileIndex = NextIndex;
            }
            else
            {
                dataGridView1.Rows[0].Selected = true;
                PreviousFileIndex = dataGridView1.RowCount - 1;
                CurrentFileIndex = 0;
            }
            var Name = dataGridView1.Rows[CurrentFileIndex].Cells[1].Value;
            var PrevName = dataGridView1.Rows[PreviousFileIndex].Cells[1].Value;
            var rename = GetRenameText();
            if (rename != "")
            {
                string Source = textBox1.Text + "\\" + Convert.ToString(PrevName);
                string destFileName = textBox1.Text + "\\" + rename + Convert.ToString(PrevName);
                File.Move(Source, destFileName);
                dataGridView1.Rows[PreviousFileIndex].Cells[1].Value = rename + Convert.ToString(PrevName);
            }
            string url = parentpath + "\\" + Convert.ToString(Name);
            axWindowsMediaPlayer1.URL = url;
        }
        private void ViewReport()
        {
            var parentpath = textBox1.Text;
            if (dataGridView1.CurrentRow != null)
            {
                CurrentFileIndex = dataGridView1.CurrentRow.Index;
            }
            dataGridView1.ClearSelection();
            if (CurrentFileIndex < dataGridView1.RowCount - 1)
            {
                int NextIndex = CurrentFileIndex + 1;

                dataGridView1.Rows[NextIndex].Selected = true;
                PreviousFileIndex = CurrentFileIndex;
                CurrentFileIndex = NextIndex;
            }
            else
            {
                dataGridView1.Rows[0].Selected = true;
                PreviousFileIndex = dataGridView1.RowCount - 1;
                CurrentFileIndex = 0;
            }
            var Name = dataGridView1.Rows[CurrentFileIndex].Cells[1].Value;
            var PrevName = dataGridView1.Rows[PreviousFileIndex].Cells[1].Value;
            var rename = GetRenameText();
            if (rename != "")
            {
                string Source = textBox1.Text + "\\" + Convert.ToString(PrevName);
                string destFileName = textBox1.Text + "\\" + rename + Convert.ToString(PrevName);
                File.Move(Source, destFileName);
                dataGridView1.Rows[PreviousFileIndex].Cells[1].Value = rename + Convert.ToString(PrevName);
            }
            string Filename = Convert.ToString(Name);
            if (Filename.Contains(".pdf"))
            {
                axAcroPDF1.src = parentpath + "/" + Convert.ToString(Name);
            }
        }
        static string SizeSuffix(Int64 value)
        {
            if (value < 0)
            {
                return "-" + SizeSuffix(-value);
            }
            int i = 0;
            decimal dValue = (decimal)value;
            while (Math.Round(dValue / 1024) >= 1)
            {
                dValue /= 1024;
                i++;
            }
            return string.Format("{0:1} {1}", dValue, SizeSuffixes[i]);
        }
        private async void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.Rows.Count > 0)
                {
                    dataGridView1.ClearSelection();
                    if (Radbtn_Video.Checked) {
                        PlaySong();
                    }
                    if (Radbtn_Docs.Checked)
                    {
                        ViewReport();
                    }
                    listBox1.SelectedItems.Clear();
                }
                else
                {
                    MessageBox.Show("No Files!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Close the App. Try again.");
            }
        }
        private async Task RenameFile(string Source, string destFileName)
        {
            if (FileAvailable(Source))
            {
                File.Move(Source, destFileName);
            }
            else
            {
                MessageBox.Show("File Already Open!. Try again.");
            }
        }
        private bool FileAvailable(string filepath)
        {
            bool fileAvailable = false;
            int maxretries = 200;
            int retryDelayms = 500;
            for (int retrycount = 0; retrycount < maxretries; retrycount++)
            {
                try
                {
                    using (FileStream fs = File.Open(filepath, FileMode.Open, FileAccess.Read))
                    {
                        fileAvailable = true;
                        break;
                    }
                }
                catch (IOException)
                {
                    Thread.Sleep(retryDelayms);
                }
                
            }
            return fileAvailable;

        }
            private void radioButton1_CheckedChanged(object sender, EventArgs e)
            {
                if (Radbtn_Video.Checked == true)
                {
                axWindowsMediaPlayer1.Visible = true;
                    axAcroPDF1.Visible = false;
                    textBox1.Text = "";
                    dataGridView1.Rows.Clear();
                    LoadControlsData();
                }
                else
                {
                axWindowsMediaPlayer1.Visible = false;
                    axAcroPDF1.Visible = true;
                }
            }
            private void radioButton2_CheckedChanged(object sender, EventArgs e)
            {
                if (Radbtn_Docs.Checked == true)
                {
                    textBox1.Text = "";
                    dataGridView1.Rows.Clear();
                axWindowsMediaPlayer1.Visible = false;
                    LoadControlsData();
                    axAcroPDF1.Visible = true;
                }
                else
                {
                axWindowsMediaPlayer1.Visible = true;
                    axAcroPDF1.Visible = false;
                }
            }
            private void button5_Click(object sender, EventArgs e)
            {
                PopulateFile(true);
            }

        
    }

}