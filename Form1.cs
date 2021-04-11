using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace WindowsFormsApp1
{
    public partial class Form1 : MetroFramework.Forms.MetroForm
    {

        public Form1()
        {
            InitializeComponent();
        }
        enum mode
        {
            Students = 0X0,
            Teachers = 0X1
        }
        int _switch = 0;
        string[] carr = { };
        private static Random random = new Random();
        public static string RandomString(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            return new string(Enumerable.Repeat(chars, length)
              .Select(s => s[random.Next(s.Length)]).ToArray());
        }
   
        void cleancache()
        {
            try
            {
                foreach (var process in Process.GetProcessesByName("EXCEL"))
                    process.Kill();
                String dtPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Cache\\";
                DirectoryInfo dir = new DirectoryInfo(dtPath);
                foreach (FileInfo flInfo in dir.GetFiles())
                {
                    File.Delete(flInfo.FullName);
                }
            }
            catch (Exception ex) { };

        }
        void fixexcel(string foldername,string savename)
        {
            cleancache();
            String dtPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\"+ foldername+ "\\";
            DirectoryInfo dir = new DirectoryInfo(dtPath);
            foreach (FileInfo flInfo in dir.GetFiles())
            {

                //Instantiate a Workbook object
                Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
                //Load the Excel file
                workbook.LoadFromFile(flInfo.FullName);

                //Get the first worksheet
                foreach (Spire.Xls.Worksheet sheet in workbook.Worksheets)
                {
                    //Delete blanks rows 
                    for (int i = sheet.Rows.Count() - 1; i >= 0; i--)
                    {
                        if (sheet.Rows[i].IsBlank)
                        {
                            sheet.DeleteRow(i + 1); //Index parameter in DeleteRow method starts from 1
                        }
                    }

                    //Delete blank columns
                    for (int j = sheet.Columns.Count() - 1; j >= 0; j--)
                    {
                        if (sheet.Columns[j].IsBlank)
                        {
                            sheet.DeleteColumn(j + 1); //Index parameter in DeleteColumn method starts from 1
                        }
                    }
                }
                //Save the file
                workbook.SaveToFile(savename + ".xlsx", Spire.Xls.ExcelVersion.Version2013);
            }
        }
        void WriteExcel()
        {
            try
            {
                String dtPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Cache\\save\\_cache.xlsx";
                var excelApp = new Microsoft.Office.Interop.Excel.Application();
                var workbook = excelApp.Workbooks.Open(dtPath);
                foreach (dynamic worksheet in workbook.Worksheets)
                {
                    worksheet.Cells.ClearContents();

                    int i = 1;
                    int i2 = 1;
                    foreach (ListViewItem lvi in listView1.Items)
                    {
                        i = 1;
                        foreach (ListViewItem.ListViewSubItem lvs in lvi.SubItems)
                        {
                            worksheet.Cells[i2, i] = lvs.Text;
                            i++;
                        }
                        i2++;
                    }
                }
                excelApp.Visible = false;
                workbook.Save();
                GC.Collect();
                GC.WaitForPendingFinalizers();
              //  Marshal.FinalReleaseComObject(worksheet);
                workbook.Close();
                Marshal.FinalReleaseComObject(workbook);
                excelApp.Quit();
                Marshal.FinalReleaseComObject(excelApp);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        void Randomize(ListView lv)
        {
            ListView.ListViewItemCollection list = lv.Items;
            Random rng = new Random();
            int n = list.Count;
            lv.BeginUpdate();
            while (n > 1)
            {
                n--;
                int k = rng.Next(n + 1);
                ListViewItem value1 = (ListViewItem)list[k];
                ListViewItem value2 = (ListViewItem)list[n];
                list[k] = new ListViewItem();
                list[n] = new ListViewItem();
                list[k] = value2;
                list[n] = value1;
            }
            lv.EndUpdate();
            lv.Invalidate();
        }
        void ConvertExcel2PDF(string path)
        {
        }
        void AddRow(string name, string Section, string Stage, string studying ,string halls)
        {
            string[] row = { name.ToString(), Section.ToString(), Stage.ToString(), studying.ToString(), halls.ToString() };
            var listViewItem = new ListViewItem(row);
            listView1.Items.Add(listViewItem);
        }
        void Load_Students()
        {
             try
            {
                listView1.Columns.Clear();

                string[] LVColumns = File.ReadAllLines("Config\\columns.ini");
                foreach (var line in LVColumns)
                {
                    string[] tokens = line.Split(',');
                    var myList = new List<string>() { tokens[0] };
                    myList.ForEach(x => listView1.Columns.Add(x));

                    c1ToolStripMenuItem.Text = LVColumns[0];
                    c2ToolStripMenuItem.Text = LVColumns[1];
                    c3ToolStripMenuItem.Text = LVColumns[2];
                    c4ToolStripMenuItem.Text = LVColumns[3];
                    c5ToolStripMenuItem.Text = LVColumns[4];

                    dc1ToolStripMenuItem.Text = LVColumns[1];
                    dc2ToolStripMenuItem.Text = LVColumns[2];
                    dc3ToolStripMenuItem.Text = LVColumns[3];
                    dc4ToolStripMenuItem.Text = LVColumns[4];

                    loadexcelToolStripMenuItem.Text = "نحميل " + LVColumns[0];
                    addhallToolStripMenuItem.Text = "تعديل " + LVColumns[4];
                    hallsToolStripMenuItem.Text = "توزيع الى " + LVColumns[4];

                    metroLabel1.Text = LVColumns[0];
                    metroLabel2.Text = LVColumns[1];
                    metroLabel3.Text = LVColumns[2];
                    metroLabel4.Text = LVColumns[3];
                    metroLabel6.Text = LVColumns[4];
                    metroLabel5.Text = LVColumns[4] + " افتراضية";
 
                    carr = LVColumns;
                }

                if (listView1.Columns[0].Text != "0")
                {
                    listView1.Columns[0].Width = 150;
                    metroLabel1.Enabled = true;
                    metroLabel1.Visible = true;
                    textboxname.Enabled = true;
                    textboxname.Visible = true;
                    c1ToolStripMenuItem.Checked = true;
                    c1ToolStripMenuItem.Visible = true;
                }
                else
                {
                    listView1.Columns[0].Width = 0;
                    metroLabel1.Enabled = false;
                    metroLabel1.Visible = false;
                    textboxname.Enabled = false;
                    textboxname.Visible = false;
                    c1ToolStripMenuItem.Checked = false;
                    c1ToolStripMenuItem.Visible = false;

                }

                if (listView1.Columns[1].Text != "0")
                {
                    listView1.Columns[1].Width = 150;
                    c2ToolStripMenuItem.Enabled = true;
                    c2ToolStripMenuItem.Visible = true;
                    metroLabel2.Enabled = true;
                    metroLabel2.Visible = true;
                    textboxsection.Enabled = true;
                    textboxsection.Visible = true;
                    dc1ToolStripMenuItem.Enabled = true;
                    dc1ToolStripMenuItem.Visible = true;
                }
                else
                {
                    listView1.Columns[1].Width = 0;
                    c2ToolStripMenuItem.Enabled = false;
                    c2ToolStripMenuItem.Visible = false;
                    metroLabel2.Enabled = false;
                    metroLabel2.Visible = false;
                    textboxsection.Enabled = false;
                    textboxsection.Visible = false;
                    dc1ToolStripMenuItem.Enabled = false;
                    dc1ToolStripMenuItem.Visible = false;
                }

                if (listView1.Columns[2].Text != "0")
                {
                    listView1.Columns[2].Width = 150;
                    c3ToolStripMenuItem.Enabled = true;
                    c3ToolStripMenuItem.Visible = true;
                    metroLabel3.Enabled = true;
                    metroLabel3.Visible = true;
                    textboxstage.Enabled = true;
                    textboxstage.Visible = true;
                    dc2ToolStripMenuItem.Enabled = true;
                    dc2ToolStripMenuItem.Visible = true;
                }
                else
                {
                    listView1.Columns[2].Width = 0;
                    c3ToolStripMenuItem.Enabled = false;
                    c3ToolStripMenuItem.Visible = false;
                    metroLabel3.Enabled = false;
                    metroLabel3.Visible = false;
                    textboxstage.Enabled = false;
                    textboxstage.Visible = false;
                    dc2ToolStripMenuItem.Enabled = false;
                    dc2ToolStripMenuItem.Visible = false;
                }

                if (listView1.Columns[3].Text != "0")
                {
                    listView1.Columns[3].Width = 150;
                    c4ToolStripMenuItem.Enabled = true;
                    c4ToolStripMenuItem.Visible = true;
                    metroLabel4.Enabled = true;
                    metroLabel4.Visible = true;
                    textboxstudying.Enabled = true;
                    textboxstudying.Visible = true;
                    dc3ToolStripMenuItem.Enabled = true;
                    dc3ToolStripMenuItem.Visible = true;
                }
                else
                {
                    listView1.Columns[3].Width = 0;
                    c4ToolStripMenuItem.Enabled = false;
                    c4ToolStripMenuItem.Visible = false;
                    metroLabel4.Enabled = false;
                    metroLabel4.Visible = false;
                    textboxstudying.Enabled = false;
                    textboxstudying.Visible = false;

                    dc3ToolStripMenuItem.Enabled = false;
                    dc3ToolStripMenuItem.Visible = false;
                }

                if (listView1.Columns[4].Text != "0")
                {
                    listView1.Columns[4].Width = 150;
                    c5ToolStripMenuItem.Enabled = true;
                    c5ToolStripMenuItem.Visible = true;
                    metroLabel6.Enabled = true;
                    metroLabel5.Enabled = true;
                    metroLabel6.Visible = true;
                    metroLabel5.Visible = true;
                    metroTextBox2.Enabled = true;
                    metroTextBox2.Visible = true;
                    dc4ToolStripMenuItem.Enabled = true;
                    dc4ToolStripMenuItem.Visible = true;
                }
                else
                {
                    listView1.Columns[4].Width = 0;
                    c5ToolStripMenuItem.Enabled = false;
                    c5ToolStripMenuItem.Visible = false;
                    metroLabel6.Enabled = false;
                    metroLabel5.Enabled = false;
                    metroLabel6.Visible = false;
                    metroLabel5.Visible = false;
                    metroTextBox2.Enabled = false;
                    metroTextBox2.Visible = false;
                    dc4ToolStripMenuItem.Visible = false;
                }
                if ((int)mode.Students == _switch)
                {
                    metroLabel8.Enabled = false;
                    numericUpDown2.Enabled = false;
                }
                listView1.Refresh();

                cleancache();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }
        void Load_teachers()
        {
            try
            {
                listView1.Columns.Clear();

                string[] LVColumns = File.ReadAllLines("Config\\t_columns.ini");
                foreach (var line in LVColumns)
                {
                    string[] tokens = line.Split(',');
                    var myList = new List<string>() { tokens[0] };
                    myList.ForEach(x => listView1.Columns.Add(x));

                    c1ToolStripMenuItem.Text = LVColumns[0];
                    c2ToolStripMenuItem.Text = LVColumns[1];
                    c3ToolStripMenuItem.Text = LVColumns[2];
                    c4ToolStripMenuItem.Text = LVColumns[3];
                    c5ToolStripMenuItem.Text = LVColumns[4];

                    dc1ToolStripMenuItem.Text = LVColumns[1];
                    dc2ToolStripMenuItem.Text = LVColumns[2];
                    dc3ToolStripMenuItem.Text = LVColumns[3];
                    dc4ToolStripMenuItem.Text = LVColumns[4];
                    loadexcelToolStripMenuItem.Text = "نحميل " + LVColumns[0] ;
                    addhallToolStripMenuItem.Text = "تعديل " + LVColumns[4];
                    hallsToolStripMenuItem.Text = "توزيع الى " + LVColumns[4];

                    metroLabel1.Text = LVColumns[0];
                    metroLabel2.Text = LVColumns[1];
                    metroLabel3.Text = LVColumns[2];
                    metroLabel4.Text = LVColumns[3];
                    metroLabel6.Text = LVColumns[4];
                    metroLabel5.Text = LVColumns[4] + " افتراضية";

  
                    carr = LVColumns;
                }

                if (listView1.Columns[0].Text != "0")
                {
                    listView1.Columns[0].Width = 150;
                    metroLabel1.Enabled = true;
                    metroLabel1.Visible = true;
                    textboxname.Enabled = true;
                    textboxname.Visible = true;
                    c1ToolStripMenuItem.Checked = true;
                    c1ToolStripMenuItem.Visible = true;
                }
                else
                {
                    listView1.Columns[0].Width = 0;
                    metroLabel1.Enabled = false;
                    metroLabel1.Visible = false;
                    textboxname.Enabled = false;
                    textboxname.Visible = false;
                    c1ToolStripMenuItem.Checked = false;
                    c1ToolStripMenuItem.Visible = false;

                }

                if (listView1.Columns[1].Text != "0")
                {
                    listView1.Columns[1].Width = 150;
                    c2ToolStripMenuItem.Enabled = true;
                    c2ToolStripMenuItem.Visible = true;
                    metroLabel2.Enabled = true;
                    metroLabel2.Visible = true;
                    textboxsection.Enabled = true;
                    textboxsection.Visible = true;
                    dc1ToolStripMenuItem.Enabled = true;
                    dc1ToolStripMenuItem.Visible = true;
                }
                else
                {
                    listView1.Columns[1].Width = 0;
                    c2ToolStripMenuItem.Enabled = false;
                    c2ToolStripMenuItem.Visible = false;
                    metroLabel2.Enabled = false;
                    metroLabel2.Visible = false;
                    textboxsection.Enabled = false;
                    textboxsection.Visible = false;
                    dc1ToolStripMenuItem.Enabled = false;
                    dc1ToolStripMenuItem.Visible = false;
                }

                if (listView1.Columns[2].Text != "0")
                {
                    listView1.Columns[2].Width = 150;
                    c3ToolStripMenuItem.Enabled = true;
                    c3ToolStripMenuItem.Visible = true;
                    metroLabel3.Enabled = true;
                    metroLabel3.Visible = true;
                    textboxstage.Enabled = true;
                    textboxstage.Visible = true;
                    dc2ToolStripMenuItem.Enabled = true;
                    dc2ToolStripMenuItem.Visible = true;
                }
                else
                {
                    listView1.Columns[2].Width = 0;
                    c3ToolStripMenuItem.Enabled = false;
                    c3ToolStripMenuItem.Visible = false;
                    metroLabel3.Enabled = false;
                    metroLabel3.Visible = false;
                    textboxstage.Enabled = false;
                    textboxstage.Visible = false;
                    dc2ToolStripMenuItem.Enabled = false;
                    dc2ToolStripMenuItem.Visible = false;
                }

                if (listView1.Columns[3].Text != "0")
                {
                    listView1.Columns[3].Width = 150;
                    c4ToolStripMenuItem.Enabled = true;
                    c4ToolStripMenuItem.Visible = true;
                    metroLabel4.Enabled = true;
                    metroLabel4.Visible = true;
                    textboxstudying.Enabled = true;
                    textboxstudying.Visible = true;
                    dc3ToolStripMenuItem.Enabled = true;
                    dc3ToolStripMenuItem.Visible = true;
                }
                else
                {
                    listView1.Columns[3].Width = 0;
                    c4ToolStripMenuItem.Enabled = false;
                    c4ToolStripMenuItem.Visible = false;
                    metroLabel4.Enabled = false;
                    metroLabel4.Visible = false;
                    textboxstudying.Enabled = false;
                    textboxstudying.Visible = false;

                    dc3ToolStripMenuItem.Enabled = false;
                    dc3ToolStripMenuItem.Visible = false;
                }

                if (listView1.Columns[4].Text != "0")
                {
                    listView1.Columns[4].Width = 150;
                    c5ToolStripMenuItem.Enabled = true;
                    c5ToolStripMenuItem.Visible = true;
                    metroLabel6.Enabled = true;
                    metroLabel5.Enabled = true;
                    metroLabel6.Visible = true;
                    metroLabel5.Visible = true;
                    metroTextBox2.Enabled = true;
                    metroTextBox2.Visible = true;
                    dc4ToolStripMenuItem.Enabled = true;
                    dc4ToolStripMenuItem.Visible = true;
                }
                else
                {
                    listView1.Columns[4].Width = 0;
                    c5ToolStripMenuItem.Enabled = false;
                    c5ToolStripMenuItem.Visible = false;
                    metroLabel6.Enabled = false;
                    metroLabel5.Enabled = false;
                    metroLabel6.Visible = false;
                    metroLabel5.Visible = false;
                    metroTextBox2.Enabled = false;
                    metroTextBox2.Visible = false;
                    dc4ToolStripMenuItem.Visible = false;

                }
                if ((int)mode.Teachers == _switch)
                {
                    metroLabel8.Enabled = true;
                    numericUpDown2.Enabled = true;
                }
                listView1.Refresh();

                cleancache();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());

            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            Load_Students();
        }
        private void metroButton1_Click(object sender, EventArgs e)
        {
            AddRow(textboxname.Text, textboxsection.Text, textboxstage.Text, textboxstudying.Text, metroTextBox2.Text);
            groupBox1.Text = listView1.Items.Count + " بيانات في لستة";

        }
        string[] ConvertToStringArray(Array values)
        {
            //create a new string array 
            string[] arrays = new string[values.Length];
            for (int i = 1; i <= values.Length; i++)
            {
                if (values.GetValue(1, i) == null)
                    arrays[i - 1] = "";
                else
                    arrays[i - 1] = (string)values.GetValue(1, i).ToString();
            }
            return arrays;
        }

        private void oppToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in listView1.SelectedItems)
            {
                listView1.Items.Remove(item);
            }
            groupBox1.Text = listView1.Items.Count + " بيانات في لستة";

        }

        private void addToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string name = "";
            string section = "";
            string stage = "";
            string studying = "";
            string halls = "";
            if(carr[0] !="0")
            name = Interaction.InputBox(carr[0], "اضافة عنصر", "");
            if (carr[1] != "0")
                section = Interaction.InputBox(carr[1], "اضافة عنصر", "");
            if (carr[2] != "0")
                stage = Interaction.InputBox(carr[2], "اضافة عنصر", "");
            if (carr[3] != "0")
                studying = Interaction.InputBox(carr[3], "اضافة عنصر", "");
            if (carr[4] != "0")
                halls = Interaction.InputBox(carr[4], "اضافة عنصر", "");

            AddRow(name, section, stage, studying, halls);
            groupBox1.Text = listView1.Items.Count + " بيانات في لستة";

        }

        private void editeToolStripMenuItem_Click(object sender, EventArgs e)
        {

            string name = "";
            string section = "";
            string stage = "";
            string studying = "";
            if (carr[0] != "0")
                name = Interaction.InputBox(carr[0], "تعديل عنصر", "");
            if (carr[1] != "0")
                section = Interaction.InputBox(carr[1], "تعديل عنصر", "");
            if (carr[2] != "0")
                stage = Interaction.InputBox(carr[2], "تعديل عنصر", "");
            if (carr[3] != "0")
                studying = Interaction.InputBox(carr[3], "تعديل عنصر", "");

            foreach (ListViewItem item in listView1.SelectedItems)
            {
                item.Text = name;
                item.SubItems[1].Text = section;
                item.SubItems[2].Text = stage;
                item.SubItems[3].Text = studying;

            }
        }

        private void loadexcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if((int)mode.Students == _switch)
                fixexcel("Students", "Cache//Temp" + RandomString(9));
                else
                fixexcel("Teachers", "Cache//Temp" + RandomString(9));
                Thread.Sleep(1000);
                foreach (var process in Process.GetProcessesByName("EXCEL"))
                {
                    process.Kill();
                }
                int bypass = (int)numericUpDown1.Value;
                bypass += 1;
                int nstat = 0;
                string init = "";
                string total = "";
                init = Interaction.InputBox("لتحميل جميع بيانات اكتب 1 لتحميل عدد بيانات اكتب 2 ", "تحميل الطلاب", "");
                if (init != "1")
                {
                    total = Interaction.InputBox("تحميل عدد بيانات لكل ملف او ورقة", "عدد بيانات", "");
                }
                String dtPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Cache\\";
                DirectoryInfo dir = new DirectoryInfo(dtPath);
                foreach (FileInfo flInfo in dir.GetFiles())
                {
                    var excelApp = new Microsoft.Office.Interop.Excel.Application();
                    var workbook = excelApp.Workbooks.Open(flInfo.FullName);
                    var sheets = workbook.Worksheets;
                    //Worksheet worksheet = (Worksheet)sheets.get_Item(1);
               
                    foreach (Worksheet ws in workbook.Worksheets)
                    {
                        // var usedRange = worksheet.UsedRange;
                        //  usedRange.RemoveDuplicates(1);
                        if (init == "1")
                        {
                            nstat = ws.UsedRange.Rows.Count;
                        }
                        else if (init == "2")
                        {
                            nstat = int.Parse(total)+1;
                        }
                        for (int i = bypass; i <= nstat; i++)
                        {
                           
                            Range range = ws.get_Range("A" + i.ToString(), "J" + i.ToString());
                            Array values = (Array)range.Cells.Value;
                            string[] strArray = ConvertToStringArray(values);
                            if (strArray[0].Length > 0)
                            {
                                listView1.Items.Add(new ListViewItem(strArray));
                            }

                        }
                        groupBox1.Text = listView1.Items.Count + " بيانات في لستة";

                    }

                    excelApp.Visible = false;
                    workbook.Save();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    // Marshal.FinalReleaseComObject(worksheet);
                    workbook.Close();
                    Marshal.FinalReleaseComObject(workbook);
                    excelApp.Quit();
                    Marshal.FinalReleaseComObject(excelApp);

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void saveexcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            metroContextMenu1.Enabled = false;
            WriteExcel();            
            using (SaveFileDialog Save = new SaveFileDialog())
            {
                Save.Filter = "Excel Files (*.xlsx)|*.xlsx";
                Save.Title = "Save";
                if (Save.ShowDialog() == DialogResult.OK)
                {
                    Thread.Sleep(1000);
                    fixexcel("Cache\\Save", Save.FileName);

                    MessageBox.Show("Done");
                    metroContextMenu1.Enabled = true;
                }
            }
            
        }

        private void sortbynameToolStripMenuItem_Click(object sender, EventArgs e)
        {
            listView1.Sorting = System.Windows.Forms.SortOrder.Ascending;

        }

        private void عشوائيToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                listView1.Sorting = System.Windows.Forms.SortOrder.None;
                Randomize(listView1);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            listView1.Items.Clear();
            groupBox1.Text = listView1.Items.Count + " بيانات في لستة";

        }

        private void frToolStripMenuItem_Click(object sender, EventArgs e)
        {
            listView1.Refresh();
            groupBox1.Text = listView1.Items.Count + " بيانات في لستة";

        }



        private void hallsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            try
            {
                foreach (var process in Process.GetProcessesByName("EXCEL"))
                    process.Kill();
                int xx = 0;
                int loop = 0;
                String dtPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Halls\\";
                DirectoryInfo dir = new DirectoryInfo(dtPath);
                foreach (FileInfo flInfo in dir.GetFiles())
                {
                    var excelApp = new Microsoft.Office.Interop.Excel.Application();
                    var workbook = excelApp.Workbooks.Open(flInfo.FullName);
                    var sheets = workbook.Worksheets;
                    // Worksheet worksheet = (Worksheet)sheets.get_Item(1);
                    // var usedRange = worksheet.UsedRange;
                    //  usedRange.RemoveDuplicates(1);
                    foreach (Worksheet ws in workbook.Worksheets)
                    {
                        for (int i = 1; i <= ws.UsedRange.Rows.Count; i++)
                        {
                            Range range = ws.get_Range("A" + i.ToString(), "J" + i.ToString());
                            Array values = (Array)range.Cells.Value;
                            string[] strArray = ConvertToStringArray(values); 
                            if (strArray[1].Length > 0)
                            {
                                if ((int)mode.Students == _switch)
                                {
                                    for (int x = 0; x < int.Parse(strArray[1]); x++) // خلايه بال عامود 2 بال اكسيل
                                    {

                                        if (xx != listView1.Items.Count)
                                        {
                                            listView1.Items[xx].SubItems[4].Text = strArray[0];
                                            xx++;
                                        }

                                    }
                                }
                                if ((int)mode.Teachers == _switch)
                                {
                                    for (int h = 0; h < numericUpDown2.Value; h++)
                                    {
                                        if (loop != listView1.Items.Count)
                                        {
                                            if (listView1.Items[loop].SubItems[0].Text != "")
                                                listView1.Items[loop].SubItems[4].Text = strArray[0];

                                            loop++;
                                        }
                                    }

                                }


                            }

                        }
                    }
                    excelApp.Visible = false;
                    workbook.Save();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    //Marshal.FinalReleaseComObject(worksheet);
                    workbook.Close();
                    Marshal.FinalReleaseComObject(workbook);
                    excelApp.Quit();
                    Marshal.FinalReleaseComObject(excelApp);

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void addhallToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string hall = "";
            hall = Interaction.InputBox("اكتب اسم", "أضافة عنصر", "");

            foreach (ListViewItem item in listView1.SelectedItems)
            {
                item.SubItems[4].Text = hall.ToString();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem items in listView1.Items)
            {
                items.SubItems[4].Text = metroTextBox1.Text;
            }
        }

        private void cdToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < listView1.Items.Count; i++)
                {
                    for (int j = 1; j < listView1.Items.Count; j++)
                    {
                        if (i != j)
                        {
                            if (listView1.Items[i].Text == listView1.Items[j].Text)
                            {
                                sb.AppendLine(listView1.Items[j].Text);
                                listView1.Items[j].Remove();
                            }
                        }
                    }
                }
                groupBox1.Text = listView1.Items.Count + " بيانات في لستة";

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void c1ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in listView1.Items)
            {
                item.SubItems[0].Text = null;
            }
        }

        private void c2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in listView1.Items)
            {
                item.SubItems[1].Text = null;
            }
        }

        private void c3ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in listView1.Items)
            {
                item.SubItems[2].Text = null;
            }
        }

        private void c4ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in listView1.Items)
            {
                item.SubItems[3].Text = null;
            }
        }

        private void c5ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in listView1.Items)
            {
                item.SubItems[4].Text = null;
            }
        }

        private void indexToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in listView1.SelectedItems)
            {
                MessageBox.Show(item.Index + " رقم الصف");
            }
        }

        private void dc1ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string firstrow = "";
            firstrow = Interaction.InputBox("رقم بداية الصف", "اضافة بيانات ", "");
            string lastrow = "";
            lastrow = Interaction.InputBox("رقم نهاية الصف", "اضافة بيانات ", "");
            string name = "";
            name = Interaction.InputBox("ادخال نص ", "اضافة بيانات ", "");
            if (firstrow != "" && lastrow != "" && name != "")
            {
                foreach (ListViewItem item in listView1.SelectedItems)
                {
                    for (int i = 0; i < listView1.Items.Count; i++)
                    {
                        for (int j = int.Parse(firstrow); j <= int.Parse(lastrow); j++)
                        {
                            listView1.Items[j].SubItems[1].Text = name;
                        }
                    }
                }
            }
        }

        private void dc2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string firstrow = "";
            firstrow = Interaction.InputBox("رقم بداية الصف", "اضافة بيانات ", "");
            string lastrow = "";
            lastrow = Interaction.InputBox("رقم نهاية الصف", "اضافة بيانات ", "");
            string name = "";
            name = Interaction.InputBox("ادخال نص", "اضافة بيانات ", "");
            if (firstrow != "" && lastrow != "" && name != "")
            {
                foreach (ListViewItem item in listView1.SelectedItems)
                {
                    for (int i = 0; i < listView1.Items.Count; i++)
                    {
                        for (int j = int.Parse(firstrow); j <= int.Parse(lastrow); j++)
                        {
                            listView1.Items[j].SubItems[2].Text = name;
                        }
                    }
                }
            }
        }

        private void dc3ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string firstrow = "";
            firstrow = Interaction.InputBox("رقم بداية الصف", "اضافة بيانات ", "");
            string lastrow = "";
            lastrow = Interaction.InputBox("رقم نهاية الصف", "اضافة بيانات ", "");
            string name = "";
            name = Interaction.InputBox("ادخال نص", "اضافة بيانات ", "");
            if (firstrow != "" && lastrow != "" && name != "")
            {
                foreach (ListViewItem item in listView1.SelectedItems)
                {
                    for (int i = 0; i < listView1.Items.Count; i++)
                    {
                        for (int j = int.Parse(firstrow); j <= int.Parse(lastrow); j++)
                        {
                            listView1.Items[j].SubItems[3].Text = name;
                        }
                    }
                }
            }
        }

        private void dc4ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string firstrow = "";
            firstrow = Interaction.InputBox("رقم بداية الصف", "اضافة بيانات ", "");
            string lastrow = "";
            lastrow = Interaction.InputBox("رقم نهاية الصف", "اضافة بيانات ", "");
            string name = "";
            name = Interaction.InputBox("ادخال نص", "اضافة بيانات ", "");
            if (firstrow != "" && lastrow != "" && name != "")
            {
                foreach (ListViewItem item in listView1.SelectedItems)
                {
                    for (int i = 0; i < listView1.Items.Count; i++)
                    {
                        for (int j = int.Parse(firstrow); j <= int.Parse(lastrow); j++)
                        {
                            listView1.Items[j].SubItems[4].Text = name;
                        }
                    }
                }
            }
        }

        private void metroRadioButton1_CheckedChanged(object sender, EventArgs e)
        {
            listView1.Items.Clear();
            _switch = 1;
            Load_teachers();

        }

        private void metroRadioButton2_CheckedChanged(object sender, EventArgs e)
        {
            listView1.Items.Clear();
            _switch = 0;
            Load_Students();

        }
    }
}
