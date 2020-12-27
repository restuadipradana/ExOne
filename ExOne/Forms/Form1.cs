using ExOne.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExOne
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public int ax = 0;

        private void button1_Click(object sender, EventArgs e)
        {
            
            var st = "AB" + ax.ToString();
            ax++;

            ProjList pl = new ProjList()
            {
                System = "My System",
                ErrKind = st,
                Desc = "Percobaan Lite DB",
                Applicant = "Restu Adi ",
                PIC = "Pradana",
                ReqFormNo = "TSH-R2092100",
                ApplyDate = "20/12/2020",
                Stage = "initial",
                CreatedAt = DateTime.Now,
                UpdatedAt = DateTime.Now
            };

            DbContext.GetInstance().GetCollection<ProjList>().Insert(pl);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var ads = DbContext.GetInstance().GetCollection<ProjList>().FindAll();
            foreach (var itm in ads)
            {
                button2.Text = itm.ErrKind;
            }
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string filePath = string.Empty;
            string fileExt = string.Empty;
            OpenFileDialog file = new OpenFileDialog(); //open dialog to choose file 
            file.Filter = "Excel Files|*.xlsx;*.xls";
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) //if there is a file choosen by the user  
            {
                filePath = file.FileName; //get the path of the file  
                fileExt = Path.GetExtension(filePath); //get the file extension  
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        textBox1.Text = filePath;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
                else
                {
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            var filePath = textBox1.Text;
            if (filePath == null)
            {
                MessageBox.Show("Please select file first!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                var excelFile = new FileInfo(filePath);
                using (var p = new ExcelPackage(excelFile))
                {
                    var ws = p.Workbook.Worksheets.FirstOrDefault();
                    var endrow = ws.Dimension.End.Row;
                    var endcol = 14;

                    for (int row = 2; row <= endrow; row++)//looking for article row
                    {
                        List<string> kind = new List<string>(); //Article
                        for (int col = 1; col <= endcol; col++)
                        {
                            if (ws.Cells[row, col].Value != null)
                            {
                                kind.Add(ws.Cells[row, col].Value.ToString()); //Model Name
                            }
                            else
                            {
                                kind.Add("");
                            }
                        }
                        ProjList pl = new ProjList()
                        {
                            System = kind[0],
                            ErrKind = kind[1],
                            Desc = kind[2],
                            Applicant = kind[3],
                            PIC = kind[4],
                            ReqFormNo = kind[5],
                            ReqFormDesc = kind[6],
                            Stage = kind[7],
                            UserExpectedDate = kind[8],
                            StageEstimateFinish = kind[9],
                            StageActualFinish = kind[10],
                            TestDateEstimate = kind[11],
                            ApplyDate = kind[12],
                            Memo = kind[13],
                            CreatedAt = DateTime.Now,
                            UpdatedAt = DateTime.Now
                        };
                        DbContext.GetInstance().GetCollection<ProjList>().Insert(pl);
                        kind.Clear();
                    }
                }
            }
        }
    }
}
