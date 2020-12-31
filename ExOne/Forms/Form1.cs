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

            //DbContext.GetInstance().GetCollection<ProjList>().Insert(pl);
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
                var fileName = Path.GetFileName(filePath);
                var excelFile = new FileInfo(filePath);
                using (var p = new ExcelPackage(excelFile))
                {
                    var ws = p.Workbook.Worksheets.FirstOrDefault();
                    var endrow = ws.Dimension.End.Row;
                    var endcol = 14;

                    for (int row = 2; row <= endrow; row++)
                    {
                        List<string> kind = new List<string>();
                        for (int col = 1; col <= endcol; col++)
                        {
                            if (ws.Cells[row, col].Value != null)
                            {
                                kind.Add(ws.Cells[row, col].Value.ToString());
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
                            FileName = fileName,
                            CreatedAt = DateTime.Now,
                            UpdatedAt = DateTime.Now
                        };
                        DbContext.GetInstance().GetCollection<ProjList>().Insert(pl);
                        kind.Clear();
                        if (ws.Cells[row + 1, 7].Value == null)
                        {
                            break;
                        }
                    }
                }
            }
        }

        private List<ProjList> GetAll()
        {
            var list = new List<ProjList>();
            
            var col = DbContext.GetInstance().GetCollection<ProjList>();
            //select * from K_tbl where k18 in (select max(k18) from K_tbl group by k06, k08)
            var awx = (from x in col.FindAll().OrderByDescending(d => d.CreatedAt)
                       group x by new { x.ReqFormNo, x.Stage }
                      into xx
                      select new { 
                          //ReqFormNo = xx.Key.ReqFormNo, 
                          //dt = (from x2 in xx select x2.CreatedAt).Max() 
                          ReqNo = xx.Key,
                          DT = xx.Max(m => m.CreatedAt)
                      })
                      .ToArray();
            
            var detx = from x in col.FindAll()
                       where awx.Any(m => m.DT == x.CreatedAt)
                       select x;
            foreach (ProjList _id in col.FindAll())
            {
                list.Add(_id);
            }
            
            return list;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = GetAll();
        }
    }
}
