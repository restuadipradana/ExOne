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
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExOne
{
    public partial class Form1 : Form
    {
        private Guid _id;
        private int level;

        public Form1()
        {
            InitializeComponent();
        }
        public int ax = 0;

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = GetAll(1);

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
            
            if (filePath == null || filePath == "")
            {
                MessageBox.Show("Please select file first!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                try
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
                            Thread.Sleep(1); //avoid same time insert 
                            if (ws.Cells[row + 1, 7].Value == null)
                            {
                                break;
                            }
                        }
                    }
                    MessageBox.Show("File " + fileName + " has been uploaded", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Detail: " + ex, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                
            }
        }

        private List<ProjList> GetAll(int a) // L1
        {
            var list = new List<ProjList>();
            
            var col = DbContext.GetInstance().GetCollection<ProjList>();
            //select * from K_tbl where CreatedAt in (select max(CreatedAt) from K_tbl group by ReqFormNo, ReqFormDesc, Stage)

            List<DateTime> awx = new List<DateTime>();

            if (a == 1)
            {
                awx = (from x in col.FindAll().OrderByDescending(d => d.CreatedAt)
                        group x by new { x.ReqFormNo, x.ReqFormDesc}
                        into xx
                        select xx.Max(m => m.CreatedAt)
                        ).ToList();
            }
            else if (a == 2)
            {
                awx = (from x in col.FindAll().OrderByDescending(d => d.CreatedAt)
                        group x by new { x.ReqFormNo, x.ReqFormDesc, x.Stage }
                        into xx
                        select xx.Max(m => m.CreatedAt)
                        ).ToList();
            }

            var detx = from x in col.FindAll().OrderByDescending(d => d.ReqFormNo)
                       where awx.Contains(x.CreatedAt)
                       select x;


            foreach (ProjList _id in detx)
            {
                list.Add(_id);
            }
            
            return list;
        }

        private List<L1list> GetHeaderData()
        {
            var list = new List<L1list>();
            var all = GetAll(1);

            foreach ( var idx in all)
            {
                L1list l1 = new L1list()
                {
                    _id = idx.id,
                    Applicant = idx.Applicant,
                    ITPIC = idx.PIC,
                    RequestFrom = idx.ReqFormNo,
                    Description = idx.ReqFormDesc
                };
                list.Add(l1);
            }

            return list;
        }

        private void Get1List()
        {
            dataGridView1.DataSource = GetHeaderData();
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].Width = 75;
            dataGridView1.Columns[2].Width = 75;
            dataGridView1.Columns[3].Width = 120;
            dataGridView1.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            level = 1;
        }

        private void Get2List()
        {
            var list = new List<L1list>();
            var all = GetAll(2);
            var ab = (from a in all
                     where a.id == _id
                     select a).FirstOrDefault();
            var ac = (from b in all
                     where b.ReqFormDesc == ab.ReqFormDesc && b.ReqFormNo == ab.ReqFormNo
                     select b).ToList();
            dataGridView1.DataSource = null;
            dataGridView1.Refresh();
            dataGridView1.DataSource = ac;
            dataGridView1.Columns[0].Visible = false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Get1List();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (level == 1)
            {
                label1.Text = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                _id = Guid.Parse(dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString());
                Get2List();
                level++;
            }
        }

        private void BackBtn_Click(object sender, EventArgs e)
        {
            if (level == 2)
            {
                Get1List();
                //level--;
            }
            else if (level == 1)
            {
                Form1_Load(this, null);
            }
        }
    }

    public class L1list
    {
        public Guid _id { get; set; }
        public string Applicant { get; set; }
        public string ITPIC { get; set; }
        public string RequestFrom { get; set; }
        public string Description { get; set; }
    }
}
