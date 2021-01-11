using ExOne.Models;
using LiteDB;
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

        private List<ProjList> GetAll(int a) 
        {
            var list = new List<ProjList>();
            
            var col = DbContext.GetInstance().GetCollection<ProjList>();
            //select * from K_tbl where CreatedAt in (select max(CreatedAt) from K_tbl group by ReqFormNo, ReqFormDesc, Stage)

            List<DateTime> awx = new List<DateTime>();

            if (a == 1) //L1
            {
                awx = (from x in col.FindAll().OrderByDescending(d => d.CreatedAt)
                        group x by new { x.ReqFormNo, x.ReqFormDesc}
                        into xx
                        select xx.Max(m => m.CreatedAt)
                        ).ToList();
            }
            else if (a == 2) //L2
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
                DateTime dt;
                _id.ApplyDate = DateTime.TryParse(_id.ApplyDate, out dt ) ? Convert.ToDateTime(_id.ApplyDate).ToString("MM/dd/yyyy") : _id.ApplyDate;
                _id.TestDateEstimate = DateTime.TryParse(_id.TestDateEstimate, out dt) ? Convert.ToDateTime(_id.TestDateEstimate).ToString("MM/dd/yyyy") : _id.TestDateEstimate;
                _id.StageActualFinish = DateTime.TryParse(_id.StageActualFinish, out dt) ? Convert.ToDateTime(_id.StageActualFinish).ToString("MM/dd/yyyy") : _id.StageActualFinish;
                _id.StageEstimateFinish = DateTime.TryParse(_id.StageEstimateFinish, out dt) ? Convert.ToDateTime(_id.StageEstimateFinish).ToString("MM/dd/yyyy") : _id.StageEstimateFinish;
                _id.UserExpectedDate = DateTime.TryParse(_id.UserExpectedDate, out dt) ? Convert.ToDateTime(_id.UserExpectedDate).ToString("MM/dd/yyyy") : _id.UserExpectedDate;
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
            subDgv.Visible = false;
            dataGridView1.DataSource = GetHeaderData();
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].Width = 75;
            dataGridView1.Columns[2].Width = 75;
            dataGridView1.Columns[2].HeaderText = "IT PIC";
            dataGridView1.Columns[3].Width = 120;
            dataGridView1.Columns[3].HeaderText = "Request Form No.";
            dataGridView1.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            level = 1;
        }

        private void Get2List() //L2
        {
            subDgv.Visible = true;
            var col = DbContext.GetInstance().GetCollection<ProjList>();
            var ab = col.FindById(_id);
            var ac = (from b in GetAll(2)
                     where b.ReqFormDesc == ab.ReqFormDesc && b.ReqFormNo == ab.ReqFormNo
                     select b).ToList();
            var ad = ac.FirstOrDefault();
            
            hierarchy.Text = ad.ReqFormNo + " -> " + (ad.ReqFormDesc.Length > 20 ? ad.ReqFormDesc.Substring(0, 20) : ad.ReqFormDesc);
            subDgv.DataSource = null;
            subDgv.Refresh();
            subDgv.DataSource = ac;
            ColSetting();
            level = 2;
        }

        private void Get3List() //L3
        {
            var col = DbContext.GetInstance().GetCollection<ProjList>();
            var selected = col.FindById(_id);
            var getData3 = col.Find(z => z.Stage == selected.Stage)
                .Where(x => x.ReqFormNo == selected.ReqFormNo)
                .Where(x => x.ReqFormDesc == selected.ReqFormDesc)
                .OrderByDescending(y => y.CreatedAt).ToList();
            hierarchy.Text = hierarchy.Text + " -> " + selected.Stage;
            subDgv.DataSource = null;
            subDgv.Refresh();
            subDgv.DataSource = getData3;
            ColSetting();
            level = 3;
        }

        private void ColSetting()
        {
            for (var i = 0; i <= 3; i++)
            {
                subDgv.Columns[i].Visible = false;
            }
            subDgv.Columns[15].Visible = false;
            subDgv.Columns[17].Visible = false;
            subDgv.Columns[4].Width = 75;
            subDgv.Columns[5].HeaderText = "IT PIC";
            subDgv.Columns[5].Width = 75;
            subDgv.Columns[6].HeaderText = "Request Form No.";
            subDgv.Columns[6].Width = 120;
            subDgv.Columns[7].HeaderText = "Request Form Description";
            subDgv.Columns[7].Width = 300;
            subDgv.Columns[8].Width = 85;
            subDgv.Columns[9].HeaderText = "Expected Finish Date";
            subDgv.Columns[9].Width = 90;
            subDgv.Columns[10].HeaderText = "Estimate Stage Finish";
            subDgv.Columns[10].Width = 90;
            subDgv.Columns[11].HeaderText = "Actual Finish Date";
            subDgv.Columns[11].Width = 90;
            subDgv.Columns[12].HeaderText = "IT Est. give Test Date";
            subDgv.Columns[12].Width = 90;
            subDgv.Columns[13].HeaderText = "Apply Date";
            subDgv.Columns[13].Width = 90;
            subDgv.Columns[14].Width = 300;
            subDgv.Columns[16].HeaderText = "Upload Time";
            subDgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            subDgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Get1List();
            string aj = "Cong";// "2020/11/10";
            string am;
            DateTime ak;
            if (DateTime.TryParse(aj, out ak))
            {
                am = aj;
                // or do some count process (convert dt and count)
            }
            else
            {
                am = aj;
            }

        }

        private void BackBtn_Click(object sender, EventArgs e)
        {
            switch(level)
            {
                case 1:
                    Get1List();
                    break;
                case 2:
                    subDgv.Visible = false;
                    hierarchy.Text = null;
                    level--;
                    break;
                case 3:
                    Get2List();
                    break;
                default:
                    break;
            }
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex == -1)
                return;
            
            _id = Guid.Parse(dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString());
            Get2List();
            
        }

        private void subDgv_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex == -1)
                return;
            if (level == 2)
            {
                _id = Guid.Parse(subDgv.Rows[e.RowIndex].Cells[0].Value.ToString());
                Get3List();
                //level++;
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
