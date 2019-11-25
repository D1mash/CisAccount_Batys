﻿using System;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace Учет_цистерн
{
    public partial class ReportForm : Form
    {
        BindingSource source = new BindingSource();
        DataTable getserv;
        string UserFIO;

        public ReportForm(string userFIO)
        {
            InitializeComponent();
            this.UserFIO = userFIO;
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            if (comboBox2.SelectedIndex == 0)
            {
                string RefreshAll = "exec dbo.GetReportAllRenderedService_v1 '" + dateTimePicker1.Value.Date.ToString() + "','" + dateTimePicker2.Value.Date.ToString() + "'";
                DataTable dt;
                dt = DbConnection.DBConnect(RefreshAll);
                source.DataSource = dt;
                dataGridView1.DataSource = source;
                dataGridView1.Columns[0].Visible = false;
                dataGridView1.Columns[14].Visible = true;
                progressBar.Maximum = TotalRow(dt);
                toolStripLabel1.Text = TotalRow(dt).ToString();

                string GetCountServiceCost = "exec dbo.GetCountServiceCost '" + dateTimePicker1.Value.Date.ToString() + "','" + dateTimePicker2.Value.Date.ToString() + "'";
                getserv = DbConnection.DBConnect(GetCountServiceCost);
            }
            else
            {
                string Refresh = "dbo.GetReportRenderedServices_v1 '" + dateTimePicker1.Value.Date.ToString() + "','" + dateTimePicker2.Value.Date.ToString() + "','" + comboBox2.SelectedValue + "'";
                DataTable dataTable;
                dataTable = DbConnection.DBConnect(Refresh);
                source.DataSource = dataTable;
                dataGridView1.DataSource = source;
                dataGridView1.Columns[0].Visible = false;
                progressBar.Maximum = TotalRow(dataTable);
                toolStripLabel1.Text = TotalRow(dataTable).ToString();

                string GetCountServiceCost = "exec dbo.GetCountServiceCost_byOwner  '" + dateTimePicker1.Value.Date.ToString() + "','" + dateTimePicker2.Value.Date.ToString() + "','" + comboBox2.SelectedValue + "'";
                getserv = DbConnection.DBConnect(GetCountServiceCost);
            }
        }

        private int TotalRow(DataTable dataTable)
        {
            int i = 0;
            foreach (DataRow dr in dataTable.Rows)
            {
                i++;
            }
            return i;
        }

        private void ReportForm_Load(object sender, EventArgs e)
        {
            String Owner = "Select * from d__Owner";
            DataTable OwnerDT = DbConnection.DBConnect(Owner);
            var dr = OwnerDT.NewRow();
            dr["Id"] = -1;
            dr["Name"] = "Все";
            OwnerDT.Rows.InsertAt(dr, 0);
            comboBox2.DataSource = OwnerDT;
            comboBox2.DisplayMember = "Name";
            comboBox2.ValueMember = "ID";

            DateTime now = DateTime.Now;
            var startDate = new DateTime(now.Year, now.Month, 1);
            var endDate = startDate.AddMonths(1).AddDays(-1);

            dateTimePicker1.Value = startDate;
            dateTimePicker2.Value = endDate;
        }

        private void Btn_Excel_Click(object sender, EventArgs e)
        {

            if (dataGridView1.Rows != null && dataGridView1.Rows.Count != 0)
            {
                if (backgroundWorker.IsBusy)
                    return;
                using (SaveFileDialog saveFileDialog = new SaveFileDialog() { Filter = "Excel file (*.xlsx)|*.xlsx|All files(*.*)|*.*" })
                {
                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        _inputParametr1.FileName = saveFileDialog.FileName;
                        _inputParametr1.owner = comboBox2.Text;
                        progressBar.Minimum = 0;
                        progressBar.Value = 0;
                        backgroundWorker.RunWorkerAsync(_inputParametr1);
                    }
                }
            }
            else
            {
                MessageBox.Show("Обновите данные!", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        struct DataParametr
        {
            public string FileName { get; set; }
            public string owner { get; set; }
        }

        DataParametr _inputParametr1;

        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

        static Process GetExcelProcess(Excel.Application excelApp)
        {
            GetWindowThreadProcessId(excelApp.Hwnd, out int id);
            return Process.GetProcessById(id);
        }

        private void BackgroundWorker_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            try
            {
                string path = AppDomain.CurrentDomain.BaseDirectory + @"ReportTemplates\Реестр  за арендованных и  собственных вагон-цистерн компании.xlsx";
                //var folder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                string fileName = ((DataParametr)e.Argument).FileName;
                string ownerName = ((DataParametr)e.Argument).owner;

                Excel.Application app = new Excel.Application();
                Process appProcess = GetExcelProcess(app);
                Excel.Workbook workbook = app.Workbooks.Open(path);
                Excel.Worksheet worksheet = workbook.Worksheets.get_Item("ТОО Казыкурт");
                app.Visible = false;

                int cellRowIndex = 0;
                int totalTOR4 = 0;

                if (ownerName == "Все")
                {
                    worksheet.Range["C4"].Value = "всех";
                }
                else
                {
                    worksheet.Range["C4"].Value = ownerName;
                }

                worksheet.Range["C6"].Value = "в ТОО Казыгурт-Юг c " + dateTimePicker1.Value.ToShortDateString() + " по " + dateTimePicker2.Value.ToShortDateString();

                FormattingExcelCells(worksheet.Range["C6"], false, false);

                worksheet.Range["K21"].Value = UserFIO;

                worksheet.Range["B15:K23"].Cut(worksheet.Cells[dataGridView1.Rows.Count + 17 + getserv.Rows.Count*2, 2]);
                
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    worksheet.Cells[i + 10, 1] = i + 1;
                    for (int j = 1; j < dataGridView1.Columns.Count-2; j++)
                    {
                        // Excel index starts from 1,1. As first Row would have the Column headers, adding a condition check.
                        //if (cellRowIndex == 1)
                        //{
                        //    worksheet.Cells[cellRowIndex, cellColumnIndex] = dataGridView1.Columns[j].HeaderText;
                        //}
                        if(j!=3 && j<4)
                        {
                            worksheet.Cells[i + 10, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                        }
                        else
                        {
                            if(j == 3)
                            {
                                if (dataGridView1.Rows[i].Cells[j].Value.ToString().Trim() == "8")
                                {
                                    worksheet.Cells[i + 10, 5] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                                }
                                else
                                {
                                    worksheet.Cells[i + 10, 4] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                                }
                            }
                        }
                        if(j>=4 && j<=5)
                        {
                            worksheet.Cells[i + 10, j + 3] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                        }
                        else
                        {
                            if(j>=6 && j <= 9)
                            {
                                if(dataGridView1.Rows[i].Cells[j].Value.ToString().Trim() == "True")
                                {
                                    worksheet.Cells[i + 10, j + 3] = "✓";
                                }
                                else
                                {
                                    worksheet.Cells[i + 10, j + 3] = " ";
                                }
                            }
                        }
                        if(j>9)
                        {
                            worksheet.Cells[i + 10, j + 3] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                        }
                    }

                    //Excel.Range priceRange = worksheet.Range[worksheet.Cells[i + 10, 15], worksheet.Cells[dataGridView1.Rows.Count + 9, 15]];
                    //priceRange.NumberFormat = "0.00";
                
                    Excel.Range range = worksheet.Range[worksheet.Cells[i + 10, 1], worksheet.Cells[i + 10, dataGridView1.Columns.Count]];
                    FormattingExcelCells(range, true, true);

                    backgroundWorker.ReportProgress(i);

                    cellRowIndex++;
                }

                worksheet.Cells[dataGridView1.Rows.Count + 12, 2] = "=C6";

                if (ownerName == "Все")
                {
                    worksheet.Cells[dataGridView1.Rows.Count + 14, 2] = "Всего обработано вагонов - цистерн всех собственников по видам операций:";
                }
                else
                {
                    worksheet.Cells[dataGridView1.Rows.Count + 14, 2] = "Всего обработано вагонов - цистерн " + ownerName + " по видам операций:";
                }

                int rowcount = 0;
                for (int i = 0; i < getserv.Rows.Count; i++)
                {
                    rowcount++;
                        for (int j = 0; j < getserv.Columns.Count; j++)
                        {
                            if (j == 0)
                            {
                                worksheet.Cells[i + cellRowIndex + 15+rowcount, j + 2] = getserv.Rows[i][j].ToString();
                            }
                            else
                            {
                                worksheet.Cells[i + cellRowIndex + 15+rowcount, j + 12] = getserv.Rows[i][j].ToString();
                            }
                        }
                }

                worksheet.Cells[dataGridView1.Rows.Count + 14, 13] = cellRowIndex;

                worksheet.Cells[dataGridView1.Rows.Count + getserv.Rows.Count*2 + 17, 13] = totalTOR4;

                worksheet.Cells[dataGridView1.Rows.Count + getserv.Rows.Count*2 + 19, 14] = TotalSum();

                Excel.Range range1 = worksheet.Range[worksheet.Cells[dataGridView1.Rows.Count + 12, 2], worksheet.Cells[dataGridView1.Rows.Count + getserv.Rows.Count * 2 + 19, 14]];
                FormattingExcelCells(range1, false, false);

                workbook.SaveAs(fileName);
                app.Quit();
                appProcess.Kill();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message,"",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
        }

        //ИТОГО СУММА
        private Double TotalSum()
        {
            Double sum = 0;
            
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                    sum += Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value);
            }
            return sum;
        }

        public void FormattingExcelCells(Excel.Range range, bool val1, bool val2)
        {
            //range.EntireColumn.AutoFit();
            range.Font.Name = "Arial Cyr";
            range.Font.Size = 9;
            range.Font.FontStyle = "Bold";
            if (val1 == true)
            {
                Excel.Borders border = range.Borders;
                border.LineStyle = Excel.XlLineStyle.xlContinuous;
                border.Weight = 2d;
            }
            if (val2 == true)
            {
                range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            }
            else
            {
                range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            }
        }

        private void BackgroundWorker_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            progressBar.Value = e.ProgressPercentage;
            LblStatus.Text = "Обработка строки.. " + e.ProgressPercentage.ToString() /*+ " из " + TotalRow()*/;
        }

        private void BackgroundWorker_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            if(e.Error == null)
            {
                Thread.Sleep(1);
                LblStatus.Text = "Данные были успешно экспортированы";
                progressBar.Value = 0;
            }
        }
    }
}
