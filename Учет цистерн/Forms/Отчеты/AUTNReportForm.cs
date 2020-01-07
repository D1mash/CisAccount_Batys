﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Учет_цистерн.Forms.Отчеты
{
    public partial class AUTNReportForm : Form
    {
        //BindingSource source = new BindingSource();
        DataTable dt;

        public AUTNReportForm()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string RefreshAll = "exec [dbo].[GetReportAUTN] '" + dateTimePicker1.Value.Date.ToString() + "','" + dateTimePicker2.Value.Date.ToString() + "'";
            DataTable dt;
            dt = DbConnection.DBConnect(RefreshAll);
            gridControl1.DataSource = dt;
            progressBar.Maximum = TotalRow(dt);
            toolStripLabel1.Text = TotalRow(dt).ToString();
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

        private void btn_Excel_Click(object sender, EventArgs e)
        {
            if (dt.Rows != null && dt.Rows.Count != 0)
            {
                if (backgroundWorker.IsBusy)
                    return;
                else
                {
                    progressBar.Minimum = 0;
                    progressBar.Value = 0;
                    backgroundWorker.RunWorkerAsync();
                }
            }
            else
            {
                MessageBox.Show("Обновите данные!", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

        static Process GetExcelProcess(Excel.Application excelApp)
        {
            GetWindowThreadProcessId(excelApp.Hwnd, out int id);
            return Process.GetProcessById(id);
        }

        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                string path = AppDomain.CurrentDomain.BaseDirectory + @"ReportTemplates\Реестр АУТН.xlsx";
                Excel.Application app = new Excel.Application();
                Process appProcess = GetExcelProcess(app);
                Excel.Workbook workbook = app.Workbooks.Open(path);
                Excel.Worksheet worksheet = workbook.Worksheets.get_Item("АУТН");
                app.Visible = false;
                object misValue = System.Reflection.Missing.Value;

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 3, 1] = i;
                        if (j < 8)
                        {
                            worksheet.Cells[i + 3, j + 2] = dt.Rows[i][j].ToString();
                        }
                        else
                        {
                            if (j >= 8 && j < 18)
                            {
                                if (dt.Rows[i][j].ToString() == "1")
                                {
                                    worksheet.Cells[i + 3, j + 2] = "Да";
                                }
                                else
                                {
                                    worksheet.Cells[i + 3, j + 2] = "Нет";
                                }
                            }
                            else
                            {
                                worksheet.Cells[i + 3, j + 2] = dt.Rows[i][j].ToString();
                            }
                        }

                    }

                    Excel.Range range = worksheet.Range[worksheet.Cells[i + 3, 1], worksheet.Cells[i + 3, dt.Columns.Count+1]];
                    FormattingExcelCells(range, true, true);

                    backgroundWorker.ReportProgress(i);
                }

                app.DisplayAlerts = false;
                workbook.SaveAs(/*fileName*/@"D:\Отчеты\Реестр АУТН.xlsx", Excel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                workbook.Close(0);
                app.Quit();
                appProcess.Kill();

                Process.Start(@"D:\Отчеты\Реестр АУТН.xls");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.Value = e.ProgressPercentage;
            LblStatus.Text = "Обработка строки.. " + e.ProgressPercentage.ToString() /*+ " из " + TotalRow()*/;
        }

        private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error == null)
            {
                Thread.Sleep(1);
                LblStatus.Text = "Данные были успешно экспортированы";
                progressBar.Value = 0;
            }
        }

        public void FormattingExcelCells(Excel.Range range, bool val1, bool val2)
        {
            range.EntireColumn.AutoFit();
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

        private void AUTNReportForm_Load(object sender, EventArgs e)
        {
            DateTime now = DateTime.Now;
            var startDate = new DateTime(now.Year, now.Month, 1);
            var endDate = startDate.AddMonths(1).AddDays(-1);

            dateTimePicker1.Value = startDate;
            dateTimePicker2.Value = endDate;
        }
    }
}