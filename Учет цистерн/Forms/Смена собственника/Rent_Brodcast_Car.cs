using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Grid;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Учет_цистерн.Forms.Смена_собственника
{
    public partial class Rent_Brodcast_Car : Form
    {
        int SelectItemRow2;
        int OwnerIDS;
        int Grid2Id, Grid1Id, Grid3Id;
        int OwnerId_v2;
        string OwnerName;
        string date;
        //string product;
        string role, User_ID;
        private TradeWright.UI.Forms.TabControlExtra tabControl1;

        public Rent_Brodcast_Car(TradeWright.UI.Forms.TabControlExtra tabControl1, string role, string UserID)
        {
            InitializeComponent();
            this.role = role;
            this.User_ID = UserID;
            this.tabControl1 = tabControl1;
        }

        private void Rent_Brodcast_Car_Load(object sender, EventArgs e)
        {
            if (role == "1" | role == "1002")
            {
                simpleButton1.Enabled = true;
            }
            else
            {
                if (role == "2")
                {

                    simpleButton1.Enabled = true;
                }
                else
                {
                    simpleButton1.Enabled = true;
                }
            }

            checkEdit1_CheckedChanged(null, null);
            checkEdit3_CheckedChanged(null, null);
            checkEdit4_CheckedChanged(null, null);
            checkEdit5_CheckedChanged(null, null);
            checkEdit6_CheckedChanged(null, null);
            checkEdit7_CheckedChanged(null, null);
            checkEdit8_CheckedChanged(null, null);
        }


        private void FillCombobox()
        {
            try
            {
                string getOwner = "Select ID, Name from d__Owner";
                DataTable dt = DbConnection.DBConnect(getOwner);

                comboBox1.DataSource = dt;
                comboBox1.ValueMember = "ID";
                comboBox1.DisplayMember = "Name";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void checkEdit1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkEdit1.Checked)
            {
                textEdit1.Enabled = (checkEdit1.CheckState == CheckState.Checked);
                textEdit1.Focus();
            }
            else
            {
                textEdit1.Enabled = false;
                textEdit1.Text = String.Empty;
            }
        }

        private void checkEdit3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkEdit3.Checked)
            {
                dateEdit1.Enabled = (checkEdit3.CheckState == CheckState.Checked);

                DateTime now = DateTime.Now;
                var startDate = new DateTime(now.Year, now.Month, 1);
                dateEdit1.EditValue = startDate;

                checkEdit5.Checked = false;
            }
            else
            {
                dateEdit1.Enabled = false;
                dateEdit1.EditValue = DBNull.Value;
            }
        }

        private void checkEdit4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkEdit4.Checked)
            {
                dateEdit2.Enabled = (checkEdit4.CheckState == CheckState.Checked);

                DateTime now = DateTime.Now;
                var startDate = new DateTime(now.Year, now.Month, 1);
                var endDate = startDate.AddMonths(1).AddDays(-1);
                dateEdit2.EditValue = endDate;

                checkEdit5.Checked = false;
            }
            else
            {
                dateEdit2.Enabled = false;
                dateEdit2.EditValue = DBNull.Value;
            }
        }

        private void checkEdit5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkEdit5.Checked)
            {
                dateEdit3.Enabled = (checkEdit5.CheckState == CheckState.Checked);

                dateEdit3.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.DateTimeAdvancingCaret;
                dateEdit3.Properties.Mask.EditMask = "d"; //'Short date' format 
                dateEdit3.Properties.Mask.UseMaskAsDisplayFormat = true;
                dateEdit3.EditValue = DateTime.Today;

                checkEdit3.Checked = false;
                checkEdit4.Checked = false;
            }
            else
            {
                dateEdit3.Enabled = false;
                dateEdit3.EditValue = DBNull.Value;
            }
        }

        private void checkEdit6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkEdit6.Checked)
            {
                textEdit2.Enabled = (checkEdit6.CheckState == CheckState.Checked);
                textEdit2.Focus();
            }
            else
            {
                textEdit2.Enabled = false;
                textEdit2.Text = String.Empty;
            }
        }

        private void checkEdit7_CheckedChanged(object sender, EventArgs e)
        {
            if (checkEdit7.Checked)
            {
                comboBox1.Enabled = (checkEdit7.CheckState == CheckState.Checked);

                FillCombobox();
            }
            else
            {
                comboBox1.Enabled = false;
                comboBox1.DataSource = null;
            }
        }

        private void checkEdit8_CheckedChanged(object sender, EventArgs e)
        {
            if (checkEdit8.Checked)
            {
                textEdit3.Enabled = (checkEdit8.CheckState == CheckState.Checked);
                textEdit3.Focus();
            }
            else
            {
                textEdit3.Enabled = false;
                textEdit3.Text = String.Empty;
            }
        }

        //Кнопка Поиск
        public void simpleButton1_Click(object sender, EventArgs e)
        {
            if (checkEdit1.Checked | checkEdit3.Checked | checkEdit4.Checked | checkEdit5.Checked | checkEdit6.Checked | checkEdit7.Checked | checkEdit8.Checked)
            {
                string Date_S;
                string Date_E;
                string Date_R;

                if (dateEdit1.DateTime.ToShortDateString() == "01.01.0001")
                {
                    Date_S = "01.01.1990";
                }
                else
                {
                    Date_S = dateEdit1.DateTime.ToShortDateString();
                }

                if (dateEdit2.DateTime.ToShortDateString() == "01.01.0001")
                {
                    Date_E = "01.01.1990";
                }
                else
                {
                    Date_E = dateEdit2.DateTime.ToShortDateString();
                }

                if (dateEdit3.DateTime.ToShortDateString() == "01.01.0001")
                {
                    Date_R = "01.01.1990";
                }
                else
                {
                    Date_R = dateEdit3.DateTime.ToShortDateString();
                }

                if (checkEdit1.Checked)
                {
                    if (textEdit1.Text != String.Empty)
                    {
                        gridControl3.DataSource = null;
                        gridView3.Columns.Clear();

                        gridControl1.DataSource = null;
                        gridView1.Columns.Clear();

                        gridControl2.DataSource = null;
                        gridView2.Columns.Clear();

                        Context1_Visible(false);

                        string Search_1 = "exec dbo.Rent_Search_By_Parametrs_2 " + "@Car_Num = '" + textEdit1.Text.Trim() + "', " + "@Type = " + 2;
                        gridControl3.DataSource = DbConnection.DBConnect(Search_1);
                        gridView3.Columns[0].Visible = false;

                        if (gridView3.RowCount > 0)
                        {
                            string Search_2 = "exec dbo.Rent_Search_By_Parametrs_1 " + "@Car_Num = '" + SelectItemRow2.ToString() + "', " + "@Date_Start = '" + Date_S + "', " + " @Date_End = '" + Date_E + "', " + "@Date_Rec = '" + Date_R + "', " + "@OwnerId = '" + comboBox1.SelectedValue + "'," + "@Product = '" + textEdit3.Text + "'," + "@Rent_Num = '" + textEdit2.Text + "'," + "@Type = " + 2;
                            gridControl1.DataSource = DbConnection.DBConnect(Search_2);
                            gridView1.Columns[0].Visible = false;
                            gridView1.Columns[6].Visible = false;
                            gridView1.Columns[7].Visible = false;

                            GridColumnSummaryItem item3 = new GridColumnSummaryItem(DevExpress.Data.SummaryItemType.Count, "Номер В/Ц", "Кол.во={0}");
                            gridView3.Columns["Номер В/Ц"].Summary.Add(item3);
                        }
                        else
                        {
                            MessageBox.Show("Запись с таким номером ВЦ в заявках отсутствует !", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Введите номер ВЦ !", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                }
                else
                {
                    gridControl3.DataSource = null;
                    gridView3.Columns.Clear();

                    gridControl2.DataSource = null;
                    gridView2.Columns.Clear();

                    gridControl1.DataSource = null;
                    gridView1.Columns.Clear();

                    Context1_Visible(true);

                    string Search = "exec dbo.Rent_Search_By_Parametrs_1 " + "@Car_Num = '" + textEdit1.Text + "', " + "@Date_Start = '" + Date_S + "', " + " @Date_End = '" + Date_E + "', " + "@Date_Rec = '" + Date_R + "', " + "@OwnerId = '" + comboBox1.SelectedValue + "'," + "@Product = '" + textEdit3.Text + "'," + "@Rent_Num = '" + textEdit2.Text + "'," + "@Type = " + 1;
                    DataTable dt = DbConnection.DBConnect(Search);

                    if (dt.Rows.Count > 0)
                    {
                        gridControl2.DataSource = dt;
                        gridView2.Columns[0].Visible = false;
                        gridView2.Columns[4].Visible = false;
                        

                        gridView2_RowCellClick(null, null);
                        gridView3_RowCellClick(null, null);
                        
                        GridColumnSummaryItem item1 = new GridColumnSummaryItem(DevExpress.Data.SummaryItemType.Count, "Номер заявки", "Кол.во={0}");
                        gridView2.Columns["Номер заявки"].Summary.Add(item1);
                    }
                    else
                    {
                        MessageBox.Show("Отсутсвуют записи в базе данных","Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            else {
                MessageBox.Show("Выберите фильтр", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void gridView3_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            try
            {
                //Ловит номер вагона из грида 3
                string Id = gridView3.GetFocusedDataRow()[1].ToString();
                SelectItemRow2 = Convert.ToInt32(Id);

                gridControl1.DataSource = null;
                gridView1.Columns.Clear();

                string Search = "exec dbo.Rent_Search_By_Parametrs_3 " + "@Car_Num = '" + SelectItemRow2.ToString() + "'";
                gridControl1.DataSource = DbConnection.DBConnect(Search);
                gridView1.Columns[0].Visible = false;
                gridView1.Columns[6].Visible = false;
                gridView1.Columns[7].Visible = false;

                GridColumnSummaryItem item2 = new GridColumnSummaryItem(DevExpress.Data.SummaryItemType.Count, "Номер заявки", "Кол.во={0}");
                gridView1.Columns["Номер заявки"].Summary.Add(item2);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void gridView2_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            try
            {
                //Ловит номер заявки из грида 2
                
                date = gridView2.GetFocusedDataRow()[1].ToString();
                //product = gridView2.GetFocusedDataRow()[3].ToString();
                string OwnerID = gridView2.GetFocusedDataRow()[4].ToString();
                string RentID = gridView2.GetFocusedDataRow()[0].ToString();
                OwnerName = gridView2.GetFocusedDataRow()[3].ToString();
                Grid2Id = Convert.ToInt32(RentID);
                OwnerIDS = Convert.ToInt32(OwnerID);

                gridControl3.DataSource = null;
                gridView3.Columns.Clear();

                string Search = "exec dbo.Rent_Search_By_Parametrs_2 " + "@Rent_id  = '" + Grid2Id.ToString() + "'," + "@Type = " + 1;
                gridControl3.DataSource = DbConnection.DBConnect(Search);
                gridView3.Columns[0].Visible = false;

                GridColumnSummaryItem item3 = new GridColumnSummaryItem(DevExpress.Data.SummaryItemType.Count, "Номер В/Ц", "Кол.во={0}");
                gridView3.Columns["Номер В/Ц"].Summary.Add(item3);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void gridView1_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            try
            {
                string Id = gridView1.GetFocusedDataRow()[6].ToString();
                date = gridView1.GetFocusedDataRow()[1].ToString();
                string OwnerID = gridView1.GetFocusedDataRow()[7].ToString();
                string HeadID = gridView1.GetFocusedDataRow()[0].ToString();
                Grid1Id = Convert.ToInt32(Id);
                Grid3Id = Convert.ToInt32(HeadID);
                OwnerId_v2 = Convert.ToInt32(OwnerID);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        public void FormattingExcelCells(Excel.Range range, bool val1, bool val2)
        {
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


        private void gridControl3_MouseClick(object sender, MouseEventArgs e)
        {
            try
            {
                if (e.Button == MouseButtons.Right)
                {
                    contextMenuStrip1.Show(gridControl3, new Point(e.X, e.Y));
                }
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void удалитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                ArrayList rows = new ArrayList();
                List<Object> aList = new List<Object>();
                string ID = string.Empty;

                List<Object> CarnumList = new List<Object>();
                string CarNum = string.Empty;
                
                Int32[] selectedRowHandles = gridView3.GetSelectedRows();

                DateTime TodayDate = DateTime.Today;
                DateTime oDate = Convert.ToDateTime(date);
                int InTime = (TodayDate - oDate).Days;

                if (selectedRowHandles.Length < gridView3.RowCount && InTime<=14)
                {
                    for (int i = 0; i < selectedRowHandles.Length; i++)
                    {
                        int selectedRowHandle = selectedRowHandles[i];
                        if (selectedRowHandle >= 0)
                            rows.Add(gridView3.GetDataRow(selectedRowHandle));
                    }

                    foreach (DataRow row in rows)
                    {
                        aList.Add(row["ID"]);
                        ID = string.Join(" ", aList);

                        CarnumList.Add(row["Номер В/Ц"]);
                        CarNum = string.Join(" ", CarnumList);

                        string delete = "exec dbo.Remove_Rent_Carriage '" + ID + "','" + CarNum + "'";
                        DbConnection.DBConnect(delete);
                    }

                    gridView2_RowCellClick(null, null);
                }
                else
                {
                    if (selectedRowHandles.Length < gridView3.RowCount && InTime <= 14)
                    {
                        MessageBox.Show("В заявке должен быть хотябы 1 вагон", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else
                    {
                        MessageBox.Show("Вагоны находящиеся в заявке сроком более 14 дней запрещены для удаления", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void создатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                ArrayList rows = new ArrayList();

                List<Object> aList = new List<Object>();
                string ID = string.Empty;

                List<Object> CarnumList = new List<Object>();
                string CarNum = string.Empty;

                Int32[] selectedRowHandles = gridView3.GetSelectedRows();

                DateTime TodayDate = DateTime.Today;
                DateTime oDate = Convert.ToDateTime(date);
                int InTime = (TodayDate - oDate).Days;

                if (selectedRowHandles.Length < gridView3.RowCount && InTime <=14)
                {
                    for (int i = 0; i < selectedRowHandles.Length; i++)
                    {
                        int selectedRowHandle = selectedRowHandles[i];
                        if (selectedRowHandle >= 0)
                            rows.Add(gridView3.GetDataRow(selectedRowHandle));
                    }

                    foreach (DataRow row in rows)
                    {
                        aList.Add(row["ID"]);
                        ID = string.Join(" ", aList);

                        CarnumList.Add(row["Номер В/Ц"]);
                        CarNum = string.Join(" ", CarnumList);

                        Clipboard.SetText(CarNum);

                        string delete = "exec dbo.Remove_Rent_Carriage '" + ID + "','" + CarNum + "'";
                        DbConnection.DBConnect(delete);
                    }

                    foreach (Form form in Application.OpenForms)
                    {
                        if (form.GetType() == typeof(Change_of_Ownership))
                        {
                            form.Activate();
                            return;
                        }
                    }

                    string DateIns = System.DateTime.Now.ToString();

                    Change_of_Ownership change_Of_Ownership = new Change_of_Ownership(tabControl1, role, User_ID);
                    tabControl1.Show();
                    TabPage chg_tabPage = new TabPage("Смена собственника от " + DateIns);
                    change_Of_Ownership.GetDate = DateIns;
                    tabControl1.TabPages.Add(chg_tabPage);
                    tabControl1.SelectedTab = chg_tabPage;
                    change_Of_Ownership.TopLevel = false;
                    change_Of_Ownership.Visible = true;
                    change_Of_Ownership.FormBorderStyle = FormBorderStyle.None;
                    change_Of_Ownership.Dock = DockStyle.Fill;
                    chg_tabPage.Controls.Add(change_Of_Ownership);
                    change_Of_Ownership.вставитьToolStripMenuItem_Click(null,null);

                    gridView2_RowCellClick(null, null);
                }
                else
                {
                    if (selectedRowHandles.Length < gridView3.RowCount && InTime <= 14)
                    {
                        MessageBox.Show("В заявке должен быть хотябы 1 вагон", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else
                    {
                        MessageBox.Show("Вагоны находящиеся в заявке сроком более 14 дней запрещены для создания заявки", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void gridControl2_MouseClick(object sender, MouseEventArgs e)
        {
            try
            {
                if (e.Button == MouseButtons.Right)
                {
                    contextMenuStrip2.Show(gridControl2, new Point(e.X, e.Y));
                }
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void изменитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DateTime TodayDate = DateTime.Today;
            DateTime oDate = Convert.ToDateTime(date);
            int InTime = (TodayDate - oDate).Days;

            if (gridView2.SelectedRowsCount > 0 && InTime <= 14)
            {
                Rent_Update_v1 update_v1 = new Rent_Update_v1(User_ID);
                update_v1.dateTimePicker1.Text = gridView2.GetFocusedDataRow()[1].ToString();
                update_v1.textEdit1.Text = gridView2.GetFocusedDataRow()[2].ToString();
                update_v1.SelectOwnerID = OwnerIDS;
                update_v1.SelectedID = Grid2Id;
                update_v1.Owner = this;
                update_v1.ShowDialog();
            }
            else
            {  
                if (gridView2.SelectedRowsCount > 0 && InTime <= 14)
                {
                    MessageBox.Show("Для редактирования записи, необходимо выбрать строку полностью!", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    MessageBox.Show("Заявки созданные более 14 дней запрещены для редактирования", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void удалитьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                DateTime TodayDate = DateTime.Today;
                DateTime oDate = Convert.ToDateTime(date);
                int InTime = (TodayDate - oDate).Days;

                if (gridView2.SelectedRowsCount > 0 && InTime <=14)
                {
                    if (MessageBox.Show("Удалить выделенные записи?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        ArrayList rows = new ArrayList();
                        List<object> aList = new List<Object>();
                        string Arrays = string.Empty;

                        Int32[] selectedRowHandles = gridView2.GetSelectedRows();
                        for (int i = 0; i < selectedRowHandles.Length; i++)
                        {
                            int selectedRowHandle = selectedRowHandles[i];
                            if (selectedRowHandle >= 0)
                                rows.Add(gridView2.GetDataRow(selectedRowHandle));
                        }
                        
                        foreach (DataRow row in rows)
                        {
                            aList.Add(row["ID"]);
                            Arrays = string.Join(" ", aList);
                            string delete = "exec dbo.Rent_Delete_v1 '" + Arrays + "'";
                            DbConnection.DBConnect(delete);
                        }
                        simpleButton1_Click(null, null);
                    }
                }
                else
                {
                    if (gridView2.SelectedRowsCount > 0 && InTime <= 14) {
                        MessageBox.Show("Для удаления записи, необходимо выбрать строку полностью!", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else
                    {
                        MessageBox.Show("Заявки созданные более 14 дней запрещены для удаления", "", MessageBoxButtons.OK, MessageBoxIcon.Warning); 
                    }
                }
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void отчетToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (gridView2.RowCount > 1)
                {
                    Excel.Application xlApp;
                    Excel.Workbook xlWorkBook;
                    Excel.Worksheet xlWorkSheet;
                    object misValue = System.Reflection.Missing.Value;

                    xlApp = new Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Add(misValue);
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                    xlWorkSheet.Range["B2"].Value = "Собственник:  " + OwnerName;

                    xlWorkSheet.Range["B4"].Value = "Заявка №: " + Grid2Id.ToString() + " от " + date;

                    //xlWorkSheet.Range["B5"].Value = "Продукт: " + product;

                    for (int i = 0; i < gridView3.RowCount; i++)
                    {
                        xlWorkSheet.Cells[i + 8, 2] = i;
                        xlWorkSheet.Cells[i + 8, 3] = gridView3.GetRowCellValue(i, "Номер В/Ц");
                        xlWorkSheet.Cells[i + 8, 4] = gridView3.GetRowCellValue(i, "Продукт");
                    }

                    Excel.Range range = xlWorkSheet.Range["B2", xlWorkSheet.Cells[gridView3.RowCount + 8, 4]];
                    FormattingExcelCells(range, false, false);

                    xlApp.DisplayAlerts = false;
                    xlWorkBook.SaveAs(AppDomain.CurrentDomain.BaseDirectory + @"Report\Заявка смена собственника.xlsx", Excel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    xlWorkBook.Close(true, misValue, misValue);
                    xlApp.Quit();

                    releaseObject(xlWorkSheet);
                    releaseObject(xlWorkBook);
                    releaseObject(xlApp);

                    Process.Start(AppDomain.CurrentDomain.BaseDirectory + @"Report\Заявка смена собственника.xlsx");
                }
                else
                {
                    MessageBox.Show("Выполните поиск", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void изменитьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            DateTime TodayDate = DateTime.Today;
            DateTime oDate = Convert.ToDateTime(date);
            int InTime = (TodayDate - oDate).Days;

            if (gridView1.SelectedRowsCount > 0 && InTime <= 14)
            {
                Rent_Update_v2 update_v2 = new Rent_Update_v2(User_ID);
                update_v2.textEdit3.Text = gridView1.GetFocusedDataRow()[3].ToString();
                update_v2.dateTimePicker1.Text = gridView1.GetFocusedDataRow()[1].ToString();
                update_v2.textEdit1.Text = gridView1.GetFocusedDataRow()[2].ToString();
                update_v2.textEdit2.Text = gridView1.GetFocusedDataRow()[4].ToString();
                update_v2.HeadID = Grid3Id;
                update_v2.BodyID = Grid1Id;
                update_v2.SelectOwnerID = OwnerId_v2;
                update_v2.Owner = this;
                update_v2.ShowDialog();
            }
            else
            {
                if(gridView1.SelectedRowsCount > 0 && InTime <= 14)
                {
                    MessageBox.Show("Для редактирования записи, необходимо выбрать строку полностью!", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    MessageBox.Show("Записи сроком более 14 дней запрещены для редактирования!", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void удалитьToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            try
            {
                DateTime TodayDate = DateTime.Today;
                DateTime oDate = Convert.ToDateTime(date);
                int InTime = (TodayDate - oDate).Days;

                if (gridView1.SelectedRowsCount > 0 && InTime <= 14)
                {
                    if (MessageBox.Show("Удалить выделенные записи?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        ArrayList rows = new ArrayList();
                        List<object> aList = new List<Object>();
                        string Arrays = string.Empty;

                        Int32[] selectedRowHandles = gridView1.GetSelectedRows();
                        for (int i = 0; i < selectedRowHandles.Length; i++)
                        {
                            int selectedRowHandle = selectedRowHandles[i];
                            if (selectedRowHandle >= 0)
                                rows.Add(gridView1.GetDataRow(selectedRowHandle));
                        }
                        foreach (DataRow row in rows)
                        {
                            aList.Add(row["BodyID"]);
                            Arrays = string.Join(" ", aList);
                            string delete = "exec dbo.Rent_Delete_v2 '" + Arrays + "'";
                            DbConnection.DBConnect(delete);
                        }
                        simpleButton1_Click(null, null);
                    }
                }
                else
                {
                    if(gridView1.SelectedRowsCount > 0 && InTime <= 14)
                    {
                        MessageBox.Show("Для удаления записи, необходимо выбрать строку полностью!", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else
                    {
                        MessageBox.Show("Заявки сроком более 14 дней запрещены!", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void отчётToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (gridView1.RowCount >= 1)
                {
                    Excel.Application xlApp;
                    Excel.Workbook xlWorkBook;
                    Excel.Worksheet xlWorkSheet;
                    object misValue = System.Reflection.Missing.Value;

                    xlApp = new Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Add(misValue);
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                    xlWorkSheet.Range["B2"].Value = "История передачи собственникам в/ц №:  " + SelectItemRow2.ToString();


                    xlWorkSheet.Range["B4"].Value = "Дата";
                    xlWorkSheet.Range["C4"].Value = "Номер заявки";
                    xlWorkSheet.Range["D4"].Value = "Номер вагона";
                    xlWorkSheet.Range["E4"].Value = "Продукт";
                    xlWorkSheet.Range["F4"].Value = "Получивший собственник";



                    for (int i = 0; i < gridView1.RowCount; i++)
                    {
                        xlWorkSheet.Cells[i + 5, 2] = gridView1.GetRowCellValue(i, "Дата");
                        xlWorkSheet.Cells[i + 5, 3] = gridView1.GetRowCellValue(i, "Номер заявки");
                        xlWorkSheet.Cells[i + 5, 4] = gridView1.GetRowCellValue(i, "Номер вагона");
                        xlWorkSheet.Cells[i + 5, 5] = gridView1.GetRowCellValue(i, "Продукт");
                        xlWorkSheet.Cells[i + 5, 6] = gridView1.GetRowCellValue(i, "Получивший собственник");
                    }

                    Excel.Range range = xlWorkSheet.Range["B2", xlWorkSheet.Cells[gridView3.RowCount + 6, 6]];
                    FormattingExcelCells(range, false, false);

                    xlApp.DisplayAlerts = false;
                    xlWorkBook.SaveAs(AppDomain.CurrentDomain.BaseDirectory + @"Report\Архив Смена собственника.xlsx", Excel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    xlWorkBook.Close(true, misValue, misValue);
                    xlApp.Quit();

                    releaseObject(xlWorkSheet);
                    releaseObject(xlWorkBook);
                    releaseObject(xlApp);

                    Process.Start(AppDomain.CurrentDomain.BaseDirectory + @"Report\Архив Смена собственника.xlsx");

                }
                else
                {
                    MessageBox.Show("Выполните поиск", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void gridControl1_MouseClick(object sender, MouseEventArgs e)
        {
            try
            {
                if (e.Button == MouseButtons.Right)
                {
                    contextMenuStrip3.Show(gridControl1, new Point(e.X, e.Y));
                }
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void изменитьToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            try
            {
                ArrayList rows = new ArrayList();
                List<Object> aList = new List<Object>();
                string ID = string.Empty;

                List<Object> CarnumList = new List<Object>();
                string CarNum = string.Empty;

                List<Object> ProductList = new List<Object>();
                string Plist = string.Empty;

                Int32[] selectedRowHandles = gridView3.GetSelectedRows();

                DateTime TodayDate = DateTime.Today;
                DateTime oDate = Convert.ToDateTime(date);
                int InTime = (TodayDate - oDate).Days;

                if (InTime <= 14)
                {
                    for (int i = 0; i < selectedRowHandles.Length; i++)
                    {
                        int selectedRowHandle = selectedRowHandles[i];
                        if (selectedRowHandle >= 0)
                            rows.Add(gridView3.GetDataRow(selectedRowHandle));
                    }

                    foreach (DataRow row in rows)
                    {
                        aList.Add(row["ID"]);
                        ID = string.Join(" ", aList);

                        CarnumList.Add(row["Номер В/Ц"]);
                        CarNum = string.Join(" ", CarnumList);
                        
                    }

                    Update_Product_v1 update_Product = new Update_Product_v1(ID, CarNum);
                    update_Product.Owner = this;
                    update_Product.ShowDialog();
                }
                else
                {
                    MessageBox.Show("Вагоны находящиеся в заявке сроком более 14 дней запрещены для изменения", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void добавитьВагонToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                DateTime TodayDate = DateTime.Today;
                DateTime oDate = Convert.ToDateTime(date);
                int InTime = (TodayDate - oDate).Days;
                DataTable ch;
                string checkCar = string.Empty;

                if (InTime <= 14)
                {

                    ArrayList Carlist = new ArrayList();
                    string Arr = string.Empty;
                    Arr = Clipboard.GetText();

                    string CarArray = string.Empty;

                    string[] word = Arr.Split(new char[] { ' ', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

                    foreach (string s in word)
                    {
                        Carlist.Add(s);
                    }
                    
                    for (int i = 0; i < Carlist.Count; i++)
                    {
                        string checkCarriage = "exec dbo.CheckCarnumber_v2 '" + Carlist[i] + "'";
                        ch = DbConnection.DBConnect(checkCarriage);
                        checkCar = ch.Rows[0][0].ToString();

                        if(checkCar == "0")
                        { 
                            CarArray = string.Join(" ", Carlist[i]);
                        
                            string newRow = "exec dbo.Rent_ADD_Body '" + CarArray + "','" + null + "','" + User_ID + "'," + Grid2Id;
                            DbConnection.DBConnect(newRow);
                        }
                        else
                        {
                            MessageBox.Show("Вагон " + Carlist[i] + " не прошел проверку на правильность. Проверьте правильность вагона.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                            continue;
                        }
                    }

                    Clipboard.Clear();
                    gridView2_RowCellClick(null, null);
                }
                else
                {
                    MessageBox.Show("В заявку сроком более 14 дней запрещено добавлять вагоны", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void textEdit3_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && number != 32 && number != 8 && (e.KeyChar <= 39 || e.KeyChar >= 46) && number != 47 && number != 61 && (e.KeyChar < 'А' || e.KeyChar > 'я'))
            {
                e.Handled = true;
            }
        }

        private void textEdit2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && number != 8)
            {
                e.Handled = true;
            }
        }

        private void gridView1_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            GridView View = sender as GridView;

            if (View.IsRowSelected(e.RowHandle))
            {
                e.Appearance.ForeColor = Color.DarkBlue;
                e.Appearance.BackColor = Color.LightBlue;
                //e.HighPriority = true;
            }
        }

        private void gridView2_RowStyle(object sender, RowStyleEventArgs e)
        {
            GridView View = sender as GridView;

            if (View.IsRowSelected(e.RowHandle))
            {
                e.Appearance.ForeColor = Color.DarkBlue;
                e.Appearance.BackColor = Color.LightBlue;
                //e.HighPriority = true;
            }
        }

        private void gridView3_RowStyle(object sender, RowStyleEventArgs e)
        {
            GridView View = sender as GridView;

            if (View.IsRowSelected(e.RowHandle))
            {
                e.Appearance.ForeColor = Color.DarkBlue;
                e.Appearance.BackColor = Color.LightBlue;
                e.HighPriority = true;
            }
        }


        private void Context1_Visible(bool val)
        {
            contextMenuStrip1.Items[0].Enabled = val;
            contextMenuStrip1.Items[1].Enabled = val;
            contextMenuStrip1.Items[2].Enabled = val;
            contextMenuStrip1.Items[3].Enabled = val;
        }

    }
}
