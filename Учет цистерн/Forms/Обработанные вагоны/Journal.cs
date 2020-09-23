using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraGrid.Views.Grid;
using Учет_цистерн.Forms.заявки_на_обработку;

namespace Учет_цистерн.Forms.Обработанные_вагоны
{
    public partial class Journal : Form
    {
        int SelectItemRow;
        public string SelectBrigadeID { get; set; }
        public string SelectProductID { get; set; }
        string role, User_ID;
        int Temp;
        int TempUpdate = -1;
        int GroupUpdate = -1;
        int UpdateAutn = -1;
        int Num;

        public Journal(string role, string UserID)
        {
            InitializeComponent();
            this.role = role;
            this.User_ID = UserID;
        }

        private void Journal_Load(object sender, EventArgs e)
        {
            try
            {
                if (role == "1" | role == "1002")
                {
                    simpleButton4.Visible = true;
                }
                else
                {
                    if (role == "2")
                    {
                        simpleButton4.Visible = false;
                    }
                    else
                    {
                        checkEdit24.Visible = true;
                        simpleButton4.Visible = false;
                        simpleButton2.Enabled = false;
                    }
                }

                DateTime now = DateTime.Now;
                dateTimePicker1.Value = now;

                Refreshh("2");
                Fillcombobox();
                Block();

                simpleButton9.Enabled = false;
                panel1.Visible = false;
                label1.Visible = false;
                memoEdit1.Visible = false;
                simpleButton7.Visible = false;
                simpleButton8.Visible = false;
                simpleButton11.Visible = false;

                if (gridView1.RowCount == 0)
                {
                    textEdit5.Text = "0";
                    textEdit6.Text = "0";
                    textEdit7.Text = "0";
                    textEdit8.Text = "0";
                    textEdit9.Text = "0";
                    textEdit10.Text = "0";
                    textEdit11.Text = "0";
                    textEdit25.Text = "0";
                    //textEdit12.Text = "0";
                    //textEdit13.Text = "0";
                    //textEdit14.Text = "0";
                    //textEdit15.Text = "0";
                    //textEdit16.Text = "0";
                    //textEdit17.Text = "0";
                    //textEdit18.Text = "0";
                    //textEdit19.Text = "0";
                    //textEdit20.Text = "0";
                    //textEdit21.Text = "0";
                    //textEdit22.Text = "0";
                    //textEdit23.Text = "0";
                }

                if (gridView1.RowCount > 0)
                {
                    gridView1_RowCellClick(null, null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void Refreshh(string Section)
        {
            if (role == "1" | role == "2" | role == "1002")
            {
                gridControl1.DataSource = null;
                gridView1.Columns.Clear();

                string Refreshh = "exec [dbo].[GetRenderedService] '" + dateTimePicker1.Value.ToShortDateString() + "', " + "@Type = " + 1;
                DataTable dt = DbConnection.DBConnect(Refreshh);
                gridControl1.DataSource = dt;
                gridView1.BestFitColumns();
                gridView1.Columns[3].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
                gridView1.Columns[0].Visible = false;
                gridView1.Columns[1].Visible = false;
                gridView1.Columns[2].Visible = false;

                gridView1.Columns[16].DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
                gridView1.Columns[16].DisplayFormat.FormatString = "dd.MM.yyyy HH:mm";
                gridView1.Columns[17].DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
                gridView1.Columns[17].DisplayFormat.FormatString = "dd.MM.yyyy HH:mm";
                gridView1.Columns[18].DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
                gridView1.Columns[18].DisplayFormat.FormatString = "dd.MM.yyyy HH:mm";
                gridView1.Columns[19].DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
                gridView1.Columns[19].DisplayFormat.FormatString = "dd.MM.yyyy HH:mm";
                gridView1.Columns[20].DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
                gridView1.Columns[20].DisplayFormat.FormatString = "dd.MM.yyyy HH:mm";
                gridView1.Columns[21].DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
                gridView1.Columns[21].DisplayFormat.FormatString = "dd.MM.yyyy HH:mm";
                gridView1.Columns[22].DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
                gridView1.Columns[22].DisplayFormat.FormatString = "dd.MM.yyyy HH:mm";

                if (Section == "2")
                {
                    gridView1.MoveLast();
                }
                else
                {
                    if (Section == "1")
                    {
                        ColumnView View = (ColumnView)gridControl1.FocusedView;
                        // obtaining the column bound to the Country field 
                        GridColumn column = View.Columns["ID"];
                        if (column != null)
                        {
                            int rhFound = View.LocateByDisplayText(View.FocusedRowHandle + 1, column, Convert.ToString(SelectItemRow));
                            View.FocusedRowHandle = rhFound;
                            View.FocusedColumn = column;
                        }
                    }
                    
                }
                

                GridColumnSummaryItem Carnumber = new GridColumnSummaryItem(DevExpress.Data.SummaryItemType.Count, "Номер вагона", "{0}");
                GridColumnSummaryItem ServiceCost = new GridColumnSummaryItem(DevExpress.Data.SummaryItemType.Sum, "Сумма услуг", "{0}");
                GridColumnSummaryItem TorCost = new GridColumnSummaryItem(DevExpress.Data.SummaryItemType.Sum, "Сумма ТОР", "{0}");
                gridView1.Columns["Номер вагона"].Summary.Add(Carnumber);
                gridView1.Columns["Сумма услуг"].Summary.Add(ServiceCost);
                gridView1.Columns["Сумма ТОР"].Summary.Add(TorCost);
            }
            else
            {
                if (role == "3")
                {
                    gridControl1.DataSource = null;
                    gridView1.Columns.Clear();

                    string Refreshh = "exec [dbo].[GetRenderedService] '" + dateTimePicker1.Value.ToShortDateString() + "', " + "@Type = " + 2;
                    DataTable dt = DbConnection.DBConnect(Refreshh);
                    gridControl1.DataSource = dt;
                    gridView1.BestFitColumns();
                    gridView1.Columns[3].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
                    gridView1.Columns[0].Visible = false;
                    gridView1.Columns[1].Visible = false;
                    gridView1.Columns[2].Visible = false;

                    gridView1.Columns[19].DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
                    gridView1.Columns[19].DisplayFormat.FormatString = "dd.MM.yyyy HH:mm";
                    gridView1.Columns[20].DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
                    gridView1.Columns[20].DisplayFormat.FormatString = "dd.MM.yyyy HH:mm";
                    gridView1.Columns[21].DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
                    gridView1.Columns[21].DisplayFormat.FormatString = "dd.MM.yyyy HH:mm";
                    gridView1.Columns[22].DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
                    gridView1.Columns[22].DisplayFormat.FormatString = "dd.MM.yyyy HH:mm";
                    gridView1.Columns[23].DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
                    gridView1.Columns[23].DisplayFormat.FormatString = "dd.MM.yyyy HH:mm";
                    gridView1.Columns[24].DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
                    gridView1.Columns[24].DisplayFormat.FormatString = "dd.MM.yyyy HH:mm";
                    gridView1.Columns[25].DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
                    gridView1.Columns[25].DisplayFormat.FormatString = "dd.MM.yyyy HH:mm";

                    if (Section == "2")
                    {
                        gridView1.MoveLast();
                    }
                    else
                    {
                        if (Section == "1")
                        {
                            ColumnView View = (ColumnView)gridControl1.FocusedView;
                            // obtaining the column bound to the Country field 
                            GridColumn column = View.Columns["ID"];
                            gridView1.FocusedRowHandle = SelectItemRow;
                            if (column != null)
                            {
                                int rhFound = View.LocateByDisplayText(View.FocusedRowHandle + 1, column, Convert.ToString(SelectItemRow));
                                View.FocusedRowHandle = SelectItemRow;
                                View.FocusedColumn = column;
                            }

                        }
                    }


                        GridColumnSummaryItem Carnumber = new GridColumnSummaryItem(DevExpress.Data.SummaryItemType.Count, "Номер вагона", "{0}");
                    GridColumnSummaryItem ServiceCost = new GridColumnSummaryItem(DevExpress.Data.SummaryItemType.Sum, "Сумма услуг", "{0}");
                    GridColumnSummaryItem TorCost = new GridColumnSummaryItem(DevExpress.Data.SummaryItemType.Sum, "Сумма ТОР", "{0}");
                    gridView1.Columns["Номер вагона"].Summary.Add(Carnumber);
                    gridView1.Columns["Сумма услуг"].Summary.Add(ServiceCost);
                    gridView1.Columns["Сумма ТОР"].Summary.Add(TorCost);
                }

            }
            
            //gridView1.ShowFindPanel();
            //textEdit1.Text = "";
            //textEdit2.Text = "";
            //textEdit3.Text = "";
            //textEdit4.Text = "";
            //textEdit5.Text = "0";
            //textEdit6.Text = "0";
            //textEdit7.Text = "0";
            //textEdit8.Text = "0";
            //textEdit9.Text = "0";
            //textEdit10.Text = "0";
            //textEdit11.Text = "0";
        }

        private void Fillcombobox()
        {
            string Brigade = "Select * from d__Brigade";
            DataTable dt = DbConnection.DBConnect(Brigade);
            comboBox1.DataSource = dt;
            comboBox1.DisplayMember = "Surname";
            comboBox1.ValueMember = "ID";

            string Product = "Select * from d__Product";
            DataTable dt1 = DbConnection.DBConnect(Product);
            comboBox2.DataSource = dt1;
            comboBox2.DisplayMember = "Name";
            comboBox2.ValueMember = "ID";
        }

        private void LastRenderedService()
        {
            string GetParameter = "select Parameter from d__Parameter";
            DataTable param = DbConnection.DBConnect(GetParameter);
            string Params = param.Rows[0][0].ToString();

            string query = "select [dbo].[CheckCarService] (" + textEdit1.Text.Trim() + ")";
            DataTable dt1 = DbConnection.DBConnect(query);
            int State = Convert.ToInt32(dt1.Rows[0][0].ToString());
            if (State == 1)
            {
                panel1.Visible = true;
                simpleButton7.Visible = true;
                label1.Visible = true;
                label1.Text = "Данная в/ц проходила обработку в течении последних " + Params + " дней";
            }
            else
            {
                panel1.Visible = false;
                label1.Visible = false;
                simpleButton7.Visible = false;
            }

            string LastRent = "exec dbo.LastRent " + textEdit1.Text.Trim();
            DataTable dt2 = DbConnection.DBConnect(LastRent);
            if (dt2.Rows.Count > 0 && State != 1)
            {
                memoEdit1.Location = panel1.Location;
                memoEdit1.Visible = true;
                memoEdit1.Text = "Последняя заявка: " + dt2.Rows[0][1] + " от " + dt2.Rows[0][2] + "" + "\r\n" + "Продукт: " + dt2.Rows[0][5] + "" + "\r\n" + "Была передача: " + dt2.Rows[0][3];
            }
            else if (dt2.Rows.Count > 0 && State == 1)
            {
                memoEdit1.Location = new Point(8, 595);
                memoEdit1.Visible = true;
                memoEdit1.Text = "Последняя заявка: " + dt2.Rows[0][1] + " от " + dt2.Rows[0][2] + "" + "\r\n" + "Продукт: " + dt2.Rows[0][5] + "" + "\r\n" + "Была передача: " + dt2.Rows[0][3];
            }
            else
            {
                memoEdit1.Location = new Point(8,595);
                memoEdit1.Visible = false;
            }
        }

        private void LastRenderedService_1()
        {
            string GetParameter = "select Parameter from d__Parameter";
            DataTable param = DbConnection.DBConnect(GetParameter);
            string Params = param.Rows[0][0].ToString();

            string query = "select [dbo].[CheckCarService_1] (" + textEdit1.Text.Trim() + "," + SelectItemRow + ")";
            DataTable dt1 = DbConnection.DBConnect(query);
            int State = Convert.ToInt32(dt1.Rows[0][0].ToString());
            if (State == 1)
            {
                panel1.Visible = true;
                simpleButton7.Visible = true;
                label1.Visible = true;
                label1.Text = "Данная в/ц проходила обработку в течении последних " + Params + " дней";
            }
            else
            {
                panel1.Visible = false;
                label1.Visible = false;
                simpleButton7.Visible = false;
            }

            string LastRent = "exec dbo.LastRent " + textEdit1.Text.Trim();
            DataTable dt2 = DbConnection.DBConnect(LastRent);
            if (dt2.Rows.Count > 0 && State != 1)
            {
                memoEdit1.Location = panel1.Location;
                memoEdit1.Visible = true;
                memoEdit1.Text = "Последняя заявка: " + dt2.Rows[0][1] + " от " + dt2.Rows[0][2] + "" + "\r\n" + "Продукт: " + dt2.Rows[0][5] + "" + "\r\n" + "Была передача: " + dt2.Rows[0][3];
            }
            else if (dt2.Rows.Count > 0 && State == 1)
            {
                memoEdit1.Location = new Point(8, 595);
                memoEdit1.Visible = true;
                memoEdit1.Text = "Последняя заявка: " + dt2.Rows[0][1] + " от " + dt2.Rows[0][2] + "" + "\r\n" + "Продукт: " + dt2.Rows[0][5] + "" + "\r\n" + "Была передача: " + dt2.Rows[0][3];
            }
            else
            {
                memoEdit1.Location = new Point(8, 595);
                memoEdit1.Visible = false;
            }
        }

        private void textEdit1_Properties_EditValueChanged(object sender, EventArgs e)
        {
            try
            {
                if (textEdit1.Text != string.Empty)
                {
                    string CheckOwner = "select c.Current_owner from d__Carriage c where c.CarNumber = " + textEdit1.Text.Trim();
                    DataTable dt = DbConnection.DBConnect(CheckOwner);
                    if (dt.Rows.Count > 0)
                    {
                        textEdit2.Text = dt.Rows[0][0].ToString();
                    }
                    else
                    {
                        textEdit2.Text = "";
                    }

                    if (textEdit1.Enabled == false)
                    {
                        LastRenderedService_1();
                    }
                    else if(SelectItemRow == 0)
                    {
                        LastRenderedService();
                    }
                    else if(SelectItemRow > 1 && textEdit1.Enabled == true)
                    {
                        LastRenderedService_1();
                    }

                    //if(dt.Rows.Count > 0)
                    //{
                    //    if (textEdit1.Text.Length > 7 && dt.Rows[0][0].ToString() == "Премиум Ойл Транс ТОО")
                    //    {
                    //        if (SelectItemRow == 0)
                    //        {
                    //            textEdit12.Enabled = true;
                    //            textEdit13.Enabled = true;
                    //            textEdit14.Enabled = true;
                    //            textEdit15.Enabled = true;
                    //            textEdit16.Enabled = true;
                    //            textEdit17.Enabled = true;
                    //            textEdit18.Enabled = true;
                    //            textEdit19.Enabled = true;
                    //            textEdit20.Enabled = true;
                    //            textEdit21.Enabled = true;
                    //            textEdit22.Enabled = true;
                    //            textEdit23.Enabled = true;
                    //            textEdit24.Enabled = true;

                    //            textEdit12.Text = "";
                    //            textEdit13.Text = "";
                    //            textEdit14.Text = "";
                    //            textEdit15.Text = "";
                    //            textEdit16.Text = "";
                    //            textEdit17.Text = "";
                    //            textEdit18.Text = "";
                    //            textEdit19.Text = "";
                    //            textEdit20.Text = "";
                    //            textEdit21.Text = "";
                    //            textEdit22.Text = "";
                    //            textEdit23.Text = "";
                    //            textEdit24.Text = "";
                    //        }
                    //        else if (TempUpdate == 0)
                    //        {
                    //            textEdit12.Enabled = true;
                    //            textEdit13.Enabled = true;
                    //            textEdit14.Enabled = true;
                    //            textEdit15.Enabled = true;
                    //            textEdit16.Enabled = true;
                    //            textEdit17.Enabled = true;
                    //            textEdit18.Enabled = true;
                    //            textEdit19.Enabled = true;
                    //            textEdit20.Enabled = true;
                    //            textEdit21.Enabled = true;
                    //            textEdit22.Enabled = true;
                    //            textEdit23.Enabled = true;
                    //            textEdit24.Enabled = true;
                    //        }
                    //    }
                    //    else
                    //    {
                    //        textEdit12.Enabled = false;
                    //        textEdit13.Enabled = false;
                    //        textEdit14.Enabled = false;
                    //        textEdit15.Enabled = false;
                    //        textEdit16.Enabled = false;
                    //        textEdit17.Enabled = false;
                    //        textEdit18.Enabled = false;
                    //        textEdit19.Enabled = false;
                    //        textEdit20.Enabled = false;
                    //        textEdit21.Enabled = false;
                    //        textEdit22.Enabled = false;
                    //        textEdit23.Enabled = false;
                    //        textEdit24.Enabled = false;

                    //        textEdit12.Text = "";
                    //        textEdit13.Text = "";
                    //        textEdit14.Text = "";
                    //        textEdit15.Text = "";
                    //        textEdit16.Text = "";
                    //        textEdit17.Text = "";
                    //        textEdit18.Text = "";
                    //        textEdit19.Text = "";
                    //        textEdit20.Text = "";
                    //        textEdit21.Text = "";
                    //        textEdit22.Text = "";
                    //        textEdit23.Text = "";
                    //        textEdit24.Text = "";
                    //    }
                    //}
                }
                else
                {
                    panel1.Visible = false;
                    label1.Visible = false;
                    memoEdit1.Visible = false;
                    simpleButton7.Visible = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            try
            {
                //Добавить
                if (SelectItemRow == 0)
                {
                    if (textEdit1.Text != string.Empty)
                    {
                        string CheckOwner = "select c.Current_owner from d__Carriage c where c.CarNumber = " + textEdit1.Text.Trim();
                        DataTable dt = DbConnection.DBConnect(CheckOwner);
                        string Check = dt.Rows[0][0].ToString();
                        if (Check != "")
                        {
                            if (textEdit1.Text.Length > 7 && dt.Rows[0][0].ToString() == "PTC Operator TOO")
                            {
                                if ((textEdit3.Text != "" && textEdit4.Text != "" && textEdit6.Text != "" && textEdit7.Text != "" && textEdit8.Text != "" && textEdit11.Text != "" && textEdit25.Text != "") &
                                    (textEdit6.Text != "0" || textEdit7.Text != "0" || textEdit8.Text != "0" || textEdit11.Text != "0" || textEdit25.Text != ""))
                                {
                                    string StartProcess, EndProcess, DateTehCart, StartTor, EndTor, ArrivalDate, DateVU19 = string.Empty;

                                    if (checkEdit25.Checked) { StartProcess = dateTimePicker9.Value.ToString(); } else { StartProcess = "null"; }
                                    if (checkEdit26.Checked) { EndProcess = dateTimePicker8.Value.ToString(); } else { EndProcess = "null"; }
                                    if (checkEdit27.Checked) { DateTehCart = dateTimePicker6.Value.ToString(); } else { DateTehCart = "null"; }
                                    if (checkEdit28.Checked) { StartTor = dateTimePicker2.Value.ToString(); } else { StartTor = "null"; }
                                    if (checkEdit29.Checked) { EndTor = dateTimePicker3.Value.ToString(); } else { EndTor = "null"; }
                                    if (checkEdit30.Checked) { ArrivalDate = dateTimePicker5.Value.ToString(); } else { ArrivalDate = "null"; }
                                    if (checkEdit31.Checked) { DateVU19 = dateTimePicker4.Value.ToString(); } else { DateVU19 = "null"; }

                                    string Add = "" +
                                        "declare @Id int; exec [dbo].[FillRenderedService] " +
                                        "'" + User_ID + "'," +
                                        "" + Num + "," +
                                        "'" + textEdit2.Text.Trim() + "'," +
                                        "'" + dateTimePicker1.Value.Date.ToString() + "'," +
                                        "" + textEdit1.Text.Trim() + "," +
                                        "" + textEdit4.Text.Trim() + "," +
                                        "" + textEdit6.Text.Trim() + "," +
                                        "" + textEdit8.Text.Trim() + "," +
                                        "" + textEdit7.Text.Trim() + "," +
                                        "" + textEdit25.Text.Trim() + "," +
                                        "" + textEdit11.Text.Trim() + "," +
                                        "" + comboBox1.SelectedValue.ToString() + "," +
                                        "'" + textEdit3.Text.Trim() + "'," +
                                        "" + comboBox2.SelectedValue.ToString() + "," +
                                        "NULL, " +
                                        "'" + StartProcess + "', " +
                                        "'" + EndProcess + "', " +
                                        "'" + DateTehCart + "'," +
                                        "'" + StartTor + "', " +
                                        "'" + EndTor + "', " +
                                        "'" + ArrivalDate + "'," +
                                        "'" + DateVU19 + "',  @CurrentID = @Id output; select @Id";

                                    DataTable HeadID = DbConnection.DBConnect(Add);
                                    if (HeadID.Rows.Count > 0)
                                    {
                                        string Id = HeadID.Rows[0][0].ToString();
                                        string Autn = "exec [dbo].[FillAutn] '" + textEdit18.Text.Trim() + "','" + textEdit12.Text.Trim() + "','" + textEdit13.Text.Trim() + "','" + textEdit14.Text.Trim() + "','" + textEdit24.Text.Trim() + "','" + textEdit15.Text.Trim() + "','" + textEdit16.Text.Trim() + "','" + textEdit23.Text.Trim() + "','" + textEdit22.Text.Trim() + "','" + textEdit21.Text.Trim() + "','" + textEdit20.Text.Trim() + "','" + textEdit19.Text.Trim() + "','" + textEdit17.Text.Trim() + "'," + Id;
                                        DbConnection.DBConnect(Autn);
                                        Refreshh("2");
                                        textEdit1.Text = "";
                                        textEdit1.Focus();
                                    }
                                    //}
                                    //else
                                    //{
                                    //    MessageBox.Show("Заполните пустые поля в АУТН!", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    //}
                                }
                                else
                                {
                                    MessageBox.Show("Заполните пустые поля в Обработке!", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                }
                            }
                            else
                            {
                                if ((textEdit3.Text != "" && textEdit4.Text != "" && textEdit6.Text != "" && textEdit7.Text != "" && textEdit8.Text != "" && textEdit11.Text != "" && textEdit25.Text != "") &
                                    (textEdit6.Text != "0" || textEdit7.Text != "0" || textEdit8.Text != "0" || textEdit11.Text != "0" || textEdit25.Text != ""))
                                {
                                    string StartProcess, EndProcess, DateTehCart, StartTor, EndTor, ArrivalDate, DateVU19 = string.Empty;

                                    if (checkEdit25.Checked) { StartProcess = dateTimePicker9.Value.ToString(); } else { StartProcess = "null"; }
                                    if (checkEdit26.Checked) { EndProcess = dateTimePicker8.Value.ToString(); } else { EndProcess = "null"; }
                                    if (checkEdit27.Checked) { DateTehCart = dateTimePicker6.Value.ToString(); } else { DateTehCart = "null"; }
                                    if (checkEdit28.Checked) { StartTor = dateTimePicker2.Value.ToString(); } else { StartTor = "null"; }
                                    if (checkEdit29.Checked) { EndTor = dateTimePicker3.Value.ToString(); } else { EndTor = "null"; }
                                    if (checkEdit30.Checked) { ArrivalDate = dateTimePicker5.Value.ToString(); } else { ArrivalDate = "null"; }
                                    if (checkEdit31.Checked) { DateVU19 = dateTimePicker4.Value.ToString(); } else { DateVU19 = "null"; }

                                    string Add = "" +
                                        "declare @Id int; exec [dbo].[FillRenderedService] " +
                                        "'" + User_ID + "'," +
                                        "" + Num + "," +
                                        "'" + textEdit2.Text.Trim() + "'," +
                                        "'" + dateTimePicker1.Value.Date.ToString() + "'," +
                                        "" + textEdit1.Text.Trim() + "," +
                                        "" + textEdit4.Text.Trim() + "," +
                                        "" + textEdit6.Text.Trim() + "," +
                                        "" + textEdit8.Text.Trim() + "," +
                                        "" + textEdit7.Text.Trim() + "," +
                                        "" + textEdit25.Text.Trim() + "," +
                                        "" + textEdit11.Text.Trim() + "," +
                                        "" + comboBox1.SelectedValue.ToString() + "," +
                                        "'" + textEdit3.Text.Trim() + "'," +
                                        "" + comboBox2.SelectedValue.ToString() + "," +
                                        "NULL, " +
                                        "'" + StartProcess + "', " +
                                        "'" + EndProcess + "', " +
                                        "'" + DateTehCart + "'," +
                                        "'" + StartTor + "', " +
                                        "'" + EndTor + "', " +
                                        "'" + ArrivalDate + "'," +
                                        "'" + DateVU19 + "',  @CurrentID = @Id output; select @Id";
                                    DataTable HeadID = DbConnection.DBConnect(Add);
                                    if (HeadID.Rows.Count > 0)
                                    {
                                        string Id = HeadID.Rows[0][0].ToString();
                                        string Autn = "exec [dbo].[FillAutn] '" + textEdit18.Text.Trim() + "','" + textEdit12.Text.Trim() + "','" + textEdit13.Text.Trim() + "','" + textEdit14.Text.Trim() + "','" + textEdit24.Text.Trim() + "','" + textEdit15.Text.Trim() + "','" + textEdit16.Text.Trim() + "','" + textEdit23.Text.Trim() + "','" + textEdit22.Text.Trim() + "','" + textEdit21.Text.Trim() + "','" + textEdit20.Text.Trim() + "','" + textEdit19.Text.Trim() + "','" + textEdit17.Text.Trim() + "'," + Id;
                                        DbConnection.DBConnect(Autn);
                                        Refreshh("2");
                                        textEdit1.Text = "";
                                        textEdit1.Focus();
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Заполните пустые поля в Обработке!", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("У данного В/Ц отсутствует собственник!", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                }
                //Редактировать группу
                else if(GroupUpdate == 1)
                {
                    if (textEdit1.Text.Length > 7 && textEdit2.Text == "PTC Operator TOO")
                    {
                        if ((textEdit4.Text != "" && textEdit6.Text != "" && textEdit7.Text != "" && textEdit8.Text != "" && textEdit25.Text != "" && textEdit11.Text != "") &
                                    (textEdit6.Text != "0" || textEdit7.Text != "0" || textEdit8.Text != "0" || textEdit25.Text != "0" || textEdit11.Text != "0"))
                        {
                            string Oper,Brigade,Product,HOL,TOR,GOR,DR,Trafar,
                                   StartProcess,EndProcess,DateTehCart,StartTor,EndTor,ArrivalDate,DateVU19 = string.Empty;
                        
                            if (checkEdit1.Checked){Oper = textEdit4.Text.Trim();}
                            else{Oper = "";}

                            if (checkEdit2.Checked) { Brigade = comboBox1.SelectedValue.ToString(); }
                            else { Brigade = ""; }

                            if (checkEdit11.Checked) { Product = comboBox2.SelectedValue.ToString(); }
                            else { Product = ""; }

                            if (checkEdit4.Checked) { HOL = textEdit6.Text.Trim(); }
                            else { HOL = ""; }

                            if (checkEdit6.Checked) { TOR = textEdit7.Text.Trim(); }
                            else { TOR = ""; }

                            if (checkEdit5.Checked) { GOR = textEdit8.Text.Trim(); }
                            else { GOR = ""; }

                            if (checkEdit32.Checked) { DR = textEdit9.Text.Trim(); }
                            else { DR = ""; }

                            if (checkEdit9.Checked) { Trafar = textEdit11.Text.Trim(); }
                            else { Trafar = ""; }

                            if (checkEdit25.Checked) { StartProcess = dateTimePicker9.Value.ToString(); } else { StartProcess = ""; }
                            if (checkEdit26.Checked) { EndProcess = dateTimePicker8.Value.ToString(); } else { EndProcess = ""; }
                            if (checkEdit27.Checked) { DateTehCart = dateTimePicker6.Value.ToString(); } else { DateTehCart = ""; }
                            if (checkEdit28.Checked) { StartTor = dateTimePicker2.Value.ToString(); } else { StartTor = ""; }
                            if (checkEdit29.Checked) { EndTor = dateTimePicker3.Value.ToString(); } else { EndTor = ""; }
                            if (checkEdit30.Checked) { ArrivalDate = dateTimePicker5.Value.ToString(); } else { ArrivalDate = ""; }
                            if (checkEdit31.Checked) { DateVU19 = dateTimePicker4.Value.ToString(); } else { DateVU19 = ""; }

                            ArrayList rows = new ArrayList();
                            List<Object> aList = new List<Object>();
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
                                aList.Add(row["ID"]);
                                Arrays = string.Join(" ", aList);
                            }
                            string[] split = Arrays.Split(new Char[] { ' ', ',', '.', ':', '\t' });
                            foreach (string s in split)
                            {
                                string UpdateAll = "exec [dbo].[UpdateRenderedServiceAll] " +
                                    "'" + User_ID + "'," +
                                    "'" + Oper + "'," +
                                    "'" + HOL + "'," +
                                    "'" + GOR + "'," +
                                    "'" + TOR + "'," +
                                    "'" + DR + "'," +
                                    "'" + Trafar + "'," +
                                    "'" + Brigade + "'," +
                                    "'" + Product + "', " +
                                    "'" + StartProcess + "'," +
                                    "'" + EndProcess + "'," +
                                    "'" + DateTehCart + "'," +
                                    "'" + StartTor + "'," +
                                    "'" + EndTor + "'," +
                                    "'" + ArrivalDate + "'," +
                                    "'" + DateVU19 + "'," +
                                    "" + Temp + ",'" + s + "'";
                                DbConnection.DBConnect(UpdateAll);
                            }
                            Refreshh("1");
                        }
                        else
                        {
                            MessageBox.Show("Заполните пустые поля в Обработке!", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    else
                    {
                        if ((textEdit4.Text != "" && textEdit6.Text != "" && textEdit7.Text != "" && textEdit8.Text != "" && textEdit25.Text != "" && textEdit11.Text != "") &
                                    (textEdit6.Text != "0" || textEdit7.Text != "0" || textEdit8.Text != "0" || textEdit25.Text != "0" || textEdit11.Text != "0"))
                        {
                            string Oper, Brigade, Product, HOL, TOR, GOR, DR, Trafar,
                                StartProcess, EndProcess, DateTehCart, StartTor, EndTor, ArrivalDate, DateVU19 = string.Empty;

                            if (checkEdit1.Checked) { Oper = textEdit4.Text.Trim(); }
                            else { Oper = ""; }

                            if (checkEdit2.Checked) { Brigade = comboBox1.SelectedValue.ToString(); }
                            else { Brigade = ""; }

                            if (checkEdit11.Checked) { Product = comboBox2.SelectedValue.ToString(); }
                            else { Product = ""; }

                            if(checkEdit4.Checked | checkEdit6.Checked | checkEdit5.Checked | checkEdit32.Checked)
                            {
                                HOL = textEdit6.Text.Trim();
                                TOR = textEdit7.Text.Trim();
                                GOR = textEdit8.Text.Trim();
                                DR = textEdit25.Text.Trim();
                            }
                            else
                            {
                                HOL = "";
                                TOR = "";
                                GOR = "";
                                DR = "";
                            }

                            if (checkEdit9.Checked) { Trafar = textEdit11.Text.Trim(); }
                            else { Trafar = ""; }

                            if (checkEdit25.Checked) { StartProcess = dateTimePicker9.Value.ToString(); } else { StartProcess = ""; }
                            if (checkEdit26.Checked) { EndProcess = dateTimePicker8.Value.ToString(); } else { EndProcess = ""; }
                            if (checkEdit27.Checked) { DateTehCart = dateTimePicker6.Value.ToString(); } else { DateTehCart = ""; }
                            if (checkEdit28.Checked) { StartTor = dateTimePicker2.Value.ToString(); } else { StartTor = ""; }
                            if (checkEdit29.Checked) { EndTor = dateTimePicker3.Value.ToString(); } else { EndTor = ""; }
                            if (checkEdit30.Checked) { ArrivalDate = dateTimePicker5.Value.ToString(); } else { ArrivalDate = ""; }
                            if (checkEdit31.Checked) { DateVU19 = dateTimePicker4.Value.ToString(); } else { DateVU19 = ""; }

                            ArrayList rows = new ArrayList();
                            List<Object> aList = new List<Object>();
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
                                aList.Add(row["ID"]);
                                Arrays = string.Join(" ", aList);
                            }
                            string[] split = Arrays.Split(new Char[] { ' ', ',', '.', ':', '\t' });
                            foreach(string s in split)
                            {
                                string UpdateAll = "exec [dbo].[UpdateRenderedServiceAll] " +
                                    "'" + User_ID + "'," +
                                    "'" + Oper + "'," +
                                    "'" + HOL + "'," +
                                    "'" + GOR + "'," +
                                    "'" + TOR + "'," +
                                    "'" + DR + "'," +
                                    "'" + Trafar + "'," +
                                    "'" + Brigade + "'," +
                                    "'" + Product + "', " +
                                    "'" + StartProcess + "'," +
                                    "'" + EndProcess + "'," +
                                    "'" + DateTehCart + "'," +
                                    "'" + StartTor + "'," +
                                    "'" + EndTor + "'," +
                                    "'" + ArrivalDate + "'," +
                                    "'" + DateVU19 + "'," +
                                    "" + Temp + ",'" + s + "'";
                                DbConnection.DBConnect(UpdateAll);
                            }
                            Refreshh("1");
                        }
                        else
                        {
                            MessageBox.Show("Заполните пустые поля в Обработке!", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                }
                //Редактировать АУТН
                else if (UpdateAutn == 1)
                {
                    string Avail, Klapan, DemSkob, PTC, Ushki, Skobi, Shaiba, Lest, Greben, Barash, TriBolt, Exp, Note = string.Empty;

                    if (checkEdit17.Checked) { Avail = textEdit18.Text.Trim(); } else { Avail = ""; } //Пригодность 
                    if (checkEdit3.Checked) { Klapan = textEdit12.Text.Trim(); } else { Klapan = ""; } //3 кл 
                    if (checkEdit12.Checked) { DemSkob = textEdit13.Text.Trim(); } else { DemSkob = ""; } //дем скобы
                    if (checkEdit13.Checked) { PTC = textEdit14.Text.Trim(); } else { PTC = ""; } //трафар ПТС
                    if (checkEdit14.Checked) { Ushki = textEdit24.Text.Trim(); } else { Ushki = ""; } //Ушки 
                    if (checkEdit15.Checked) { Skobi = textEdit15.Text.Trim(); } else { Skobi = ""; } //Скобы 
                    if (checkEdit22.Checked) { Shaiba = textEdit16.Text.Trim(); } else { Shaiba = ""; } //Шайба и валик 
                    if (checkEdit21.Checked) { Lest = textEdit23.Text.Trim(); } else { Lest = ""; } //ЛЕстница 
                    if (checkEdit20.Checked) { Greben = textEdit22.Text.Trim(); } else { Greben = ""; } //Гребнь 
                    if (checkEdit19.Checked) { Barash = textEdit21.Text.Trim(); } else { Barash = ""; } //ВЦ бараш тип
                    if (checkEdit18.Checked) { TriBolt = textEdit20.Text.Trim(); } else { TriBolt = ""; } //3 болта
                    if (checkEdit16.Checked) { Exp = textEdit19.Text.Trim(); } else { Exp = ""; } //годный на эксп
                    if (checkEdit23.Checked) { Note = textEdit17.Text.Trim(); } else { Note = ""; } //Причина негод

                    ArrayList rows = new ArrayList();
                    List<Object> aList = new List<Object>();
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
                        aList.Add(row["ID"]);
                        Arrays = string.Join(" ", aList);
                        string UpdateAutnAll = "exec dbo.UpdateAutnAll '" + Avail + "','" + Klapan + "','" + DemSkob + "','" + PTC + "','" + Ushki + "','" + Skobi + "','" + Shaiba + "','" + Lest + "','" + Greben + "','" + Barash + "','" + TriBolt + "','" + Exp + "','" + Note + "','" + Arrays + "'";
                        DbConnection.DBConnect(UpdateAutnAll);
                    }
                    Refreshh("1");
                }
                //Редактировать
                else
                {
                    string CheckOwner = "select c.Current_owner from d__Carriage c where c.CarNumber = " + textEdit1.Text.Trim();
                    DataTable dt = DbConnection.DBConnect(CheckOwner);
                    string Check = dt.Rows[0][0].ToString();
                    if (Check == "")
                    {
                        MessageBox.Show("У данного В/Ц отсутствует собственник!", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else
                    {
                        if (textEdit1.Text.Length > 7 && dt.Rows[0][0].ToString() == "PTC Operator TOO")
                        {
                            if ((textEdit3.Text != "" && textEdit4.Text != "" && textEdit6.Text != "" && textEdit7.Text != "" && textEdit8.Text != "" && textEdit11.Text != "" && textEdit25.Text != "") &
                                    (textEdit6.Text != "0" || textEdit7.Text != "0" || textEdit8.Text != "0" || textEdit11.Text != "0" || textEdit25.Text != ""))
                            {

                                string StartProcess, EndProcess, DateTehCart, StartTor, EndTor, ArrivalDate, DateVU19 = string.Empty;

                                if (checkEdit25.Checked) { StartProcess = dateTimePicker9.Value.ToString(); } else { StartProcess = "null"; }
                                if (checkEdit26.Checked) { EndProcess = dateTimePicker8.Value.ToString(); } else { EndProcess = "null"; }
                                if (checkEdit27.Checked) { DateTehCart = dateTimePicker6.Value.ToString(); } else { DateTehCart = "null"; }
                                if (checkEdit28.Checked) { StartTor = dateTimePicker2.Value.ToString(); } else { StartTor = "null"; }
                                if (checkEdit29.Checked) { EndTor = dateTimePicker3.Value.ToString(); } else { EndTor = "null"; }
                                if (checkEdit30.Checked) { ArrivalDate = dateTimePicker5.Value.ToString(); } else { ArrivalDate = "null"; }
                                if (checkEdit31.Checked) { DateVU19 = dateTimePicker4.Value.ToString(); } else { DateVU19 = "null"; }

                                string Update = "exec [dbo].[UpdateRenderedService] " +
                                    "'"+User_ID+"'," + textEdit1.Text.Trim() + "," +
                                    "'"+textEdit2.Text.Trim()+"'," +
                                    "" + textEdit4.Text.Trim() + "," +
                                    "" + textEdit6.Text.Trim() + "," +
                                    "" + textEdit8.Text.Trim() + "," +
                                    "" + textEdit7.Text.Trim() + "," +
                                    "" + textEdit25.Text.Trim() + "," +
                                    "" + textEdit11.Text.Trim() + "," +
                                    "" + comboBox1.SelectedValue.ToString() + "," +
                                    "'" + textEdit3.Text.Trim() + "'," +
                                    "" + comboBox2.SelectedValue.ToString() + "," +
                                    "'" + StartProcess + "', " +
                                    "'" + EndProcess + "', " +
                                    "'" + DateTehCart + "'," +
                                    "'" + StartTor + "', " +
                                    "'" + EndTor + "', " +
                                    "'" + ArrivalDate + "'," +
                                    "'" + DateVU19 + "'," +
                                    "" + SelectItemRow;
                                    DbConnection.DBConnect(Update);
                                    Refreshh("1");
                                    Block();
                                    TempUpdate = -1;
                                //}
                                //else
                                //{
                                //    MessageBox.Show("Заполните пустые поля в АУТН!", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                //}
                            }
                            else
                            {
                                MessageBox.Show("Заполните пустые поля в Обработке!", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
                        else
                        {
                            if ((textEdit3.Text != "" && textEdit4.Text != "" && textEdit6.Text != "" && textEdit7.Text != "" && textEdit8.Text != "" && textEdit11.Text != "" && textEdit25.Text != "") &
                                    (textEdit6.Text != "0" || textEdit7.Text != "0" || textEdit8.Text != "0" || textEdit11.Text != "0" || textEdit25.Text != ""))
                            {
                                string StartProcess, EndProcess, DateTehCart, StartTor, EndTor, ArrivalDate, DateVU19 = string.Empty;

                                if (checkEdit25.Checked) { StartProcess = dateTimePicker9.Value.ToString(); } else { StartProcess = "null"; }
                                if (checkEdit26.Checked) { EndProcess = dateTimePicker8.Value.ToString(); } else { EndProcess = "null"; }
                                if (checkEdit27.Checked) { DateTehCart = dateTimePicker6.Value.ToString(); } else { DateTehCart = "null"; }
                                if (checkEdit28.Checked) { StartTor = dateTimePicker2.Value.ToString(); } else { StartTor = "null"; }
                                if (checkEdit29.Checked) { EndTor = dateTimePicker3.Value.ToString(); } else { EndTor = "null"; }
                                if (checkEdit30.Checked) { ArrivalDate = dateTimePicker5.Value.ToString(); } else { ArrivalDate = "null"; }
                                if (checkEdit31.Checked) { DateVU19 = dateTimePicker4.Value.ToString(); } else { DateVU19 = "null"; }

                                string Update = "exec [dbo].[UpdateRenderedService] " +
                                    "'" + User_ID + "'," + textEdit1.Text.Trim() + "," +
                                    "'" + textEdit2.Text.Trim() + "'," +
                                    "" + textEdit4.Text.Trim() + "," +
                                    "" + textEdit6.Text.Trim() + "," +
                                    "" + textEdit8.Text.Trim() + "," +
                                    "" + textEdit7.Text.Trim() + "," +
                                    "" + textEdit25.Text.Trim() + "," +
                                    "" + textEdit11.Text.Trim() + "," +
                                    "" + comboBox1.SelectedValue.ToString() + "," +
                                    "'" + textEdit3.Text.Trim() + "'," +
                                    "" + comboBox2.SelectedValue.ToString() + "," +
                                    "'" + StartProcess + "', " +
                                    "'" + EndProcess + "', " +
                                    "'" + DateTehCart + "'," +
                                    "'" + StartTor + "', " +
                                    "'" + EndTor + "', " +
                                    "'" + ArrivalDate + "'," +
                                    "'" + DateVU19 + "'," +
                                    "" + SelectItemRow;
                                DbConnection.DBConnect(Update);
                                Refreshh("1");
                                Block();
                                TempUpdate = -1;
                            }
                            else
                            {
                                MessageBox.Show("Заполните пустые поля в Обработке!", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void gridView1_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            try
            {
                string Id = gridView1.GetFocusedDataRow()[0].ToString();
                SelectItemRow = Convert.ToInt32(Id);

                GetCurrent();
                simpleButton9.Enabled = false;
                TempUpdate = -1;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void GetCurrent()
        {
            try
            {
                if(GroupUpdate == 1)
                {
                    SelectBrigadeID = gridView1.GetFocusedDataRow()[1].ToString();
                    SelectProductID = gridView1.GetFocusedDataRow()[2].ToString();
                    comboBox1.DataBindings.Clear();
                    comboBox2.DataBindings.Clear();
                    Update(SelectItemRow);
                }
                else if(UpdateAutn == 1)
                {
                    SelectBrigadeID = gridView1.GetFocusedDataRow()[1].ToString();
                    SelectProductID = gridView1.GetFocusedDataRow()[2].ToString();
                    comboBox1.DataBindings.Clear();
                    comboBox2.DataBindings.Clear();
                    Update(SelectItemRow);
                }
                else
                {
                    SelectBrigadeID = gridView1.GetFocusedDataRow()[1].ToString();
                    SelectProductID = gridView1.GetFocusedDataRow()[2].ToString();
                    comboBox1.DataBindings.Clear();
                    comboBox2.DataBindings.Clear();
                    Block();
                    Update(SelectItemRow);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void Update(int Id)
        {
            try
            {
                string GetValue = "select * from d__RenderedService where ID = " + Id;
                DataTable dt = DbConnection.DBConnect(GetValue);

                bool HOL = Convert.ToBoolean(dt.Rows[0][3].ToString());
                int HOL1;
                if (HOL == true)
                {
                    HOL1 = 1;
                }
                else
                {
                    HOL1 = 0;
                }

                bool GOR = Convert.ToBoolean(dt.Rows[0][4].ToString());
                int GOR1;
                if (GOR == true)
                {
                    GOR1 = 1;
                }
                else
                {
                    GOR1 = 0;
                }

                bool TOR = Convert.ToBoolean(dt.Rows[0][5].ToString());
                int TOR1;
                if (TOR == true)
                {
                    TOR1 = 1;
                }
                else
                {
                    TOR1 = 0;
                }

                bool DR1 = Convert.ToBoolean(dt.Rows[0][6].ToString());
                int DR11;
                if (DR1 == true)
                {
                    DR11 = 1;
                }
                else
                {
                    DR11 = 0;
                }

                bool DR2 = Convert.ToBoolean(dt.Rows[0][7].ToString());
                int DR22;
                if (DR2 == true)
                {
                    DR22 = 1;
                }
                else
                {
                    DR22 = 0;
                }

                bool Trafar = Convert.ToBoolean(dt.Rows[0][8].ToString());
                int Trafar1;
                if (Trafar == true)
                {
                    Trafar1 = 1;
                }
                else
                {
                    Trafar1 = 0;
                }

                bool Naruzhka = Convert.ToBoolean(dt.Rows[0][9].ToString());
                int Naruzhka1;
                if (Naruzhka == true)
                {
                    Naruzhka1 = 1;
                }
                else
                {
                    Naruzhka1 = 0;
                }

                textEdit1.Text = dt.Rows[0][1].ToString();
                textEdit3.Text = dt.Rows[0][12].ToString();
                textEdit4.Text = dt.Rows[0][2].ToString();
                textEdit6.Text = HOL1.ToString();
                textEdit8.Text = GOR1.ToString();
                textEdit7.Text = TOR1.ToString();
                textEdit9.Text = DR11.ToString();
                textEdit10.Text = DR22.ToString();
                textEdit11.Text = Trafar1.ToString();
                textEdit5.Text = Naruzhka1.ToString();

                if (dt.Rows[0][28] == System.DBNull.Value) { dateTimePicker9.Value = System.DateTime.Now; } else { dateTimePicker9.Value = Convert.ToDateTime(dt.Rows[0][28]); }
                if (dt.Rows[0][29] == System.DBNull.Value) { dateTimePicker8.Value = System.DateTime.Now; } else { dateTimePicker8.Value = Convert.ToDateTime(dt.Rows[0][29]); }
                if (dt.Rows[0][30] == System.DBNull.Value) { dateTimePicker6.Value = System.DateTime.Now; } else { dateTimePicker6.Value = Convert.ToDateTime(dt.Rows[0][30]); }
                if (dt.Rows[0][31] == System.DBNull.Value) { dateTimePicker2.Value = System.DateTime.Now; } else { dateTimePicker2.Value = Convert.ToDateTime(dt.Rows[0][31]); }
                if (dt.Rows[0][32] == System.DBNull.Value) { dateTimePicker3.Value = System.DateTime.Now; } else { dateTimePicker3.Value = Convert.ToDateTime(dt.Rows[0][32]); }
                if (dt.Rows[0][33] == System.DBNull.Value) { dateTimePicker5.Value = System.DateTime.Now; } else { dateTimePicker5.Value = Convert.ToDateTime(dt.Rows[0][33]); }
                if (dt.Rows[0][34] == System.DBNull.Value) { dateTimePicker4.Value = System.DateTime.Now; } else { dateTimePicker4.Value = Convert.ToDateTime(dt.Rows[0][34]); }

                Num = Convert.ToInt32(dt.Rows[0][26].ToString());

                string GetAutn = "select * from d__Autn_2 where Head_ID = " + Id;
                DataTable Autn = DbConnection.DBConnect(GetAutn);

                textEdit18.Text = Autn.Rows[0][1].ToString();
                textEdit12.Text = Autn.Rows[0][2].ToString();
                textEdit13.Text = Autn.Rows[0][3].ToString();
                textEdit14.Text = Autn.Rows[0][4].ToString();
                textEdit24.Text = Autn.Rows[0][5].ToString();
                textEdit15.Text = Autn.Rows[0][6].ToString();
                textEdit16.Text = Autn.Rows[0][7].ToString();
                textEdit23.Text = Autn.Rows[0][8].ToString();
                textEdit22.Text = Autn.Rows[0][9].ToString();
                textEdit21.Text = Autn.Rows[0][10].ToString();
                textEdit20.Text = Autn.Rows[0][11].ToString();
                textEdit19.Text = Autn.Rows[0][12].ToString();
                textEdit17.Text = Autn.Rows[0][13].ToString();

                string Brigade = "Select * from d__Brigade";
                DataTable dt1 = DbConnection.DBConnect(Brigade);
                comboBox1.DataSource = dt1;
                comboBox1.DisplayMember = "Surname";
                comboBox1.ValueMember = "ID";
                comboBox1.DataBindings.Clear();
                comboBox1.DataBindings.Add("SelectedValue", this, "SelectBrigadeID", true, DataSourceUpdateMode.OnPropertyChanged);

                string Product = "Select * from d__Product";
                DataTable dt2 = DbConnection.DBConnect(Product);
                comboBox2.DataSource = dt2;
                comboBox2.DisplayMember = "Name";
                comboBox2.ValueMember = "ID";
                comboBox2.DataBindings.Add("SelectedValue", this, "SelectProductID", true, DataSourceUpdateMode.OnPropertyChanged);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Block()
        {
            textEdit1.Enabled = false;
            //textEdit3.Enabled = false;
            textEdit4.Enabled = false;
            textEdit6.Enabled = false;
            textEdit8.Enabled = false;
            textEdit7.Enabled = false;
            textEdit9.Enabled = false;
            textEdit10.Enabled = false;
            textEdit11.Enabled = false;
            textEdit5.Enabled = false;
            textEdit25.Enabled = false;
            comboBox1.Enabled = false;
            comboBox2.Enabled = false;
            simpleButton1.Enabled = false;

            textEdit12.Enabled = false;
            textEdit13.Enabled = false;
            textEdit14.Enabled = false;
            textEdit15.Enabled = false;
            textEdit16.Enabled = false;
            textEdit17.Enabled = false;
            textEdit18.Enabled = false;
            textEdit19.Enabled = false;
            textEdit20.Enabled = false;
            textEdit21.Enabled = false;
            textEdit22.Enabled = false;
            textEdit23.Enabled = false;
            textEdit24.Enabled = false;

            dateTimePicker2.Enabled = false;
            dateTimePicker3.Enabled = false;
            dateTimePicker4.Enabled = false;
            dateTimePicker5.Enabled = false;
            dateTimePicker6.Enabled = false;
            dateTimePicker8.Enabled = false;
            dateTimePicker9.Enabled = false;

            checkEdit25.Checked = false;
            checkEdit26.Checked = false;
            checkEdit27.Checked = false;
            checkEdit28.Checked = false;
            checkEdit29.Checked = false;
            checkEdit30.Checked = false;
            checkEdit31.Checked = false;
        }
        private void Unblock()
        {
            textEdit1.Enabled = true;
            //textEdit3.Enabled = true;
            textEdit4.Enabled = true;
            textEdit6.Enabled = true;
            textEdit8.Enabled = true;
            textEdit7.Enabled = true;
            textEdit9.Enabled = true;
            textEdit10.Enabled = true;
            textEdit11.Enabled = true;
            textEdit5.Enabled = true;
            textEdit25.Enabled = true;
            comboBox1.Enabled = true;
            comboBox2.Enabled = true;
            simpleButton1.Enabled = true;
        }

        private void simpleButton4_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("Удалить выделенную запись?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    ArrayList rows = new ArrayList();
                    List<Object> aList = new List<Object>();
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
                        aList.Add(row["ID"]);
                        Arrays = string.Join(" ", aList);
                        string delete = "exec dbo.DeleteRenderedService '" + Arrays + "'";
                        DbConnection.DBConnect(delete);
                    }
                    Refreshh("2");
                    if(gridView1.RowCount > 0)
                    {
                        gridView1_RowCellClick(null,null);
                    }
                    else
                    {
                        textEdit1.Text = "";
                    }
                }
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void simpleButton9_Click(object sender, EventArgs e)
        {
            try
            {
                string GetInvoice = "dbo.GetDocNum";
                DataTable dt = DbConnection.DBConnect(GetInvoice);
                textEdit3.Text = dt.Rows[0][0].ToString();
                Num = Convert.ToInt32(dt.Rows[0][1].ToString());
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void textEdit6_Properties_KeyPress(object sender, KeyPressEventArgs e)
        {
            var validKeys = new[] { Keys.Back, Keys.D0, Keys.D1 };

            e.Handled = !validKeys.Contains((Keys)e.KeyChar);
        }

        private void textEdit8_Properties_KeyPress(object sender, KeyPressEventArgs e)
        {
            var validKeys = new[] { Keys.Back, Keys.D0, Keys.D1 };

            e.Handled = !validKeys.Contains((Keys)e.KeyChar);
        }

        private void textEdit7_Properties_KeyPress(object sender, KeyPressEventArgs e)
        {
            var validKeys = new[] { Keys.Back, Keys.D0, Keys.D1 };

            e.Handled = !validKeys.Contains((Keys)e.KeyChar);
        }

        private void textEdit9_Properties_KeyPress(object sender, KeyPressEventArgs e)
        {
            var validKeys = new[] { Keys.Back, Keys.D0, Keys.D1 };

            e.Handled = !validKeys.Contains((Keys)e.KeyChar);
        }

        private void textEdit10_Properties_KeyPress(object sender, KeyPressEventArgs e)
        {
            var validKeys = new[] { Keys.Back, Keys.D0, Keys.D1 };

            e.Handled = !validKeys.Contains((Keys)e.KeyChar);
        }

        private void textEdit11_Properties_KeyPress(object sender, KeyPressEventArgs e)
        {
            var validKeys = new[] { Keys.Back, Keys.D0, Keys.D1 };

            e.Handled = !validKeys.Contains((Keys)e.KeyChar);
        }

        private void textEdit5_Properties_KeyPress(object sender, KeyPressEventArgs e)
        {
            var validKeys = new[] { Keys.Back, Keys.D0, Keys.D1 };

            e.Handled = !validKeys.Contains((Keys)e.KeyChar);
        }

        private void textEdit9_Properties_EditValueChanged(object sender, EventArgs e)
        {
            if(textEdit9.Text == "1")
            {
                textEdit7.Text = "0";
                textEdit9.Text = "1";
                textEdit8.Text = "1";
                textEdit6.Text = "0";
                textEdit10.Text = "0";
            }
            else if(textEdit9.Text == "0")
            {
                textEdit7.Text = "0";
                textEdit8.Text = "0";
                textEdit10.Text = "0";
            }
        }

        private void textEdit10_Properties_EditValueChanged(object sender, EventArgs e)
        {
            if (textEdit10.Text == "1")
            {
                textEdit7.Text = "0";
                textEdit10.Text = "1";
                textEdit8.Text = "1";
                textEdit6.Text = "0";
                textEdit25.Text = "0";
            }
            else if(textEdit10.Text == "0")
            {
                textEdit7.Text = "0";
                textEdit8.Text = "0";
                textEdit25.Text = "0";
            }
        }

        private void textEdit7_Properties_EditValueChanged(object sender, EventArgs e)
        {
            if(textEdit7.Text == "1")
            {
                textEdit6.Text = "0";
                textEdit8.Text = "1";
                textEdit7.Text = "1";
                textEdit25.Text = "0";
            }
            else if(textEdit7.Text == "0")
            {
                //textEdit8.Text = "0";
                textEdit25.Text = "0";
            }
        }
        private void textEdit25_EditValueChanged(object sender, EventArgs e)
        {
            if (textEdit25.Text == "1")
            {
                textEdit6.Text = "0";
                textEdit8.Text = "1";
                textEdit7.Text = "0";
            }
        }

        private void textEdit6_Properties_EditValueChanged(object sender, EventArgs e)
        {
            if(textEdit6.Text == "1")
            {
                textEdit7.Text = "0";
                textEdit8.Text = "0";
                textEdit25.Text = "0";
            }
        }

        private void textEdit8_Properties_EditValueChanged(object sender, EventArgs e)
        {
            if (textEdit8.Text == "1")
            {
                textEdit6.Text = "0";
            }
            else if(textEdit8.Text == "0")
            {
                textEdit7.Text = "0";
                textEdit25.Text = "0";
            }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            DateTime today = DateTime.Today;
            DateTime dateTime = Convert.ToDateTime(dateTimePicker1.Value.ToShortDateString());

            if(dateTime > today)  
            {
                dateTimePicker1.Value = DateTime.Today;

                Refreshh("2");
                if (gridView1.RowCount > 0)
                {
                    gridView1_RowCellClick(null, null);
                }
            }
            else
            {
                Refreshh("2");
                if (gridView1.RowCount > 0)
                {
                    gridView1_RowCellClick(null, null);
                }
            }
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            Unblock();
            textEdit1.Focus();
            textEdit1.Text = "";
            simpleButton9.Enabled = true;
            Temp = SelectItemRow;
            SelectItemRow = 0;

            checkEdit25.Visible = true;
            checkEdit26.Visible = true;
            checkEdit27.Visible = true;
            checkEdit28.Visible = true;
            checkEdit29.Visible = true;
            checkEdit30.Visible = true;
            checkEdit31.Visible = true;

            Clear_AUTN();
        }

        private void simpleButton3_Click(object sender, EventArgs e)
        {
            TempUpdate = 0;
            Unblock();
            string CarNum = textEdit1.Text.Trim();
            textEdit1.Text = "";
            textEdit1.Text = CarNum;


            checkEdit25.Visible = true;
            checkEdit26.Visible = true;
            checkEdit27.Visible = true;
            checkEdit28.Visible = true;
            checkEdit29.Visible = true;
            checkEdit30.Visible = true;
            checkEdit31.Visible = true;
        }

        private void simpleButton5_Click(object sender, EventArgs e)
        {
            ChangeCurrentOwner change = new ChangeCurrentOwner();
            change.Owner = this;
            change.ShowDialog();
        }

        public void ChangeOwner(int Id)
        {
            try
            {
                ArrayList rows = new ArrayList();
                List<Object> CarList = new List<Object>();
                List<Object> IdList = new List<Object>();
                string CarArrays = string.Empty;
                string IdArrays = string.Empty;

                Int32[] selectedRowHandles = gridView1.GetSelectedRows();
                for (int i = 0; i < selectedRowHandles.Length; i++)
                {
                    int selectedRowHandle = selectedRowHandles[i];
                    if (selectedRowHandle >= 0)
                        rows.Add(gridView1.GetDataRow(selectedRowHandle));
                }
                foreach (DataRow row in rows)
                {
                    CarList.Add(row["Номер вагона"]);
                    CarArrays = string.Join(" ", CarList);
                    IdList.Add(row["ID"]);
                    IdArrays = string.Join(" ", IdList);
                    string UpdateOwner = "exec dbo.UpdateCurrentOwner '"+ CarArrays + "','"+ IdArrays + "',"+Id;
                    DbConnection.DBConnect(UpdateOwner);
                }
                Refreshh("1");
                gridView1_RowCellClick(null,null);
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void simpleButton7_Click(object sender, EventArgs e)
        {
            try
            {
                if (textEdit1.Enabled == false)
                {
                    string query = "exec [dbo].[LastRenderedService_1] " + textEdit1.Text.Trim() + "," + SelectItemRow;
                    DataTable dt = DbConnection.DBConnect(query);
                    LastRenderedServiceForm last = new LastRenderedServiceForm(dt);
                    last.ShowDialog();
                }
                else if (SelectItemRow == 0)
                {
                    string query = "exec [dbo].[LastRenderedService] " + textEdit1.Text.Trim();
                    DataTable dt = DbConnection.DBConnect(query);
                    LastRenderedServiceForm last = new LastRenderedServiceForm(dt);
                    last.ShowDialog();
                }
                else if (SelectItemRow > 1 && textEdit1.Enabled == true)
                {
                    string query = "exec [dbo].[LastRenderedService_1] " + textEdit1.Text.Trim() + "," + SelectItemRow;
                    DataTable dt = DbConnection.DBConnect(query);
                    LastRenderedServiceForm last = new LastRenderedServiceForm(dt);
                    last.ShowDialog();
                }
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void simpleButton6_Click(object sender, EventArgs e)
        {
            Temp = SelectItemRow;
            GroupUpdate = 1;
            simpleButton6.Visible = false;
            simpleButton8.Visible = true;
            simpleButton1.Enabled = true;

            textEdit1.Enabled = false;
            textEdit4.Enabled = false;
            textEdit3.Enabled = false;
            comboBox1.Enabled = false;
            comboBox2.Enabled = false;

            textEdit6.Enabled = false;
            textEdit7.Enabled = false;
            textEdit8.Enabled = false;
            textEdit9.Enabled = false;
            textEdit10.Enabled = false;
            textEdit11.Enabled = false;
            textEdit5.Enabled = false;

            dateTimePicker1.Enabled = false;
            simpleButton2.Enabled = false;
            simpleButton3.Enabled = false;
            simpleButton4.Enabled = false;
            simpleButton5.Enabled = false;

            checkEdit1.Visible = true;
            checkEdit2.Visible = true;
            checkEdit11.Visible = true;
            checkEdit4.Visible = true;
            checkEdit5.Visible = true;
            checkEdit6.Visible = true;
            checkEdit9.Visible = true;
            checkEdit32.Visible = true;

            checkEdit25.Visible = true;
            checkEdit26.Visible = true;
            checkEdit27.Visible = true;
            checkEdit28.Visible = true;
            checkEdit29.Visible = true;
            checkEdit30.Visible = true;
            checkEdit31.Visible = true;

            textEdit1.Text = "";
            textEdit3.Text = "";
        }
        private void simpleButton8_Click(object sender, EventArgs e)
        {
            GroupUpdate = -1;
            Block();

            if (gridView1.RowCount > 0)
            {
                gridView1_RowCellClick(null,null);
            }

            simpleButton8.Visible = false;
            simpleButton6.Visible = true;
            simpleButton1.Enabled = false;

            dateTimePicker1.Enabled = true;
            simpleButton2.Enabled = true;
            simpleButton3.Enabled = true;
            simpleButton4.Enabled = true;
            simpleButton5.Enabled = true;

            checkEdit1.Visible = false;
            checkEdit2.Visible = false;
            checkEdit11.Visible = false;
            checkEdit4.Visible = false;
            checkEdit5.Visible = false;
            checkEdit6.Visible = false;
            checkEdit9.Visible = false;
            checkEdit32.Visible = false;

            checkEdit1.Checked = false;
            checkEdit2.Checked = false;
            checkEdit11.Checked = false;
            checkEdit4.Checked = false;
            checkEdit5.Checked = false;
            checkEdit6.Checked = false;
            checkEdit7.Checked = false;
            checkEdit8.Checked = false;
            checkEdit9.Checked = false;
            checkEdit10.Checked = false;

            checkEdit25.Checked = false;
            checkEdit26.Checked = false;
            checkEdit27.Checked = false;
            checkEdit28.Checked = false;
            checkEdit29.Checked = false;
            checkEdit30.Checked = false;
            checkEdit31.Checked = false;
            checkEdit32.Checked = false;

            checkEdit25.Visible = false;
            checkEdit26.Visible = false;
            checkEdit27.Visible = false;
            checkEdit28.Visible = false;
            checkEdit29.Visible = false;
            checkEdit30.Visible = false;
            checkEdit31.Visible = false;
        }

        private void checkEdit1_Properties_CheckStateChanged(object sender, EventArgs e)
        {
            textEdit4.Enabled = (checkEdit1.CheckState == CheckState.Checked);
        }
        private void checkEdit32_Properties_CheckStateChanged(object sender, EventArgs e)
        {
            textEdit25.Enabled = (checkEdit32.CheckState == CheckState.Checked);
        }

        private void checkEdit2_Properties_CheckStateChanged(object sender, EventArgs e)
        {
            comboBox1.Enabled = (checkEdit2.CheckState == CheckState.Checked);
        }

        private void checkEdit11_Properties_CheckStateChanged(object sender, EventArgs e)
        {
            comboBox2.Enabled = (checkEdit11.CheckState == CheckState.Checked);
        }

        private void checkEdit4_Properties_CheckStateChanged(object sender, EventArgs e)
        {
            textEdit6.Enabled = (checkEdit4.CheckState == CheckState.Checked);
        }

        private void checkEdit5_Properties_CheckStateChanged(object sender, EventArgs e)
        {
            textEdit8.Enabled = (checkEdit5.CheckState == CheckState.Checked);
        }

        private void checkEdit6_Properties_CheckStateChanged(object sender, EventArgs e)
        {
            textEdit7.Enabled = (checkEdit6.CheckState == CheckState.Checked);
        }

        private void checkEdit7_Properties_CheckStateChanged(object sender, EventArgs e)
        {
            textEdit9.Enabled = (checkEdit7.CheckState == CheckState.Checked);
        }

        private void checkEdit8_Properties_CheckStateChanged(object sender, EventArgs e)
        {
            textEdit10.Enabled = (checkEdit8.CheckState == CheckState.Checked);
        }

        private void checkEdit9_Properties_CheckStateChanged(object sender, EventArgs e)
        {
            textEdit11.Enabled = (checkEdit9.CheckState == CheckState.Checked);
        }

        private void checkEdit10_Properties_CheckStateChanged(object sender, EventArgs e)
        {
            textEdit5.Enabled = (checkEdit10.CheckState == CheckState.Checked);
        }

        //AUTN
        private void checkEdit3_Properties_CheckStateChanged(object sender, EventArgs e)
        {
            textEdit12.Enabled = (checkEdit3.CheckState == CheckState.Checked);
        }

        private void checkEdit17_Properties_CheckStateChanged(object sender, EventArgs e)
        {
            textEdit18.Enabled = (checkEdit17.CheckState == CheckState.Checked);
        }

        private void textEdit18_Properties_KeyPress(object sender, KeyPressEventArgs e)
        {
            var validKeys = new[] { Keys.Back, Keys.D0, Keys.D1 };

            e.Handled = !validKeys.Contains((Keys)e.KeyChar);
        }

        private void checkEdit12_Properties_CheckStateChanged(object sender, EventArgs e)
        {
            textEdit13.Enabled = (checkEdit12.CheckState == CheckState.Checked);
        }

        private void textEdit13_Properties_KeyPress(object sender, KeyPressEventArgs e)
        {
            var validKeys = new[] { Keys.Back, Keys.D0, Keys.D1 };

            e.Handled = !validKeys.Contains((Keys)e.KeyChar);
        }

        private void checkEdit13_Properties_CheckStateChanged(object sender, EventArgs e)
        {
            textEdit14.Enabled = (checkEdit13.CheckState == CheckState.Checked);
        }

        private void textEdit14_Properties_KeyPress(object sender, KeyPressEventArgs e)
        {
            var validKeys = new[] { Keys.Back, Keys.D0, Keys.D1 };

            e.Handled = !validKeys.Contains((Keys)e.KeyChar);
        }

        private void checkEdit14_Properties_CheckStateChanged(object sender, EventArgs e)
        {
            textEdit24.Enabled = (checkEdit14.CheckState == CheckState.Checked);
        }

        private void textEdit24_Properties_KeyPress(object sender, KeyPressEventArgs e)
        {
            var validKeys = new[] { Keys.Back, Keys.D0, Keys.D1 };

            e.Handled = !validKeys.Contains((Keys)e.KeyChar);
        }

        private void textEdit15_Properties_KeyPress(object sender, KeyPressEventArgs e)
        {
            var validKeys = new[] { Keys.Back, Keys.D0, Keys.D1 };

            e.Handled = !validKeys.Contains((Keys)e.KeyChar);
        }

        private void checkEdit15_Properties_CheckStateChanged(object sender, EventArgs e)
        {
            textEdit15.Enabled = (checkEdit15.CheckState == CheckState.Checked);
        }

        private void textEdit16_Properties_KeyPress(object sender, KeyPressEventArgs e)
        {
            var validKeys = new[] { Keys.Back, Keys.D0, Keys.D1 };

            e.Handled = !validKeys.Contains((Keys)e.KeyChar);
        }

        private void checkEdit22_Properties_CheckStateChanged(object sender, EventArgs e)
        {
            textEdit16.Enabled = (checkEdit22.CheckState == CheckState.Checked);
        }

        private void textEdit23_Properties_KeyPress(object sender, KeyPressEventArgs e)
        {
            var validKeys = new[] { Keys.Back, Keys.D0, Keys.D1 };

            e.Handled = !validKeys.Contains((Keys)e.KeyChar);
        }

        private void checkEdit21_Properties_CheckStateChanged(object sender, EventArgs e)
        {
            textEdit23.Enabled = (checkEdit21.CheckState == CheckState.Checked);
        }

        private void textEdit22_Properties_KeyPress(object sender, KeyPressEventArgs e)
        {
            var validKeys = new[] { Keys.Back, Keys.D0, Keys.D1 };

            e.Handled = !validKeys.Contains((Keys)e.KeyChar);
        }

        private void checkEdit20_Properties_CheckStateChanged(object sender, EventArgs e)
        {
            textEdit22.Enabled = (checkEdit20.CheckState == CheckState.Checked);
        }

        private void textEdit21_Properties_KeyPress(object sender, KeyPressEventArgs e)
        {
            var validKeys = new[] { Keys.Back, Keys.D0, Keys.D1 };

            e.Handled = !validKeys.Contains((Keys)e.KeyChar);
        }

        private void checkEdit19_Properties_CheckStateChanged(object sender, EventArgs e)
        {
            textEdit21.Enabled = (checkEdit19.CheckState == CheckState.Checked);
        }

        private void textEdit20_Properties_KeyPress(object sender, KeyPressEventArgs e)
        {
            var validKeys = new[] { Keys.Back, Keys.D0, Keys.D1 };

            e.Handled = !validKeys.Contains((Keys)e.KeyChar);
        }

        private void checkEdit18_Properties_CheckStateChanged(object sender, EventArgs e)
        {
            textEdit20.Enabled = (checkEdit18.CheckState == CheckState.Checked);
        }

        private void textEdit19_Properties_KeyPress(object sender, KeyPressEventArgs e)
        {
            var validKeys = new[] { Keys.Back, Keys.D0, Keys.D1 };

            e.Handled = !validKeys.Contains((Keys)e.KeyChar);
        }

        private void checkEdit16_Properties_CheckStateChanged(object sender, EventArgs e)
        {
            textEdit19.Enabled = (checkEdit16.CheckState == CheckState.Checked);
        }

        private void checkEdit23_Properties_CheckStateChanged(object sender, EventArgs e)
        {
            textEdit17.Enabled = (checkEdit23.CheckState == CheckState.Checked);
        }

        private void textEdit12_Properties_KeyPress(object sender, KeyPressEventArgs e)
        {
            var validKeys = new[] { Keys.Back, Keys.D0, Keys.D1 };

            e.Handled = !validKeys.Contains((Keys)e.KeyChar);
        }

        private void checkEdit24_CheckedChanged(object sender, EventArgs e)
        {
            if(checkEdit24.Checked)
            {
                gridControl1.DataSource = null;
                gridView1.Columns.Clear();

                string Refreshh = "exec [dbo].[GetRenderedService] '" + dateTimePicker1.Value.ToShortDateString() + "', " + "@Type = " + 1;
                DataTable dt = DbConnection.DBConnect(Refreshh);
                gridControl1.DataSource = dt;
                gridView1.BestFitColumns();
                gridView1.Columns[3].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
                gridView1.Columns[0].Visible = false;
                gridView1.Columns[1].Visible = false;
                gridView1.Columns[2].Visible = false;
            }
            else
            {
                Refreshh("2");
            }
        }

        private void gridView1_RowStyle(object sender, RowStyleEventArgs e)
        {
            GridView View = sender as GridView;

            if (View.Columns.Count == 32)
            {
                if (e.RowHandle >= 0)
                {
                    string category = View.GetRowCellDisplayText(e.RowHandle, View.Columns["Удалён"]);
                    if (category == "Отмечено")
                    {
                        e.Appearance.BackColor = Color.LightPink;
                        //e.HighPriority = true;
                    }
                }
            }
            else
            {
                e.Appearance.BackColor = Color.White;
                //e.HighPriority = true;
            }

            //if (View.IsRowSelected(e.RowHandle))
            //{
            //    e.Appearance.ForeColor = Color.DarkBlue;
            //    e.Appearance.BackColor = Color.LightBlue;
            //    //e.HighPriority = true;
            //}

            if(View.FocusedRowHandle == e.RowHandle)
            {
                e.Appearance.ForeColor = Color.DarkBlue;
                e.Appearance.BackColor = Color.LightBlue;
            }
        }

        private void textEdit1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down)
            {
                textEdit4.Focus();
            }
            else if (e.KeyCode == Keys.Up)
            {
                simpleButton1.Focus();
            }
        }

        private void textEdit4_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down)
            {
                textEdit6.Focus();
            }
            else if (e.KeyCode == Keys.Up)
            {
                textEdit1.Focus();
            }
        }

        private void textEdit6_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down)
            {
                textEdit7.Focus();
            }
            else if (e.KeyCode == Keys.Up)
            {
                textEdit4.Focus();
            }
        }

        private void textEdit7_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down)
            {
                textEdit8.Focus();
            }
            else if (e.KeyCode == Keys.Up)
            {
                textEdit6.Focus();
            }
        }

        private void textEdit8_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down)
            {
                textEdit25.Focus();
            }
            else if (e.KeyCode == Keys.Up)
            {
                textEdit7.Focus();
            }
        }
        private void textEdit25_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down)
            {
                textEdit11.Focus();
            }
            else if (e.KeyCode == Keys.Up)
            {
                textEdit8.Focus();
            }
        }

        private void textEdit9_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down)
            {
                textEdit10.Focus();
            }
            else if (e.KeyCode == Keys.Up)
            {
                textEdit8.Focus();
            }
        }

        private void textEdit10_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down)
            {
                textEdit11.Focus();
            }
            else if (e.KeyCode == Keys.Up)
            {
                textEdit9.Focus();
            }
        }

        private void textEdit11_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down)
            {
                textEdit5.Focus();
            }
            else if (e.KeyCode == Keys.Up)
            {
                textEdit10.Focus();
            }
        }

        private void textEdit5_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down)
            {
                if(textEdit18.Enabled == false)
                {
                    simpleButton1.Focus();
                }
                else
                {
                    textEdit18.Focus();
                }
            }
            else if (e.KeyCode == Keys.Up)
            {
                textEdit11.Focus();
            }
        }

        private void textEdit18_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down)
            {
                textEdit12.Focus();
            }
            else if (e.KeyCode == Keys.Up)
            {
                textEdit5.Focus();
            }
        }

        private void textEdit12_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down)
            {
                textEdit13.Focus();
            }
            else if (e.KeyCode == Keys.Up)
            {
                textEdit18.Focus();
            }
        }

        private void textEdit13_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down)
            {
                textEdit14.Focus();
            }
            else if (e.KeyCode == Keys.Up)
            {
                textEdit12.Focus();
            }
        }

        private void textEdit14_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down)
            {
                textEdit24.Focus();
            }
            else if (e.KeyCode == Keys.Up)
            {
                textEdit13.Focus();
            }
        }

        private void textEdit24_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down)
            {
                textEdit15.Focus();
            }
            else if (e.KeyCode == Keys.Up)
            {
                textEdit14.Focus();
            }
        }

        private void textEdit15_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down)
            {
                textEdit16.Focus();
            }
            else if (e.KeyCode == Keys.Up)
            {
                textEdit24.Focus();
            }
        }

        private void textEdit16_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down)
            {
                textEdit23.Focus();
            }
            else if (e.KeyCode == Keys.Up)
            {
                textEdit15.Focus();
            }
        }

        private void textEdit23_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down)
            {
                textEdit22.Focus();
            }
            else if (e.KeyCode == Keys.Up)
            {
                textEdit16.Focus();
            }
        }

        private void textEdit22_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down)
            {
                textEdit21.Focus();
            }
            else if (e.KeyCode == Keys.Up)
            {
                textEdit23.Focus();
            }
        }

        private void textEdit21_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down)
            {
                textEdit20.Focus();
            }
            else if (e.KeyCode == Keys.Up)
            {
                textEdit22.Focus();
            }
        }

        private void textEdit20_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down)
            {
                textEdit19.Focus();
            }
            else if (e.KeyCode == Keys.Up)
            {
                textEdit21.Focus();
            }
        }

        private void textEdit19_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down)
            {
                textEdit17.Focus();
            }
            else if (e.KeyCode == Keys.Up)
            {
                textEdit20.Focus();
            }
        }

        private void textEdit17_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down)
            {
                simpleButton1.Focus();
            }
            else if (e.KeyCode == Keys.Up)
            {
                textEdit19.Focus();
            }
        }

        private void simpleButton1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down)
            {
                textEdit1.Focus();
            }
            else if (e.KeyCode == Keys.Up)
            {
                if(textEdit17.Enabled == false)
                {
                    textEdit5.Focus();
                }
                else
                {
                    textEdit17.Focus();
                }
            }
        }

        private void simpleButton10_Click(object sender, EventArgs e)
        {
            UpdateAutn = 1;
            //textEdit12.Enabled = true;
            //textEdit13.Enabled = true;
            //textEdit14.Enabled = true;
            //textEdit15.Enabled = true;
            //textEdit16.Enabled = true;
            //textEdit17.Enabled = true;
            //textEdit18.Enabled = true;
            //textEdit19.Enabled = true;
            //textEdit20.Enabled = true;
            //textEdit21.Enabled = true;
            //textEdit22.Enabled = true;
            //textEdit23.Enabled = true;
            //textEdit24.Enabled = true;
            simpleButton1.Enabled = true;
            simpleButton10.Visible = false;
            simpleButton11.Visible = true;

            checkEdit3.Visible = true;
            checkEdit12.Visible = true;
            checkEdit13.Visible = true;
            checkEdit14.Visible = true;
            checkEdit15.Visible = true;
            checkEdit16.Visible = true;
            checkEdit17.Visible = true;
            checkEdit18.Visible = true;
            checkEdit19.Visible = true;
            checkEdit20.Visible = true;
            checkEdit21.Visible = true;
            checkEdit22.Visible = true;
            checkEdit23.Visible = true;
        }

        private void simpleButton11_Click(object sender, EventArgs e)
        {
            simpleButton11.Visible = false;
            simpleButton10.Visible = true;
            UpdateAutn = -1;
            checkEdit3.Visible = false;
            checkEdit12.Visible = false;
            checkEdit13.Visible = false;
            checkEdit14.Visible = false;
            checkEdit15.Visible = false;
            checkEdit16.Visible = false;
            checkEdit17.Visible = false;
            checkEdit18.Visible = false;
            checkEdit19.Visible = false;
            checkEdit20.Visible = false;
            checkEdit21.Visible = false;
            checkEdit22.Visible = false;
            checkEdit23.Visible = false;

            checkEdit3.Checked = false;
            checkEdit12.Checked = false;
            checkEdit13.Checked = false;
            checkEdit14.Checked = false;
            checkEdit15.Checked = false;
            checkEdit16.Checked = false;
            checkEdit17.Checked = false;
            checkEdit18.Checked = false;
            checkEdit19.Checked = false;
            checkEdit20.Checked = false;
            checkEdit21.Checked = false;
            checkEdit22.Checked = false;
            checkEdit23.Checked = false;

            Block();
            Clear_AUTN();
        }

        private void Clear_AUTN()
        {
            textEdit12.Text = string.Empty;
            textEdit13.Text = string.Empty;
            textEdit14.Text = string.Empty;
            textEdit15.Text = string.Empty;
            textEdit16.Text = string.Empty;
            textEdit17.Text = string.Empty;
            textEdit18.Text = string.Empty;
            textEdit19.Text = string.Empty;
            textEdit20.Text = string.Empty;
            textEdit21.Text = string.Empty;
            textEdit22.Text = string.Empty;
            textEdit23.Text = string.Empty;
            textEdit24.Text = string.Empty;
        }

        private void gridView1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Up)
            {
                gridView1_RowCellClick(null, null);
            }
            else
            {
                if (e.KeyCode == Keys.Down)
                {
                    gridView1_RowCellClick(null, null);
                }
            }
        }

        private void checkEdit25_Properties_CheckStateChanged(object sender, EventArgs e)
        {
            dateTimePicker9.Enabled = (checkEdit25.CheckState == CheckState.Checked);
        }

        private void checkEdit26_Properties_CheckStateChanged(object sender, EventArgs e)
        {
            dateTimePicker8.Enabled = (checkEdit26.CheckState == CheckState.Checked);
        }

        private void checkEdit27_Properties_CheckStateChanged(object sender, EventArgs e)
        {
            dateTimePicker6.Enabled = (checkEdit27.CheckState == CheckState.Checked);
        }

        private void checkEdit28_Properties_CheckStateChanged(object sender, EventArgs e)
        {
            dateTimePicker2.Enabled = (checkEdit28.CheckState == CheckState.Checked);
        }

        private void checkEdit29_Properties_CheckStateChanged(object sender, EventArgs e)
        {
            dateTimePicker3.Enabled = (checkEdit29.CheckState == CheckState.Checked);
        }

        private void checkEdit30_Properties_CheckStateChanged(object sender, EventArgs e)
        {
            dateTimePicker5.Enabled = (checkEdit30.CheckState == CheckState.Checked);
        }

        private void checkEdit31_Properties_CheckStateChanged(object sender, EventArgs e)
        {
            dateTimePicker4.Enabled = (checkEdit31.CheckState == CheckState.Checked);
        }

        

        private void Enter_Key_Down(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                simpleButton1.PerformClick();
            }
        }
    }
}
