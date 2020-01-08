﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraGrid;

namespace Учет_цистерн.Forms.заявки_на_обработку
{
    public partial class OrderUpdateForm : Form
    {
        TradeWright.UI.Forms.TabControlExtra TabControlExtra;
        public OrderUpdateForm(TradeWright.UI.Forms.TabControlExtra tabControlExtra)
        {
            InitializeComponent();
            TabControlExtra = tabControlExtra;
            this.TabControlExtra.TabClosing += new System.EventHandler<System.Windows.Forms.TabControlCancelEventArgs>(this.tabControl1_TabClosing);
        }
        int SelectItemRow;
        public string GetStatus { get; set; }
        public string GetDate { get; set; }
        public int SelectedID { get; set; }
        public string SelectContrID { get; set; }
        public string SelectBrigadeID { get; set; }
        public string SelectStationID { get; set; }
        //Кнопка добавить
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                string GetID = "select ID from d__RenderedServiceHead where NUM = '" + GetStatus + "'";
                DataTable dt = DbConnection.DBConnect(GetID);
                string ID = dt.Rows[0][0].ToString();
                string RenrederServiceBody = "exec dbo.FillRenderedServiceBody NULL, NULL, NULL, " + ID;
                DbConnection.DBConnect(RenrederServiceBody);
                UpdateBody();
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        //Вставить вагоны из буфера обмена
        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                string GetID = "select ID from d__RenderedServiceHead where NUM = '" + GetStatus +"'";
                DataTable dt = DbConnection.DBConnect(GetID);
                string Insert = "exec dbo.InsertCarnumber '" + Clipboard.GetText() + "'," + dt.Rows[0][0].ToString();
                DbConnection.DBConnect(Insert);
                gridControl1.DataSource = null;
                GetData();
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        //удалить строку из таблицы
        private void button5_Click(object sender, EventArgs e)
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
                        aList.Add(row["BodyID"]);
                        Arrays = string.Join(" ", aList);
                        string delete = "exec dbo.DeleteRenderedBody '" + Arrays + "'";
                        DbConnection.DBConnect(delete);
                    }
                    UpdateBody();
                }
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        //скопировать строку
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                string CopyBody = "exec dbo.CopyRenderedServiceBody " + SelectItemRow;
                DbConnection.DBConnect(CopyBody);
                UpdateBody();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        //Обновить кнопка
        private void button6_Click(object sender, EventArgs e)
        {
            UpdateBody();
        }
        //Сохранить и закрыть документ
        private void button1_Click(object sender, EventArgs e)
        {
            SaveDoc();
        }
        private void SaveDoc()
        {
            try
            {
                string GetID = "select ID from d__RenderedServiceHead where NUM = '" + GetStatus +"'";
                DataTable dt = DbConnection.DBConnect(GetID);

                string CheckProduct = "select Carnumber from d__RenderedServiceBody b left join temp t on t.body_id = b.ID where b.Product_ID is NULL and (t.drkr = 0 and t.dr1 = 0 and t.ChangeTrafar = 0 and t.NaruzhChistka = 0) and b.Head_ID = " + dt.Rows[0][0].ToString(); 
                DataTable dt1 = DbConnection.DBConnect(CheckProduct);

                string CheckCarNumber = "select * from d__RenderedServiceBody where CarNumber is NULL and Head_ID = " + dt.Rows[0][0].ToString();
                DataTable dt2 = DbConnection.DBConnect(CheckCarNumber);

                string CheckRows = "select * from d__RenderedServiceBody where Head_ID = " + dt.Rows[0][0].ToString();
                DataTable dt3 = DbConnection.DBConnect(CheckRows);
                if (dt3.Rows.Count > 0)
                {
                    if (dt2.Rows.Count > 0)
                    {
                        MessageBox.Show("Введите вагон или удалите пустую строку!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else
                    {
                        if (dt1.Rows.Count > 0)
                        {
                            MessageBox.Show("У вагона " + dt1.Rows[0][0].ToString() + " не выбран продукт!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        else
                        {
                            string RenrederServiceHead = "exec dbo.RenderedServiceCreate '" + GetStatus + "','" + textBox1.Text.Trim() + "','" + dateTimePicker1.Value.Date.ToString() + "','" + textBox2.Text.Trim() + "','" + comboBox3.SelectedValue.ToString() + "','" + comboBox2.SelectedValue.ToString() + "',1,1";
                            DbConnection.DBConnect(RenrederServiceHead);
                            UpdateBody();
                            TabControlExtra.TabPages.Remove(TabControlExtra.SelectedTab);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("В табличной части отсутствуют вагоны!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void UpdateBody()
        {
            try
            {
                gridControl1.DataSource = null;
                GetData();
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OrderUpdateForm_Load(object sender, EventArgs e)
        {
            try
            {
                GetData();
                FillCombobox();
                string GetData1 = "select h.Invoice, ServiceDate, Operation, Invoice, s.Name,u.FIO from d__RenderedServiceHead h left join d__DocState s on s.ID = h.ID_DocState left join dbo.Users u on u.AID = h.ID_USER_INS where h.ID = " + SelectedID;
                DataTable dt = DbConnection.DBConnect(GetData1);
                this.textBox1.Text = dt.Rows[0][0].ToString();
                this.dateTimePicker1.Text = dt.Rows[0][1].ToString();
                this.textBox2.Text = dt.Rows[0][2].ToString();
                this.textBox3.Text = dt.Rows[0][3].ToString();
                this.label9.Text = dt.Rows[0][4].ToString();
                this.label10.Text = dt.Rows[0][5].ToString();
                textEdit1.Visible = false;
                memoEdit1.Visible = false;
                simpleButton1.Visible = false;

                GridColumnSummaryItem Carnumber = new GridColumnSummaryItem(DevExpress.Data.SummaryItemType.Count, "№ вагона", "{0}");
                GridColumnSummaryItem Cost = new GridColumnSummaryItem(DevExpress.Data.SummaryItemType.Sum, "Цена", "{0}");
                gridView1.Columns["№ вагона"].Summary.Add(Carnumber);
                gridView1.Columns["Цена"].Summary.Add(Cost);

                string GetDocState = "select ID_DocState from d__RenderedServiceHead where ID = " + SelectedID;
                DataTable DocStateDt = DbConnection.DBConnect(GetDocState);
                int DocState = Convert.ToInt32(DocStateDt.Rows[0][0]);
                if (DocState > 0 && DocState < 2)
                {
                    button1.Enabled = true;
                    button2.Enabled = true;
                    textBox1.Enabled = true;
                    textBox2.Enabled = true;
                    textBox3.Enabled = true;
                    dateTimePicker1.Enabled = true;
                    comboBox1.Enabled = true;
                    comboBox2.Enabled = true;
                    comboBox3.Enabled = true;
                    gridControl1.Enabled = true;
                    label12.Visible = false;
                    button3.Enabled = true;
                    button4.Enabled = true;
                    button5.Enabled = true;
                    button6.Enabled = true;
                }
                else if (DocState > 1 && DocState < 3)
                {
                    button1.Enabled = false;
                    button2.Enabled = false;
                    textBox1.Enabled = false;
                    textBox2.Enabled = false;
                    textBox3.Enabled = false;
                    dateTimePicker1.Enabled = false;
                    comboBox1.Enabled = false;
                    comboBox2.Enabled = false;
                    comboBox3.Enabled = false;
                    gridControl1.Enabled = false;
                    label12.Visible = true;
                    button3.Enabled = false;
                    button4.Enabled = false;
                    button5.Enabled = false;
                    button6.Enabled = false;
                }
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        //Получаю данные из таблицы d__RenderedServiceBody, также комбобоксы по внешним ключам - продукты и услуги
        private void GetData()
        {
            string ID = "select ID from d__RenderedServiceHead where NUM = '" + GetStatus +"'";
            DataTable GetIDDT = DbConnection.DBConnect(ID);
            string GetDocument = "exec dbo.GetRenderedServiceBody_v1 " + GetIDDT.Rows[0][0].ToString();
            DataTable GetDocumentDT = DbConnection.DBConnect(GetDocument);
            gridControl1.DataSource = GetDocumentDT;
            gridView1.Columns[0].Visible = false;
            //Продукты
            string GetProduct = "select d.ID as [ProductID], d.Name as [Продукт],qh.Name as [Обработка] from d__Product d left join qHangling qh on qh.ID = d.Handling_id";
            DataTable GetProductDT = DbConnection.DBConnect(GetProduct);
            //DataGridViewComboBoxColumn ProductComboBox = new DataGridViewComboBoxColumn();
            //ProductComboBox.DataSource = GetProductDT;
            //ProductComboBox.FlatStyle = FlatStyle.Flat;
            //ProductComboBox.HeaderText = "Продукт";
            //ProductComboBox.Name = "ProductComboBox";
            //ProductComboBox.DisplayMember = "Name";
            //ProductComboBox.ValueMember = "ID";
            //ProductComboBox.DataPropertyName = "Product_ID";
            //gridView1.Columns.Insert(3,ProductComboBox);

            RepositoryItemLookUpEdit riLookup = new RepositoryItemLookUpEdit();
            riLookup.DataSource = GetProductDT;
            riLookup.ValueMember = "ProductID";
            riLookup.DisplayMember = "Продукт";
            riLookup.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
            riLookup.SearchMode = DevExpress.XtraEditors.Controls.SearchMode.AutoComplete;
            riLookup.AutoSearchColumnIndex = 1;
            gridView1.Columns["Продукт"].ColumnEdit = riLookup;
            gridView1.BestFitColumns();
        }
        private void FillCombobox()
        {
            //Контрагенты
            string Owner = "select * from d__Owner";
            DataTable OwnerDT = DbConnection.DBConnect(Owner);
            comboBox1.DataSource = OwnerDT;
            comboBox1.DisplayMember = "Name";
            comboBox1.ValueMember = "ID";
            comboBox1.DataBindings.Add("SelectedValue", this, "SelectContrID", true, DataSourceUpdateMode.OnPropertyChanged);
            //Бригадиры
            string Brigade = "select * from d__Brigade";
            DataTable BrigadeDT = DbConnection.DBConnect(Brigade);
            comboBox2.DataSource = BrigadeDT;
            comboBox2.DisplayMember = "Name";
            comboBox2.ValueMember = "ID";
            comboBox2.DataBindings.Add("SelectedValue", this, "SelectBrigadeID", true, DataSourceUpdateMode.OnPropertyChanged);
            //Станции
            string Station = "select * from d__Station";
            DataTable StationDT = DbConnection.DBConnect(Station);
            comboBox3.DataSource = StationDT;
            comboBox3.DisplayMember = "Name";
            comboBox3.ValueMember = "ID";
            comboBox3.DataBindings.Add("SelectedValue", this, "SelectStationID", true, DataSourceUpdateMode.OnPropertyChanged);
        }

        private void tabControl1_TabClosing(object sender, TabControlCancelEventArgs e)
        {
            string GetID = "select ID from d__RenderedServiceHead where NUM = '" + GetStatus+"'";
            DataTable dt = DbConnection.DBConnect(GetID);

            if (e.TabPage.Text == "Акт № " + GetStatus + ". Редактирование от " + GetDate)
            {
                string CheckState = "select ID_DOCSTATE from d__RenderedServiceHead where ID = "+dt.Rows[0][0].ToString();
                DataTable CheckStateDt = DbConnection.DBConnect(CheckState);
                int State = Convert.ToInt32(CheckStateDt.Rows[0][0].ToString());
                if (State == 1)
                {
                    DialogResult result = MessageBox.Show("Документ изменён. Нажмите Да, если вы хотите сохранить изменения и закрыть документ, Нет - для возврата в документ.", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                    {
                        try
                        {
                            string CheckProduct = "select Carnumber from d__RenderedServiceBody b left join temp t on t.body_id = b.ID where b.Product_ID is NULL and (t.drkr = 0 and t.dr1 = 0 and t.ChangeTrafar = 0 and t.NaruzhChistka = 0) and b.Head_ID = " + dt.Rows[0][0].ToString();
                            DataTable dt1 = DbConnection.DBConnect(CheckProduct);

                            string CheckCarNumber = "select * from d__RenderedServiceBody where CarNumber is NULL and Head_ID = " + dt.Rows[0][0].ToString();
                            DataTable dt2 = DbConnection.DBConnect(CheckCarNumber);

                            string CheckRows = "select * from d__RenderedServiceBody where Head_ID = " + dt.Rows[0][0].ToString();
                            DataTable dt3 = DbConnection.DBConnect(CheckRows);
                            if (dt3.Rows.Count > 0)
                            {
                                if (dt2.Rows.Count > 0)
                                {
                                    MessageBox.Show("Введите вагон или удалите пустую строку!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    e.Cancel = true;
                                    return;
                                }
                                else
                                {
                                    if (dt1.Rows.Count > 0)
                                    {
                                        MessageBox.Show("У вагона " + dt1.Rows[0][0].ToString() + " не выбран продукт!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                        e.Cancel = true;
                                        return;
                                    }
                                    else
                                    {
                                        string RenrederServiceHead = "exec dbo.RenderedServiceCreate '" + GetStatus + "','" + textBox1.Text.Trim() + "','" + dateTimePicker1.Value.Date.ToString() + "','" + textBox2.Text.Trim() + "','" + comboBox3.SelectedValue.ToString() + "','" + comboBox2.SelectedValue.ToString() + "',1,1";
                                        DbConnection.DBConnect(RenrederServiceHead);
                                        UpdateBody();
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("В табличной части отсутствуют вагоны!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                e.Cancel = true;
                                return;
                            }
                        }
                        catch (SqlException ex)
                        {
                            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            e.Cancel = true;
                            return;
                        }
                        catch (Exception exp)
                        {
                            MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            e.Cancel = true;
                            return;
                        }
                    }
                    else
                    {
                        e.Cancel = true;
                        return;
                    }
                }
                else
                {

                }
            }
            else
            {

            }
        }

        private void gridView1_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            try
            {
                string Id = gridView1.GetFocusedDataRow()[0].ToString();
                SelectItemRow = Convert.ToInt32(Id);

                string CarNumber = gridView1.GetFocusedDataRow()[1].ToString();
                if (gridView1.DataRowCount == 0)
                {
                    textEdit1.Visible = false;
                    memoEdit1.Visible = false;
                    simpleButton1.Visible = false;
                }
                else
                {
                    if (CarNumber != string.Empty)
                    {
                        string query = "select [dbo].[CheckCarService] (" + CarNumber + "," + Id + ")";
                        DataTable dt = DbConnection.DBConnect(query);
                        int State = Convert.ToInt32(dt.Rows[0][0].ToString());
                        if (State == 1)
                        {
                            textEdit1.Visible = true;
                            simpleButton1.Visible = true;
                            string Message = "В/Ц " + CarNumber + " проходил обработку в течении последних 14 дней.";
                            textEdit1.Text = Message;
                        }
                        else
                        {
                            textEdit1.Visible = false;
                            simpleButton1.Visible = false;
                        }

                        string LastRent = "exec dbo.LastRent " + CarNumber;
                        DataTable dt1 = DbConnection.DBConnect(LastRent);
                        if (dt1.Rows.Count > 0)
                        {
                            memoEdit1.Visible = true;
                            memoEdit1.Text = "Последняя заявка: " + dt1.Rows[0][1] + " от " + dt1.Rows[0][2] + "" + "\r\n" + "Продукт: " + dt1.Rows[0][5] + "" + "\r\n" + "Была передача: " + dt1.Rows[0][3];
                        }
                        else
                        {
                            memoEdit1.Visible = false;
                        }
                    }
                    else
                    {
                        textEdit1.Visible = false;
                        memoEdit1.Visible = false;
                        simpleButton1.Visible = false;
                    }
                }
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void simpleButton1_Click(object sender, EventArgs e)
        {
            string CarNumber = gridView1.GetFocusedDataRow()[1].ToString();
            string query = "exec [dbo].[LastRenderedService] " + CarNumber + ", " + SelectItemRow;
            DataTable dt = DbConnection.DBConnect(query);
            LastRenderedServiceForm last = new LastRenderedServiceForm(dt);
            last.ShowDialog();
        }
        private void gridView1_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            try
            {
                //Проверка для вагона
                if (gridView1.GetFocusedDataRow()[1].ToString() == string.Empty)
                {
                    string UpdateBody = "exec dbo.UpdateRenderedServiceBody NULL, 1," + SelectItemRow;
                    DbConnection.DBConnect(UpdateBody);
                }
                else
                {
                    string CarNumber = gridView1.GetFocusedDataRow()[1].ToString();
                    string ServiceDate = dateTimePicker1.Value.Date.ToShortDateString();

                    string CheckCarnumber = "select * from d__Carriage where CarNumber = " + CarNumber;
                    DataTable dt1 = DbConnection.DBConnect(CheckCarnumber);
                    if (dt1.Rows.Count > 0)
                    {
                        string GetCarnumber = "select * from d__RenderedServiceHead h left join d__RenderedServiceBody b on h.ID = b.Head_ID where b.CarNumber = " + CarNumber + " and h.ServiceDate = '" + ServiceDate + "' and h.NUM != '" + GetStatus + "'";
                        DataTable dt = DbConnection.DBConnect(GetCarnumber);
                        if (dt.Rows.Count == 0)
                        {
                            string UpdateBody = "exec dbo.UpdateRenderedServiceBody '" + gridView1.GetFocusedDataRow()[1].ToString() + "',1," + SelectItemRow;
                            DbConnection.DBConnect(UpdateBody);
                        }
                        else
                        {
                            MessageBox.Show("Вагон " + CarNumber + " уже имеется в заявках на эту дату " + ServiceDate, "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            UpdateBody();
                        }
                    }
                    else
                    {
                        MessageBox.Show(CarNumber + " отсутствует в справочнике Вагоны!", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        UpdateBody();
                    }
                }
                //Проверка для операции
                if (gridView1.GetFocusedDataRow()[2].ToString() == string.Empty)
                {
                    string UpdateBody = "exec dbo.UpdateRenderedServiceBody NULL, 3, " + SelectItemRow;
                    DbConnection.DBConnect(UpdateBody);
                }
                else
                {
                    string UpdateBody = "exec dbo.UpdateRenderedServiceBody '" + gridView1.GetFocusedDataRow()[2].ToString() + "',3," + SelectItemRow;
                    DbConnection.DBConnect(UpdateBody);
                }
                //Проверка для продукта
                if (gridView1.GetFocusedDataRow()[3].ToString() == string.Empty)
                {
                    string UpdateBody = "exec dbo.UpdateRenderedServiceBody NULL, 2, " + SelectItemRow;
                    DbConnection.DBConnect(UpdateBody);
                }
                else
                {
                    string UpdateBody = "exec dbo.UpdateRenderedServiceBody '" + gridView1.GetFocusedDataRow()[3].ToString() + "',2," + SelectItemRow;
                    DbConnection.DBConnect(UpdateBody);
                }

                //axis
                string UpdateAxis = "update temp set axis = " + gridView1.GetFocusedDataRow()[4].ToString() + " where body_id = " + SelectItemRow;
                DbConnection.DBConnect(UpdateAxis);
                //gor
                string UpdateGor = "exec [dbo].[UpdateBodyTemp] '" + gridView1.GetFocusedDataRow()[5].ToString() + "',2," + SelectItemRow;
                DbConnection.DBConnect(UpdateGor);
                //hol
                string UpdateHol = "exec [dbo].[UpdateBodyTemp] '" + gridView1.GetFocusedDataRow()[6].ToString() + "',3," + SelectItemRow;
                DbConnection.DBConnect(UpdateHol);
                //tor
                string UpdateTor = "exec [dbo].[UpdateBodyTemp] '" + gridView1.GetFocusedDataRow()[7].ToString() + "',4," + SelectItemRow;
                DbConnection.DBConnect(UpdateTor);
                //drkr 1
                string UpdateDrkr = "exec [dbo].[UpdateBodyTemp] '" + gridView1.GetFocusedDataRow()[8].ToString() + "',5," + SelectItemRow;
                DbConnection.DBConnect(UpdateDrkr);
                //drkr 2
                string UpdateDrkr2 = "exec [dbo].[UpdateBodyTemp] '" + gridView1.GetFocusedDataRow()[9].ToString() + "',17," + SelectItemRow;
                DbConnection.DBConnect(UpdateDrkr2);
                //changeTrafar
                string changeTrafar = "exec [dbo].[UpdateBodyTemp] '" + gridView1.GetFocusedDataRow()[10].ToString() + "',18," + SelectItemRow;
                DbConnection.DBConnect(changeTrafar);
                //NaruzhOchistka
                string NaruzhOchistka = "exec [dbo].[UpdateBodyTemp] '" + gridView1.GetFocusedDataRow()[11].ToString() + "',19," + SelectItemRow;
                DbConnection.DBConnect(NaruzhOchistka);
                //AvailableAUTN
                string UpdateAvailableAUTN = "exec [dbo].[UpdateBodyTemp] '" + gridView1.GetFocusedDataRow()[14].ToString() + "',6," + SelectItemRow;
                DbConnection.DBConnect(UpdateAvailableAUTN);
                //Klapan
                string UpdateKlapan = "exec [dbo].[UpdateBodyAutn] '" + gridView1.GetFocusedDataRow()[15].ToString() + "',1," + SelectItemRow;
                DbConnection.DBConnect(UpdateKlapan);
                //DemSkobi
                string UpdateDemSkobi = "exec [dbo].[UpdateBodyTemp] '" + gridView1.GetFocusedDataRow()[16].ToString() + "',7," + SelectItemRow;
                DbConnection.DBConnect(UpdateDemSkobi);
                //Trafaret_PTC_Holding
                string UpdateTrafaret_PTC_Holding = "exec [dbo].[UpdateBodyTemp] '" + gridView1.GetFocusedDataRow()[17].ToString() + "',8," + SelectItemRow;
                DbConnection.DBConnect(UpdateTrafaret_PTC_Holding);
                //UshkiZavareni
                string UpdateUshkiZavareni = "exec [dbo].[UpdateBodyTemp] '" + gridView1.GetFocusedDataRow()[18].ToString() + "',9," + SelectItemRow;
                DbConnection.DBConnect(UpdateUshkiZavareni);
                //SkobiZavareni
                string UpdateSkobiZavareni = "exec [dbo].[UpdateBodyTemp] '" + gridView1.GetFocusedDataRow()[19].ToString() + "',10," + SelectItemRow;
                DbConnection.DBConnect(UpdateSkobiZavareni);
                //ShaibaValikZavareni
                string UpdateShaibaValikZavareni = "exec [dbo].[UpdateBodyTemp] '" + gridView1.GetFocusedDataRow()[20].ToString() + "',11," + SelectItemRow;
                DbConnection.DBConnect(UpdateShaibaValikZavareni);
                //VnutrLestnica
                string UpdateVnutrLestnica = "exec [dbo].[UpdateBodyTemp] '" + gridView1.GetFocusedDataRow()[21].ToString() + "',12," + SelectItemRow;
                DbConnection.DBConnect(UpdateVnutrLestnica);
                //Greben
                string UpdateGreben = "exec [dbo].[UpdateBodyTemp] '" + gridView1.GetFocusedDataRow()[22].ToString() + "',13," + SelectItemRow;
                DbConnection.DBConnect(UpdateGreben);
                //BarashkTip
                string UpdateBarashkTip = "exec [dbo].[UpdateBodyTemp] '" + gridView1.GetFocusedDataRow()[23].ToString() + "',14," + SelectItemRow;
                DbConnection.DBConnect(UpdateBarashkTip);
                //AvailTriBoltov
                string UpdateAvailTriBoltov = "exec [dbo].[UpdateBodyTemp] '" + gridView1.GetFocusedDataRow()[24].ToString() + "',15," + SelectItemRow;
                DbConnection.DBConnect(UpdateAvailTriBoltov);
                //AvailToExp
                string UpdateAvailToExp = "exec [dbo].[UpdateBodyTemp] '" + gridView1.GetFocusedDataRow()[25].ToString() + "',16," + SelectItemRow;
                DbConnection.DBConnect(UpdateAvailToExp);
                //Note
                string UpdateNote = "exec [dbo].[UpdateBodyAutn] '" + gridView1.GetFocusedDataRow()[26].ToString() + "',2," + SelectItemRow;
                DbConnection.DBConnect(UpdateNote);

                //for (int i = 0; i < dataGridView1.Rows.Count; i++)
                //{
                //    string FillBody = "exec dbo.UpdateRenderedServiceBody_Filter " + dataGridView1.Rows[i].Cells[3].Value + "," + dataGridView1.Rows[i].Cells[4].Value + "," + dataGridView1.Rows[i].Cells[5].Value + "," + dataGridView1.Rows[i].Cells[6].Value + "," + dataGridView1.Rows[i].Cells[7].Value + "," + SelectItemRow;
                //    DbConnection.DBConnect(FillBody);
                //}

                UpdateBody();
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
