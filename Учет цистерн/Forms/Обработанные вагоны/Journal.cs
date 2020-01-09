﻿using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraGrid;

namespace Учет_цистерн.Forms.Обработанные_вагоны
{
    public partial class Journal : Form
    {
        int SelectItemRow;
        public string SelectBrigadeID { get; set; }
        public string SelectProductID { get; set; }

        public Journal()
        {
            InitializeComponent();
        }

        private void Refresh()
        {
            gridControl1.DataSource = null;
            gridView1.Columns.Clear();
            string refresh = "exec [dbo].[GetRenderedService] '"+dateTimePicker1.Value.ToShortDateString()+"'";
            DataTable dt = DbConnection.DBConnect(refresh);
            gridControl1.DataSource = dt;
            gridView1.Columns[0].Visible = false;
            gridView1.Columns[1].Visible = false;
            gridView1.Columns[2].Visible = false;
            gridView1.Columns[3].Width = 150;
            gridView1.Columns[4].Width = 150;
            gridView1.Columns[9].Width = 150;
            gridView1.Columns[10].Width = 150;
            gridView1.Columns[11].Width = 150;
            gridView1.Columns[12].Width = 150;
            gridView1.Columns[13].Width = 150;
            gridView1.Columns[14].Width = 150;
            gridView1.Columns[16].Width = 350;
            gridView1.Columns[17].Width = 150;
            gridView1.Columns[18].Width = 150;

            textEdit1.Text = "";
            textEdit2.Text = "";
            textEdit3.Text = "";
            textEdit4.Text = "";
            textEdit5.Text = "0";
            textEdit6.Text = "0";
            textEdit7.Text = "0";
            textEdit8.Text = "0";
            textEdit9.Text = "0";
            textEdit10.Text = "0";
            textEdit11.Text = "0";
        }

        private void Fillcombobox()
        {
            string Brigade = "Select * from d__Brigade";
            DataTable dt = DbConnection.DBConnect(Brigade);
            comboBox1.DataSource = dt;
            comboBox1.DisplayMember = "Name";
            comboBox1.ValueMember = "ID";

            string Product = "Select * from d__Product";
            DataTable dt1 = DbConnection.DBConnect(Product);
            comboBox2.DataSource = dt1;
            comboBox2.DisplayMember = "Name";
            comboBox2.ValueMember = "ID";
        }

        private void Journal_Load(object sender, EventArgs e)
        {
            DateTime now = DateTime.Now;
            dateTimePicker1.Value = now;

            Refresh();
            Fillcombobox();
            Block();

            GridColumnSummaryItem Carnumber = new GridColumnSummaryItem(DevExpress.Data.SummaryItemType.Count, "Номер вагона", "{0}");
            GridColumnSummaryItem ServiceCost = new GridColumnSummaryItem(DevExpress.Data.SummaryItemType.Sum, "Сумма услуг", "{0}");
            GridColumnSummaryItem TorCost = new GridColumnSummaryItem(DevExpress.Data.SummaryItemType.Sum, "Сумма ТОР", "{0}");
            gridView1.Columns["Номер вагона"].Summary.Add(Carnumber);
            gridView1.Columns["Сумма услуг"].Summary.Add(ServiceCost);
            gridView1.Columns["Сумма ТОР"].Summary.Add(TorCost);
        }

        private void textEdit1_Properties_EditValueChanged(object sender, EventArgs e)
        {
            try
            {
                if (textEdit1.Text != string.Empty)
                {
                    string CheckOwner = "select o.Name from d__Carriage c left join d__Owner o on o.ID = c.Owner_ID where c.CarNumber = " + textEdit1.Text.Trim();
                    DataTable dt = DbConnection.DBConnect(CheckOwner);
                    if (dt.Rows.Count > 0)
                    {
                        textEdit2.Text = dt.Rows[0][0].ToString();
                    }
                    else
                    {
                        textEdit2.Text = "";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void textEdit1_Properties_KeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {
                char chr = e.KeyChar;
                if (!Char.IsDigit(chr) && chr != 8)
                {
                    e.Handled = true;
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
                if(SelectItemRow == 0)
                {
                    string Add = "exec [dbo].[FillRenderedService] " + textEdit1.Text.Trim() + "," + textEdit4.Text.Trim() + "," + textEdit6.Text.Trim() + "," + textEdit8.Text.Trim() + "," + textEdit7.Text.Trim() + "," + textEdit9.Text.Trim() + "," + textEdit10.Text.Trim() + "," + textEdit11.Text.Trim() + "," + textEdit5.Text.Trim() + "," + comboBox1.SelectedValue.ToString() + ",'" + textEdit3.Text.Trim() + "'," + comboBox2.SelectedValue.ToString() + ",NULL";
                    DbConnection.DBConnect(Add);
                    Refresh();
                }
                else
                {
                    string Update = "exec [dbo].[UpdateRenderedService] " + textEdit1.Text.Trim() + "," + textEdit4.Text.Trim() + "," + textEdit6.Text.Trim() + "," + textEdit8.Text.Trim() + "," + textEdit7.Text.Trim() + "," + textEdit9.Text.Trim() + "," + textEdit10.Text.Trim() + "," + textEdit11.Text.Trim() + "," + textEdit5.Text.Trim() + "," + comboBox1.SelectedValue.ToString() + ",'" + textEdit3.Text.Trim() + "'," + comboBox2.SelectedValue.ToString() + ","+SelectItemRow;
                    DbConnection.DBConnect(Update);
                    Refresh();
                    Block();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void gridView1_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            string Id = gridView1.GetFocusedDataRow()[0].ToString();
            SelectItemRow = Convert.ToInt32(Id);
            SelectBrigadeID = gridView1.GetFocusedDataRow()[1].ToString();
            SelectProductID = gridView1.GetFocusedDataRow()[2].ToString();
            comboBox1.DataBindings.Clear();
            comboBox2.DataBindings.Clear();
            Block();
            Update(SelectItemRow);
        }

        public void Update(int Id)
        {
            string GetValue = "select * from d__RenderedService where ID = "+Id;
            DataTable dt = DbConnection.DBConnect(GetValue);
            bool HOL = Convert.ToBoolean(dt.Rows[0][3].ToString());
            int HOL1;
            if(HOL == true)
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

            string Brigade = "Select * from d__Brigade";
            DataTable dt1 = DbConnection.DBConnect(Brigade);
            comboBox1.DataSource = dt1;
            comboBox1.DisplayMember = "Name";
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

        private void Block()
        {
            textEdit1.Enabled = false;
            textEdit3.Enabled = false;
            textEdit4.Enabled = false;
            textEdit6.Enabled = false;
            textEdit8.Enabled = false;
            textEdit7.Enabled = false;
            textEdit9.Enabled = false;
            textEdit10.Enabled = false;
            textEdit11.Enabled = false;
            textEdit5.Enabled = false;
            comboBox1.Enabled = false;
            comboBox2.Enabled = false;
            simpleButton1.Enabled = false;
        }
        private void Unblock()
        {
            textEdit1.Enabled = true;
            textEdit3.Enabled = true;
            textEdit4.Enabled = true;
            textEdit6.Enabled = true;
            textEdit8.Enabled = true;
            textEdit7.Enabled = true;
            textEdit9.Enabled = true;
            textEdit10.Enabled = true;
            textEdit11.Enabled = true;
            textEdit5.Enabled = true;
            comboBox1.Enabled = true;
            comboBox2.Enabled = true;
            simpleButton1.Enabled = true;
        }

        private void simpleButton4_Click(object sender, EventArgs e)
        {
            try
            {
                string delete = "delete from d__RenderedService where ID = " + SelectItemRow;
                DbConnection.DBConnect(delete);
                Refresh();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                textEdit8.Text = "1";
                textEdit6.Text = "0";
                textEdit10.Text = "0";
            }
        }

        private void textEdit10_Properties_EditValueChanged(object sender, EventArgs e)
        {
            if (textEdit10.Text == "1")
            {
                textEdit7.Text = "0";
                textEdit8.Text = "1";
                textEdit6.Text = "0";
                textEdit9.Text = "0";
            }
        }

        private void textEdit7_Properties_EditValueChanged(object sender, EventArgs e)
        {
            if(textEdit7.Text == "1")
            {
                textEdit6.Text = "0";
                textEdit8.Text = "1";
                textEdit9.Text = "0";
                textEdit10.Text = "0";
            }
        }

        private void textEdit6_Properties_EditValueChanged(object sender, EventArgs e)
        {
            if(textEdit6.Text == "1")
            {
                textEdit7.Text = "0";
                textEdit8.Text = "0";
                textEdit9.Text = "0";
                textEdit10.Text = "0";
            }
        }

        private void textEdit8_Properties_EditValueChanged(object sender, EventArgs e)
        {
            if (textEdit8.Text == "1")
            {
                textEdit6.Text = "0";
            }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            Refresh();
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            Unblock();
            textEdit1.Text = "";
            SelectItemRow = 0;
        }

        private void simpleButton3_Click(object sender, EventArgs e)
        {
            Unblock();
        }

        private void simpleButton5_Click(object sender, EventArgs e)
        {
            Refresh();
        }
    }
}
