﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Учет_цистерн.Forms.Услуги.СНО_Приход
{
    public partial class SnoComForm : Form
    {
        BindingSource source = new BindingSource();

        public SnoComForm()
        {
            InitializeComponent();
        }

        int SelectItemRow;
 
        private void dataGridView1_FilterStringChanged(object sender, EventArgs e)
        {
            this.source.Filter = this.dataGridView1.FilterString;
        }

        private void dataGridView1_SortStringChanged(object sender, EventArgs e)
        {
            this.source.Sort = this.dataGridView1.SortString;
        }

        private void SnoComForm_Load(object sender, EventArgs e)
        {
            GetSNO();

            panel1.Visible = false;
            panel2.Visible = false;
            panel3.Visible = false;

            textBox1.Visible = false;
            textBox2.Visible = false;
            textBox3.Visible = false;
            textBox4.Visible = false;
        }

        public void GetSNO()
        {
            try
            {
                string GetSNO = "exec dbo.GetCurrentSNO";
                DataTable dataTable = new DataTable();
                dataTable = DbConnection.DBConnect(GetSNO);
                source.DataSource = dataTable;
                dataGridView1.DataSource = source;
                dataGridView1.Columns[0].Visible = false;
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

        private void dataGridView1_CellPainting_1(object sender, DataGridViewCellPaintingEventArgs e)
        {
            try
            {
                Decimal Capacity = 0;
                Decimal Cost = 0;
                Decimal SumOutVat = 0;
                Decimal SumVat = 0;
                for (int i = 0; i < this.dataGridView1.Rows.Count; i++)
                {
                    if (dataGridView1.Rows[i].Cells[3].Value.ToString() != string.Empty)
                    {
                        Capacity += Convert.ToDecimal(this.dataGridView1[3, i].Value);
                    }
                    if (dataGridView1.Rows[i].Cells[4].Value.ToString() != string.Empty)
                    {
                        Cost += Convert.ToDecimal(this.dataGridView1[4, i].Value);
                    }
                    if (dataGridView1.Rows[i].Cells[5].Value.ToString() != string.Empty)
                    {
                        SumOutVat += Convert.ToDecimal(this.dataGridView1[5, i].Value);
                    }
                    if (dataGridView1.Rows[i].Cells[7].Value.ToString() != string.Empty)
                    {
                        SumVat += Convert.ToDecimal(this.dataGridView1[7, i].Value);
                    }
                }

                panel4.Width = this.dataGridView1.RowHeadersWidth - 3;
                panel4.Location = new Point(5, this.dataGridView1.Height - (panel4.Height - 15));
                panel4.Visible = true;

                int Xdgvx_Panel1 = this.dataGridView1.GetCellDisplayRectangle(1, -1, true).Location.X;
                panel1.Width = this.dataGridView1.Columns[2].Width;
                Xdgvx_Panel1 = this.dataGridView1.GetCellDisplayRectangle(2, -1, true).Location.X;
                panel1.Location = new Point(Xdgvx_Panel1, this.dataGridView1.Height - (panel1.Height - 15));
                panel1.Visible = true;

                int Xdgvx_TextBox1 = this.dataGridView1.GetCellDisplayRectangle(2, -1, true).Location.X;
                textBox1.Width = this.dataGridView1.Columns[3].Width;
                Xdgvx_TextBox1 = this.dataGridView1.GetCellDisplayRectangle(3, -1, true).Location.X;
                textBox1.Location = new Point(Xdgvx_TextBox1, this.dataGridView1.Height - (textBox1.Height - 15));
                textBox1.Text = Capacity.ToString();
                textBox1.Visible = true;

                int Xdgvx_TextBox2 = this.dataGridView1.GetCellDisplayRectangle(3, -1, true).Location.X;
                textBox2.Width = this.dataGridView1.Columns[4].Width;
                Xdgvx_TextBox2 = this.dataGridView1.GetCellDisplayRectangle(4, -1, true).Location.X;
                textBox2.Location = new Point(Xdgvx_TextBox2, this.dataGridView1.Height - (textBox2.Height - 15));
                textBox2.Text = Cost.ToString();
                textBox2.Visible = true;

                int Xdgvx_TextBox3 = this.dataGridView1.GetCellDisplayRectangle(4, -1, true).Location.X;
                textBox3.Width = this.dataGridView1.Columns[5].Width;
                Xdgvx_TextBox3 = this.dataGridView1.GetCellDisplayRectangle(5, -1, true).Location.X;
                textBox3.Location = new Point(Xdgvx_TextBox3, this.dataGridView1.Height - (textBox3.Height - 15));
                textBox3.Text = SumOutVat.ToString();
                textBox3.Visible = true;

                int Xdgvx_Panel2 = this.dataGridView1.GetCellDisplayRectangle(5, -1, true).Location.X;
                panel2.Width = this.dataGridView1.Columns[6].Width;
                Xdgvx_Panel2 = this.dataGridView1.GetCellDisplayRectangle(6, -1, true).Location.X;
                panel2.Location = new Point(Xdgvx_Panel2, this.dataGridView1.Height - (panel2.Height - 15));
                panel2.Visible = true;

                int Xdgvx_TextBox4 = this.dataGridView1.GetCellDisplayRectangle(6, -1, true).Location.X;
                textBox4.Width = this.dataGridView1.Columns[7].Width;
                Xdgvx_TextBox4 = this.dataGridView1.GetCellDisplayRectangle(7, -1, true).Location.X;
                textBox4.Location = new Point(Xdgvx_TextBox4, this.dataGridView1.Height - (textBox4.Height - 15));
                textBox4.Text = SumVat.ToString();
                textBox4.Visible = true;

                int Xdgvx_Panel3 = this.dataGridView1.GetCellDisplayRectangle(7, -1, true).Location.X;
                panel3.Width = this.dataGridView1.Columns[8].Width + 2;
                Xdgvx_Panel3 = this.dataGridView1.GetCellDisplayRectangle(8, -1, true).Location.X;
                panel3.Location = new Point(Xdgvx_Panel3, this.dataGridView1.Height - (panel3.Height - 15));
                panel3.Visible = true;
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                SnoComAddForm snoComAddForm = new SnoComAddForm();
                snoComAddForm.Show();
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                SnoComUpdateFrom snoComUpdateFrom = new SnoComUpdateFrom();
                snoComUpdateFrom.SelectID = SelectItemRow;
                snoComUpdateFrom.textBox1.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                snoComUpdateFrom.textBox2.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                snoComUpdateFrom.textBox3.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                snoComUpdateFrom.dateTimePicker1.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                snoComUpdateFrom.Show();
            }
            catch(Exception ex)
            {
                MessageBox.Show("Для редактирования записи, необходимо указать строку! " + ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if (SelectItemRow > 0) {
                    if (MessageBox.Show("Вы действительно хотите удалить эту запись?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        string Delete = "delete from d__CurrentSNO where ID = " + SelectItemRow;
                        DbConnection.DBConnect(Delete);
                        MessageBox.Show("Запись удалена!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GetSNO();
                    }
                }
                else
                {
                    MessageBox.Show("Для удаления записи, необходимо выбрать строку!", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            GetSNO();
        }

        private void dataGridView1_CellClick_1(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.RowIndex >= 0)
                {
                    DataGridViewRow row = this.dataGridView1.Rows[e.RowIndex];
                    string Id = row.Cells["ID"].Value.ToString();
                    SelectItemRow = Convert.ToInt32(Id);
                }
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
