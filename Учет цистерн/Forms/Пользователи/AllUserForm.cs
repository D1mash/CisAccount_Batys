using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using DevExpress.XtraGrid;

namespace Учет_цистерн.Forms.Пользователи
{
    public partial class AllUserForm : Form
    {
        string role = string.Empty;
        int SelectItemRow = 0;
        int SelectRoleId = 0;

        public AllUserForm(string role)
        {
            InitializeComponent();
            this.role = role;
        }

        private void Refreshh()
        {
            string Refreshh = "exec dbo.GetUser";
            DataTable dt = DbConnection.DBConnect(Refreshh);
            gridControl1.DataSource = dt;
            gridView1.Columns[0].Visible = false;
            gridView1.Columns[1].Visible = false;

            GridColumnSummaryItem Count = new GridColumnSummaryItem(DevExpress.Data.SummaryItemType.Count, "Имя", "{0}");
            gridView1.Columns["Имя"].Summary.Add(Count);
        }

        public void AllUserForm_Load(object sender, EventArgs e)
        {
            Refreshh();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                foreach (Form form in Application.OpenForms)
                {
                    if (form.GetType() == typeof(AddUserForm))
                    {
                        form.Activate();
                        return;
                    }
                }
                AddUserForm snoComAddForm = new AddUserForm();
                snoComAddForm.Owner = this;
                snoComAddForm.ShowDialog();
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
                if(SelectItemRow != 0)
                {
                    foreach (Form form in Application.OpenForms)
                    {
                        if (form.GetType() == typeof(UpdateUserForm))
                        {
                            form.Activate();
                            return;
                        }
                    }
                    UpdateUserForm update = new UpdateUserForm(role);
                    update.SelectID = SelectItemRow;
                    update.SelectRoleId = SelectRoleId;
                    update.Owner = this;
                    update.ShowDialog();
                }
                else
                {
                    MessageBox.Show("Для редактирования записи, необходимо указать строку! ", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if (gridView1.SelectedRowsCount > 0)
                {
                    try
                    {
                        if (SelectItemRow != 0)
                        {
                            if (MessageBox.Show("Вы действительно хотите удалить эту запись?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                            {
                                string Delete = "update Users set IsDeleted = 1 where AID = " + SelectItemRow;
                                DbConnection.DBConnect(Delete);
                                MessageBox.Show("Пользователь удалён!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                Refreshh();
                            }
                        }
                        else
                        {
                            MessageBox.Show("Для удаления записи, необходимо указать строку! ", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    catch (Exception exp)
                    {
                        MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Для удаления записи, необходимо выбрать строку полностью!", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void gridView1_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            try
            {
                string Id = gridView1.GetFocusedDataRow()[0].ToString();
                string RoleId = gridView1.GetFocusedDataRow()[1].ToString();
                SelectItemRow = Convert.ToInt32(Id);
                SelectRoleId = Convert.ToInt32(RoleId);
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
