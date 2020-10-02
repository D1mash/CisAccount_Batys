using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace Учет_цистерн.Forms.Пользователи
{
    public partial class AddUserForm : Form
    {
        string Fristname = string.Empty;
        string Surname = string.Empty;

        public AddUserForm()
        {
            InitializeComponent();
        }

        private void FillCombobox()
        {
            string Role = "select * from d__Role";
            DataTable dt = DbConnection.DBConnect(Role);
            comboBox1.DataSource = dt;
            comboBox1.DisplayMember = "Name";
            comboBox1.ValueMember = "ID";
        }

        private void AddUserForm_Load(object sender, EventArgs e)
        {
            FillCombobox();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string FillUser = "exec dbo.FillUser '"+textBox2.Text.Trim()+"','"+ textBox3.Text.Trim() + "','"+ textBox4.Text.Trim() + "','"+ textBox1.Text.Trim() + "','"+comboBox1.SelectedValue.ToString()+"'";
                DbConnection.DBConnect(FillUser);
                MessageBox.Show("Пользователь добавлен!","",MessageBoxButtons.OK,MessageBoxIcon.Information);

                AllUserForm allUserForm = this.Owner as AllUserForm;
                allUserForm.Refresh();

                this.Close();
            }
            catch(SqlException ex)
            {
                MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            Fristname = textBox2.Text;
            textBox4_Enter(null, null);
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            Surname = textBox3.Text;
            textBox4_Enter(null, null);
        }

        private void textBox4_Enter(object sender, EventArgs e)
        {
            textBox4.Text = Surname + " " + Fristname;
        }

        private void AddUserForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            string Fristname = string.Empty;
            string Surname = string.Empty;
        }
    }
}
