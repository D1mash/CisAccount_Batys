using System;
using System.Data;
using System.Windows.Forms;

namespace Учет_цистерн.Forms.Пользователи
{
    public partial class UpdateUserForm : Form
    {
        int selectID;
        string role = string.Empty;
        string Fristname = string.Empty;
        string Surname = string.Empty;

        public int SelectID
        {
            get { return SelectID; }
            set { selectID = value; }
        }

        public int SelectRoleId { get; set; }

        public UpdateUserForm(string role)
        {
            InitializeComponent();
            this.role = role;
        }

        private void Fillcombobox()
        {
            try
            {
                string Role = "select * from d__Role";
                DataTable dt = DbConnection.DBConnect(Role);
                comboBox1.DataSource = dt;
                comboBox1.DisplayMember = "Name";
                comboBox1.ValueMember = "ID";
                comboBox1.DataBindings.Add("SelectedValue", this, "SelectRoleId", true, DataSourceUpdateMode.OnPropertyChanged);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UpdateUserForm_Load(object sender, EventArgs e)
        {
            Fillcombobox();
            SelectUser();
        }

        private void SelectUser()
        {
            try
            {
                string Role = "SELECT * FROM [Users] where AID =" + selectID;
                DataTable dt = DbConnection.DBConnect(Role);

                textBox2.Text = dt.Rows[0][1].ToString();
                textBox3.Text = dt.Rows[0][2].ToString();
                textBox4.Text = dt.Rows[0][3].ToString();
                textBox1.Text = dt.Rows[0][4].ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if(textBox1.Text == string.Empty)
                {
                    MessageBox.Show("Введите пароль!", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    string Update = "exec dbo.UpdateUser '" + textBox2.Text.Trim() + "','" + textBox3.Text.Trim() + "','" + textBox4.Text.Trim() + "','" + textBox1.Text.Trim() + "'," + comboBox1.SelectedValue.ToString() + "," + selectID;
                    DbConnection.DBConnect(Update);

                    MessageBox.Show("Пользователь изменён!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    AllUserForm allUserForm = this.Owner as AllUserForm;
                    allUserForm.Refresh();

                    selectID = 0;
                    role = string.Empty;
                    Fristname = string.Empty;
                    Surname = string.Empty;

                    this.Close();
                }
            }
            catch(Exception ex)
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
    }
}
