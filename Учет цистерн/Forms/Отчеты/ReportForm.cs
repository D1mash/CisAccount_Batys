using System;
using System.Data;
using System.Windows.Forms;
using Учет_цистерн.Forms.Отчеты;


namespace Учет_цистерн
{
    public partial class ReportForm : Form
    {
        DataTable dt;
        string role;

        public ReportForm(string role)
        {
            InitializeComponent();
            this.role = role;
        }

        private new void Refresh()
        {
            if (checkBox1.Checked)
            {
                if (comboBox2.SelectedIndex == 0)
                {
                    string Itog_All_Report = "exec dbo.Itog_All_Report '" + dateTimePicker1.Value.Date.ToString() + "','" + dateTimePicker2.Value.Date.ToString() + "'";
                    dt = DbConnection.DBConnect(Itog_All_Report);

                    toolStripLabel1.Text = TotalRow(dt).ToString();
                }
                else
                {
                    string Itog_Report = "exec dbo.Itog_Report  '" + dateTimePicker1.Value.Date.ToString() + "','" + dateTimePicker2.Value.Date.ToString() + "','" + comboBox2.SelectedValue + "'";
                    dt = DbConnection.DBConnect(Itog_Report);

                    toolStripLabel1.Text = TotalRow(dt).ToString();
                }
            }
            else
            {
                if (checkBox2.Checked)
                {
                    string RefreshAll = "exec dbo.GetReportAllRenderedService_v1 '" + dateTimePicker1.Value.Date.ToString() + "','" + dateTimePicker2.Value.Date.ToString() + "'," + "@Type = " + 1;
                    dt = DbConnection.DBConnect(RefreshAll);

                    toolStripLabel1.Text = TotalRow(dt).ToString();
                }
                else
                {
                    if (checkBox5.Checked)
                    {
                        string RefreshAll = "exec [dbo].[GetReportAUTN] '" + dateTimePicker1.Value.Date.ToString() + "','" + dateTimePicker2.Value.Date.ToString() + "'";
                        dt = DbConnection.DBConnect(RefreshAll);

                        toolStripLabel1.Text = TotalRow(dt).ToString();
                    }
                    else
                    {
                        if (checkBox4.Checked)
                        {
                            string GetSNO = "exec dbo.GetSNO '" + dateTimePicker1.Value.Date.ToShortDateString() + "', '" + dateTimePicker2.Value.Date.ToShortDateString() + "'";
                            dt = DbConnection.DBConnect(GetSNO);
                        }
                        else
                        {
                            if (checkBox6.Checked)
                            {
                                string GetSNO = "exec dbo.GetCurrentSNO '" + dateTimePicker1.Value.Date.ToShortDateString() + "', '" + dateTimePicker2.Value.Date.ToShortDateString() + "'";
                                dt = DbConnection.DBConnect(GetSNO);
                            }
                        }
                    }
                }
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
            if (role == "1" || role == "2")
            {
                checkBox4.Enabled = false;
                checkBox6.Enabled = false;
            }
            else
            {
                return;
            }

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

            var startDate = now;
            var endDate = now;

            dateTimePicker1.Value = startDate;
            dateTimePicker2.Value = endDate;

            //Refresh();
            checkBox3_CheckedChanged(null, null);
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            try
            {
                if (checkBox1.Checked)
                {

                    //Итоговый реестр

                    if (comboBox2.SelectedIndex == 0)
                    {
                        Refresh();

                        var Tbr = new TotalByStation();
                        Tbr.TBS(dateTimePicker1.Value.ToShortDateString(), dateTimePicker2.Value.ToShortDateString(), dt);

                        dt.Clear();
                    }
                    else
                    {
                        // Итоговый реестр по собственнику

                        var Tbr = new TotalByStation();
                        Tbr.TotalByCompany(dateTimePicker1.Value.Date.ToString(), dateTimePicker2.Value.Date.ToString(), comboBox2.SelectedValue, comboBox2.Text, dateTimePicker1.Value.ToShortDateString(), dateTimePicker2.Value.ToShortDateString());
                    }
                }
                else
                {
                    // АУТН

                    if (checkBox5.Checked)
                    {
                        Refresh();

                        var Autn = new AUTN();
                        Autn.AUTN_Report(dt);

                        dt.Clear();
                    }
                    else
                    {
                        //Общий реестр

                        if (checkBox2.Checked)
                        {
                            if (comboBox2.SelectedIndex == 0)
                            {
                                var GR = new General_Reestr();
                                GR.General_Reesters(dateTimePicker1.Value.Date.ToString(), dateTimePicker2.Value.Date.ToString());
                            }
                            else
                            {

                                // СНО Реализация
                                if (checkBox4.Checked)
                                {
                                    Refresh();
                                    var Sno = new SNO();
                                    Sno.SNO_OUT(dateTimePicker1.Value.Date.ToShortDateString(), dateTimePicker2.Value.Date.ToShortDateString(), dt);
                                    dt.Clear();
                                }
                                else
                                {
                                    // СНО Приход
                                    if (checkBox6.Checked)
                                    {
                                        Refresh();

                                        var Sno = new SNO();
                                        Sno.SNO_IN(dateTimePicker1.Value.Date.ToShortDateString(), dateTimePicker2.Value.Date.ToShortDateString(), dt);
                                        dt.Clear();
                                    }
                                }
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

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                checkBox2.Checked = false;
                checkBox5.Checked = false;
                checkBox4.Checked = false;
                checkBox6.Checked = false;
            }
            else
            {
                checkBox1.Checked = false;
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            if (dateTimePicker1.Checked)
            {
                if (checkBox3.Checked)
                {
                    return;
                }
                else
                {
                    dateTimePicker2.Value = dateTimePicker1.Value;
                }
            }
            else
            {
                return;
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                checkBox1.Checked = false;
                comboBox2.Enabled = false;
                checkBox5.Checked = false;
                comboBox2.SelectedIndex = 0;
                checkBox4.Checked = false;
                checkBox6.Checked = false;
            }
            else
            {
                comboBox2.Enabled = true;
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked)
            {
                dateTimePicker2.Enabled = (checkBox3.CheckState == CheckState.Checked);
                dateTimePicker2.Value = DateTime.Today;
            }
            else
            {
                dateTimePicker2.Enabled = false;
                dateTimePicker2.Value = dateTimePicker1.Value;
            }
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked)
            {
                comboBox2.SelectedIndex = 0;
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                comboBox2.Enabled = false;
                checkBox4.Checked = false;
                checkBox6.Checked = false;
            }
            else
            {
                checkBox5.Checked = false;
                comboBox2.Enabled = true;
                //Refresh();
            }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked)
            {
                comboBox2.SelectedIndex = 0;
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                comboBox2.Enabled = false;
                checkBox6.Checked = false;
                checkBox5.Checked = false;
            }
            else
            {
                checkBox4.Checked = false;
                comboBox2.Enabled = true;
            }
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox6.Checked)
            {
                comboBox2.SelectedIndex = 0;
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                comboBox2.Enabled = false;
                checkBox5.Checked = false;
                checkBox4.Checked = false;
            }
            else
            {
                checkBox5.Checked = false;
                comboBox2.Enabled = true;
            }
        }
    }
}