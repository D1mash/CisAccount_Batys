using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Учет_цистерн.Forms.заявки_на_обработку;
using Учет_цистерн.Forms.СНО;
using Учет_цистерн.Forms.Справка;
using Учет_цистерн.Forms.Услуги.СНО_Приход;
using AutoUpdaterDotNET;
using NLog;
using System.Configuration;
using System.Collections.Specialized;
using Учет_цистерн.Forms;
using System.Collections;
using Учет_цистерн.Forms.Пользователи;
using Учет_цистерн.Forms.Обработанные_вагоны;
using Учет_цистерн.Forms.Смена_собственника;
using Учет_цистерн.Forms.Доп.параметры;
using Учет_цистерн.Forms.АУТН;

namespace Учет_цистерн
{
    public partial class MainForm : Form
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();

        string sUrl = ConfigurationManager.AppSettings["Url"].ToString();

        string UserFIO, role, UserID;

        public MainForm(string FIO, string role, string aid)
        {
            InitializeComponent();
            this.tabControl2.SelectedTab = tabPage2;
            this.splitContainer1.SplitterDistance = 25;
            string GetConnection = "exec dbo.GetConnection";
            DataTable dt = DbConnection.DBConnect(GetConnection);
            toolStripTextBox1.Text = "SPID: " + dt.Rows[0][1].ToString() + "; UID: " + dt.Rows[0][0].ToString() + ";";
            this.Text = "Учет вагонов-цистерн. " + dt.Rows[0][2].ToString();
            this.UserFIO = FIO;
            this.role = role;
            this.UserID = aid;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            try
            {
                GetFilter();
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                logger.Error(exp, "MainForm_Load");
            }
        }

        public void GetFilter()
        {
            try
            {
                string Reffresh = "exec dbo.GetFilter '"+UserID+"'";
                gridControl1.DataSource = DbConnection.DBConnect(Reffresh); ;
                gridView1.Columns[0].Visible = false;
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                logger.Error(exp, "GetFilter");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                contextMenuStrip_Product.Show(button1, new Point(0, button1.Height));
                if(role == "1")
                {
                    contextMenuStrip_Product.Items[0].Enabled = true;
                    contextMenuStrip_Product.Items[1].Enabled = true;
                    contextMenuStrip_Product.Items[2].Enabled = true;
                    contextMenuStrip_Product.Items[3].Enabled = true;
                    contextMenuStrip_Product.Items[4].Enabled = true;
                }
                else
                {
                    if(role == "2")
                    {
                        contextMenuStrip_Product.Items[0].Enabled = true;
                        contextMenuStrip_Product.Items[1].Enabled = true;
                        contextMenuStrip_Product.Items[2].Enabled = true;
                        contextMenuStrip_Product.Items[3].Enabled = true;
                        contextMenuStrip_Product.Items[4].Enabled = true;
                    }
                    else
                    {
                        contextMenuStrip_Product.Items[0].Enabled = true;
                        contextMenuStrip_Product.Items[1].Enabled = true;
                        contextMenuStrip_Product.Items[2].Enabled = true;
                        contextMenuStrip_Product.Items[3].Enabled = true;
                        contextMenuStrip_Product.Items[4].Enabled = true;
                    }
                }
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                logger.Error(exp, "button1_Click");
            }
        }

        //Глобальный фильтр, вставить вагоны
        private void вставитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                string Insert = "exec dbo.InsertGlobalFilter '"+UserID+"','" + Clipboard.GetText() + "'";
                DbConnection.DBConnect(Insert);
                GetFilter();
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                logger.Error(exp, "Глобальный фильтр, вставить вагоны");
            }
        }

        //Глобальный фильтр, обновить вагоны
        private void обновитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GetFilter();
        }
        //Глобальный фильтр, удалить выделенные вагоны
        private void удалитьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
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
                    string delete = "exec dbo.RemoveGlobalFilter '"+UserID+"','" + Arrays + "'";
                    DbConnection.DBConnect(delete);
                }
                GetFilter();
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                logger.Error(ex, "удалитьToolStripMenuItem1_Click_SQL");
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                logger.Error(exp, "удалитьToolStripMenuItem1_Click");
            }
        }
        //Глобальный фильтр, удалить все вагоны
        private void удалитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                string DeleteAll = "delete from dbo.GlobalFilter where UserID = '"+UserID+"'";
                DbConnection.DBConnect(DeleteAll);
                GetFilter();
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                logger.Error(ex, "GlobalудалитьToolStripMenuItem1_Click_SQL");
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                logger.Error(exp, "GlobalудалитьToolStripMenuItem1_Click_SQL");
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                dynamic result = MessageBox.Show("Завершить работу с программой?", "Завершение программы", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    string Truncate = "exec dbo.Delete_Temp_MultiCar '"+UserID+"'";
                    DbConnection.DBConnect(Truncate);

                    string UpdateAuditUser = "UPDATE AUDIT_USER SET DATE_OUT = GETDATE(), IS_DEAD = 1 WHERE ID_USER = '"+UserID+"' and (IS_DEAD IS NULL OR DATE_OUT IS NULL)";
                    DbConnection.DBConnect(UpdateAuditUser);
                    
                    Environment.Exit(0);
                }
                else
                {
                    e.Cancel = true;
                }
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                logger.Error(exp, "MainForm_FormClosing");
                logger.Info(exp, "MainForm_FormClosing");
            }
        }

        private void ToolStripMenuItem1_Product_Click(object sender, EventArgs e)
        {
            try
            {
                foreach (Form form in Application.OpenForms)
                {
                    if (form.GetType() == typeof(Form_Product))
                    {
                        form.Activate();
                        return;
                    }
                }

                Form_Product frm = new Form_Product(role, UserID);
                tabControl1.Show();
                TabPage ProductTabPage = new TabPage("Продукты");
                tabControl1.TabPages.Add(ProductTabPage);
                tabControl1.SelectedTab = ProductTabPage;
                frm.TopLevel = false;
                frm.Visible = true;
                frm.FormBorderStyle = FormBorderStyle.None;
                frm.Dock = DockStyle.Fill;
                ProductTabPage.Controls.Add(frm);
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                logger.Error(exp, "ToolStripMenuItem1_Product_Click");
            }
        }

        private void toolStripMenuItem2_Station_Click(object sender, EventArgs e)
        {
            try
            {
                foreach (Form form in Application.OpenForms)
                {
                    if (form.GetType() == typeof(StationForm))
                    {
                        form.Activate();
                        return;
                    }
                }

                StationForm frm = new StationForm(role,UserID);
                tabControl1.Show();
                TabPage StationTabPage = new TabPage("Станции");
                tabControl1.TabPages.Add(StationTabPage);
                tabControl1.SelectedTab = StationTabPage;
                frm.TopLevel = false;
                frm.Visible = true;
                frm.FormBorderStyle = FormBorderStyle.None;
                frm.Dock = DockStyle.Fill;
                StationTabPage.Controls.Add(frm);
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                logger.Error(exp, "toolStripMenuItem2_Station_Click");
            }
        }

        private void ToolStripMenuItem3_Brigade_Click(object sender, EventArgs e)
        {
            try
            {
                foreach (Form form in Application.OpenForms)
                {
                    if (form.GetType() == typeof(BrigadeForm))
                    {
                        form.Activate();
                        return;
                    }
                }

                BrigadeForm frm = new BrigadeForm(role);
                tabControl1.Show();
                TabPage BrigadeTabPage = new TabPage("Бригадиры");
                tabControl1.TabPages.Add(BrigadeTabPage);
                tabControl1.SelectedTab = BrigadeTabPage;
                frm.TopLevel = false;
                frm.Visible = true;
                frm.FormBorderStyle = FormBorderStyle.None;
                frm.Dock = DockStyle.Fill;
                BrigadeTabPage.Controls.Add(frm);
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                logger.Error(exp, "ToolStripMenuItem3_Brigade_Click");
            }
        }

        private void toolStripMenuItem4_Owner_Click(object sender, EventArgs e)
        {
            try
            {
                foreach (Form form in Application.OpenForms)
                {
                    if (form.GetType() == typeof(OwnerForm))
                    {
                        form.Activate();
                        return;
                    }
                }

                OwnerForm frm = new OwnerForm(role);
                tabControl1.Show();
                TabPage OwnerTabPage = new TabPage("Собственники");
                tabControl1.TabPages.Add(OwnerTabPage);
                tabControl1.SelectedTab = OwnerTabPage;
                frm.TopLevel = false;
                frm.Visible = true;
                frm.FormBorderStyle = FormBorderStyle.None;
                frm.Dock = DockStyle.Fill;
                OwnerTabPage.Controls.Add(frm);
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                logger.Error(exp, "toolStripMenuItem4_Owner_Click");
            }
        }

        private void ToolStripMenuItem_Carriage_Click(object sender, EventArgs e)
        {
            try
            {
                foreach (Form form in Application.OpenForms)
                {
                    if (form.GetType() == typeof(CarriageForm))
                    {
                        form.Activate();
                        return;
                    }
                }

                CarriageForm carriageForm = new CarriageForm(role);
                tabControl1.Show();
                TabPage CarriageTabPage = new TabPage("Вагоны");
                tabControl1.TabPages.Add(CarriageTabPage);
                tabControl1.SelectedTab = CarriageTabPage;
                carriageForm.TopLevel = false;
                carriageForm.Visible = true;
                carriageForm.FormBorderStyle = FormBorderStyle.None;
                carriageForm.Dock = DockStyle.Fill;
                CarriageTabPage.Controls.Add(carriageForm);
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                logger.Error(exp, "ToolStripMenuItem_Carriage_Click");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                contextMenuStrip_Services.Show(button2, new Point(0, button2.Height));
                if(role == "1")
                {
                    contextMenuStrip_Services.Items[0].Enabled = true;
                    contextMenuStrip_Services.Items[1].Enabled = false;
                }
                else
                {
                    if(role == "2")
                    {
                        contextMenuStrip_Services.Items[0].Enabled = true;
                        contextMenuStrip_Services.Items[1].Enabled = false;
                    }
                    else
                    {
                        contextMenuStrip_Services.Items[0].Enabled = true;
                        contextMenuStrip_Services.Items[1].Enabled = true;
                    }
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                logger.Error(ex, "button2_Click_MainForm");
            }
        }


        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                RenderedServiceForm RenderedServiceForm = new RenderedServiceForm(this.toolStripProgressBar1, this.toolStripLabel1, this.button1, this.button2, this.button3, this.button4, this.btn_Refrence, this.tabControl1);
                tabControl1.Show();
                TabPage CarriageTabPage = new TabPage("Обработанные вагоны");
                tabControl1.TabPages.Add(CarriageTabPage);
                tabControl1.SelectedTab = CarriageTabPage;
                RenderedServiceForm.TopLevel = false;
                RenderedServiceForm.Visible = true;
                RenderedServiceForm.FormBorderStyle = FormBorderStyle.None;
                RenderedServiceForm.Dock = DockStyle.Fill;
                CarriageTabPage.Controls.Add(RenderedServiceForm);
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                logger.Error(exp, "button4_Click_MainForm");
            }
        }


        private Point DragStartPosition = Point.Empty;

        private void tabControl1_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            DragStartPosition = new Point(e.X, e.Y);
        }


        private void tabControl1_MouseMove(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            try
            {
                if (e.Button != MouseButtons.Left) return;

                Rectangle r = new Rectangle(DragStartPosition, Size.Empty);
                r.Inflate(SystemInformation.DragSize);

                TabPage tp = HoverTab();

                if (tp != null)
                {
                    if (!r.Contains(e.X, e.Y))
                        tabControl1.DoDragDrop(tp, DragDropEffects.All);
                }
                DragStartPosition = Point.Empty;
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                logger.Error(exp, "tabControl1_MouseMove");
            }
        }


        private void tabControl1_DragOver(object sender, System.Windows.Forms.DragEventArgs e)
        {
            try
            {
                TabPage hover_Tab = HoverTab();
                if (hover_Tab == null)
                    e.Effect = DragDropEffects.None;
                else
                {
                    if (e.Data.GetDataPresent(typeof(TabPage)))
                    {
                        e.Effect = DragDropEffects.Move;
                        TabPage drag_tab = (TabPage)e.Data.GetData(typeof(TabPage));

                        if (hover_Tab == drag_tab) return;

                        Rectangle TabRect = tabControl1.GetTabRect(tabControl1.TabPages.IndexOf(hover_Tab));
                        TabRect.Inflate(-3, -3);
                        if (TabRect.Contains(tabControl1.PointToClient(new Point(e.X, e.Y))))
                        {
                            SwapTabPages(drag_tab, hover_Tab);
                            tabControl1.SelectedTab = drag_tab;
                        }
                    }
                }
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                logger.Error(exp, "tabControl1_DragOver");
            }
        }


        private TabPage HoverTab()
        {
            for (int index = 0; index <= tabControl1.TabCount - 1; index++)
            {
                if (tabControl1.GetTabRect(index).Contains(tabControl1.PointToClient(Cursor.Position)))
                return tabControl1.TabPages[index];
            }
            return null;
        }


        private void SwapTabPages(TabPage tp1, TabPage tp2)
        {
            try
            {
                int Index1 = tabControl1.TabPages.IndexOf(tp1);
                int Index2 = tabControl1.TabPages.IndexOf(tp2);
                tabControl1.TabPages[Index1] = tp2;
                tabControl1.TabPages[Index2] = tp1;
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                logger.Error(exp, "SwapTabPages");
            }
        }

        private void ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                
                foreach (Form form in Application.OpenForms)
                {
                    if (form.GetType() == typeof(ReportForm))
                    {
                        form.Activate();
                        return;
                    }
                }

                ReportForm reportForm = new ReportForm(role);
                reportForm.ShowDialog();
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                logger.Error(exp, "ToolStripMenuItem1_Click");
            }
        }

        private void button4_MouseHover(object sender, EventArgs e)
        {
            this.button4.BackColor = System.Drawing.Color.Pink;
        }

        private void button4_MouseLeave(object sender, EventArgs e)
        {
            this.button4.BackColor = System.Drawing.SystemColors.Control;
        }

        private void button4_MouseMove(object sender, MouseEventArgs e)
        {
            this.button4.BackColor = System.Drawing.Color.Pink;
        }

        private void tabControl2_Click(object sender, EventArgs e)
        {
            try
            {
                if (tabControl2.SelectedTab == tabPage1)
                {
                    this.splitContainer1.SplitterDistance = 250;
                    GetFilter();
                }
                else if (tabControl2.SelectedTab == tabPage2)
                {
                    this.splitContainer1.SplitterDistance = 25;
                }
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                logger.Error(exp, "tabControl2_Click");
            }
        }

        private void РасценкиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                foreach (Form form in Application.OpenForms)
                {
                    if (form.GetType() == typeof(ServiceCostForm))
                    {
                        form.Activate();
                        return;
                    }
                }

                ServiceCostForm ServiceCostForm = new ServiceCostForm(role);
                tabControl1.Show();
                TabPage CarriageTabPage = new TabPage("Расценки");
                tabControl1.TabPages.Add(CarriageTabPage);
                tabControl1.SelectedTab = CarriageTabPage;
                ServiceCostForm.TopLevel = false;
                ServiceCostForm.Visible = true;
                ServiceCostForm.FormBorderStyle = FormBorderStyle.None;
                ServiceCostForm.Dock = DockStyle.Fill;
                CarriageTabPage.Controls.Add(ServiceCostForm);
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                logger.Error(exp, "РасценкиToolStripMenuItem_Click");
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                OrderAllForm orderAllForm = new OrderAllForm(this.tabControl1, this.toolStripProgressBar1, this.toolStripLabel1, this.button1, this.button2, this.button3, this.button4, this.btn_Refrence, this.button7, role);
                tabControl1.Show();
                TabPage OrderAllTabPage = new TabPage("Журнал актов");
                tabControl1.TabPages.Add(OrderAllTabPage);
                OrderAllTabPage.Name = "OrderAllTabPage";
                tabControl1.SelectedTab = OrderAllTabPage;
                orderAllForm.TopLevel = false;
                orderAllForm.Visible = true;
                orderAllForm.FormBorderStyle = FormBorderStyle.None;
                orderAllForm.Dock = DockStyle.Fill;
                OrderAllTabPage.Controls.Add(orderAllForm);
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                logger.Error(exp, "button7_Click_MainForm");
            }
        }

        private void сНОРеализацияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                foreach (Form form in Application.OpenForms)
                {
                    if (form.GetType() == typeof(SnoImplForm))
                    {
                        form.Activate();
                        return;
                    }
                }

                SnoImplForm SnoImplForm = new SnoImplForm(role, UserID);
                tabControl1.Show();
                TabPage SnotabPage = new TabPage("СНО Реализация");
                tabControl1.TabPages.Add(SnotabPage);
                tabControl1.SelectedTab = SnotabPage;
                SnoImplForm.TopLevel = false;
                SnoImplForm.Visible = true;
                SnoImplForm.FormBorderStyle = FormBorderStyle.None;
                SnoImplForm.Dock = DockStyle.Fill;
                SnotabPage.Controls.Add(SnoImplForm);
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                logger.Error(exp, "сНОРеализацияToolStripMenuItem_Click");
            }
        }

        private void сНОПриходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                foreach (Form form in Application.OpenForms)
                {
                    if (form.GetType() == typeof(SnoComForm))
                    {
                        form.Activate();
                        return;
                    }
                }
                SnoComForm snoComForm = new SnoComForm(role, UserID);
                tabControl1.Show();
                TabPage SnotabPage = new TabPage("СНО приход");
                tabControl1.TabPages.Add(SnotabPage);
                tabControl1.SelectedTab = SnotabPage;
                snoComForm.TopLevel = false;
                snoComForm.Visible = true;
                snoComForm.FormBorderStyle = FormBorderStyle.None;
                snoComForm.Dock = DockStyle.Fill;
                SnotabPage.Controls.Add(snoComForm);
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                logger.Error(exp, "сНОПриходToolStripMenuItem_Click");
            }
        }

        private void проверитьОбновлениеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                AutoUpdater.Start(sUrl);
                AutoUpdater.ShowSkipButton = false;
                AutoUpdater.ShowRemindLaterButton = false;
                AutoUpdater.ReportErrors = true;
                AutoUpdater.DownloadPath = Environment.CurrentDirectory;
                AutoUpdater.RunUpdateAsAdmin = false;
                AutoUpdater.UpdateFormSize = new System.Drawing.Size(800, 600);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                logger.Error(ex, "проверитьОбновлениеToolStripMenuItem_Click");
            }
        }

        private void просмотрСправкиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Данная вкладка нахидится в разработке! Извините за неудобства ...", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (Form form in Application.OpenForms)
            {
                if (form.GetType() == typeof(AboutBox))
                {
                    form.Activate();
                    return;
                }
            }
            AboutBox aboutBox = new AboutBox();
            aboutBox.Show();
        }

        private void btn_Refrence_Click(object sender, EventArgs e)
        {
            try
            {
                contextMenuStrip_Refrence.Show(btn_Refrence, new Point(0, button2.Height));
                if(role == "1002")
                {
                    contextMenuStrip_Refrence.Items[0].Enabled = true;
                    contextMenuStrip_Refrence.Items[1].Enabled = true;
                    contextMenuStrip_Refrence.Items[2].Enabled = true;
                    contextMenuStrip_Refrence.Items[3].Enabled = true;
                }
                else
                {
                    if (role == "1")
                    {
                        contextMenuStrip_Refrence.Items[0].Enabled = false;
                        contextMenuStrip_Refrence.Items[1].Enabled = true;
                        contextMenuStrip_Refrence.Items[2].Enabled = true;
                        contextMenuStrip_Refrence.Items[3].Enabled = true;
                    }
                    else
                    {
                        contextMenuStrip_Refrence.Items[0].Enabled = false;
                        contextMenuStrip_Refrence.Items[1].Enabled = false;
                        contextMenuStrip_Refrence.Items[2].Enabled = true;
                        contextMenuStrip_Refrence.Items[3].Enabled = true;
                    }
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                logger.Error(ex, "btn_Refrence_Click_MainForm");
            }
        }

        private void gridControl1_MouseClick(object sender, MouseEventArgs e)
        {
            try
            {
                if (e.Button == MouseButtons.Right)
                {
                    //int currentMouseOverRow = gridView1.HitTest(e.X, e.Y).RowIndex;
                    contextMenuStrip_GlobalFilter.Show(gridControl1, new Point(e.X, e.Y));
                }
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                logger.Error(exp, "dataGridView1_MouseClick_1");
            }
        }

        private void пользователиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                foreach (Form form in Application.OpenForms)
                {
                    if (form.GetType() == typeof(AllUserForm))
                    {
                        form.Activate();
                        return;
                    }
                }

                AllUserForm all = new AllUserForm(role);
                tabControl1.Show();
                TabPage UserTabPage = new TabPage("Пользователи");
                tabControl1.TabPages.Add(UserTabPage);
                tabControl1.SelectedTab = UserTabPage;
                all.TopLevel = false;
                all.Visible = true;
                all.FormBorderStyle = FormBorderStyle.None;
                all.Dock = DockStyle.Fill;
                UserTabPage.Controls.Add(all);
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            contextMenuStrip_Rent_Car.Show(button5, new Point(0, button5.Height));
            if (role == "1")
            {
                contextMenuStrip_Rent_Car.Items[0].Enabled = true;
                contextMenuStrip_Rent_Car.Items[1].Enabled = true;
            }
            else
            {
                if (role == "2")
                {
                    contextMenuStrip_Rent_Car.Items[0].Enabled = true;
                    contextMenuStrip_Rent_Car.Items[1].Enabled = true;
                }
                else
                {
                    contextMenuStrip_Rent_Car.Items[0].Enabled = true;
                    contextMenuStrip_Rent_Car.Items[1].Enabled = true;
                }
            }
            }

        private void button6_Click(object sender, EventArgs e)
        {
            foreach (Form form in Application.OpenForms)
            {
                if (form.GetType() == typeof(Journal))
                {
                    form.Activate();
                    return;
                }
            }

            Journal journalForm = new Journal(role,UserID);
            tabControl1.Show();
            TabPage JournalPage = new TabPage("Журнал обработанных вагонов");
            tabControl1.TabPages.Add(JournalPage);
            tabControl1.SelectedTab = JournalPage;
            journalForm.TopLevel = false;
            journalForm.Visible = true;
            journalForm.FormBorderStyle = FormBorderStyle.None;
            journalForm.Dock = DockStyle.Fill;
            JournalPage.Controls.Add(journalForm);
        }

        private void сменаСобственникаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (Form form in Application.OpenForms)
            {
                if (form.GetType() == typeof(Change_of_Ownership))
                {
                    form.Activate();
                    return;
                }
            }

            string DateIns = System.DateTime.Now.ToString();

            Change_of_Ownership change_Of_Ownership = new Change_of_Ownership(this.tabControl1, role, UserID);
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
        }

        private void заявкаНаПередачуВцToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (Form form in Application.OpenForms)
            {
                if (form.GetType() == typeof(Rent_Brodcast_Car))
                {
                    form.Activate();
                    return;
                }
            }

            Rent_Brodcast_Car rent_Brodcast_Car = new Rent_Brodcast_Car(this.tabControl1,role, UserID);
            tabControl1.Show();
            TabPage chg_tabPage = new TabPage("Заявка на передачу в/ц");
            tabControl1.TabPages.Add(chg_tabPage);
            tabControl1.SelectedTab = chg_tabPage;
            rent_Brodcast_Car.TopLevel = false;
            rent_Brodcast_Car.Visible = true;
            rent_Brodcast_Car.FormBorderStyle = FormBorderStyle.None;
            rent_Brodcast_Car.Dock = DockStyle.Fill;
            chg_tabPage.Controls.Add(rent_Brodcast_Car);
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            contextMenuStrip_Report.Show(button3, new Point(0, button3.Height));
        }

        private void button8_Click(object sender, EventArgs e)
        {
            foreach (Form form in Application.OpenForms)
            {
                if (form.GetType() == typeof(AutnForm))
                {
                    form.Activate();
                    return;
                }
            }

            AutnForm autnForm = new AutnForm(role);
            tabControl1.Show();
            TabPage autnPage = new TabPage("Журнал АУТН");
            tabControl1.TabPages.Add(autnPage);
            tabControl1.SelectedTab = autnPage;
            autnForm.TopLevel = false;
            autnForm.Visible = true;
            autnForm.FormBorderStyle = FormBorderStyle.None;
            autnForm.Dock = DockStyle.Fill;
            autnPage.Controls.Add(autnForm);
        }

        private void дополнительныеПараметрыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string GetParameter = "select Parameter from d__Parameter";
            DataTable param = DbConnection.DBConnect(GetParameter);
            LastRent_Days last = new LastRent_Days();
            last.textEdit1.Text = param.Rows[0][0].ToString();
            last.ShowDialog();
        }

        private void добавитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AddGlobalFilterForm addGlobalFilterForm = new AddGlobalFilterForm(UserID);
            addGlobalFilterForm.Owner = this;
            addGlobalFilterForm.ShowDialog();
        }
    }
}
