using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Учет_цистерн.Forms.Отчеты
{
    class General_Reestr
    {
        public void General_Reesters(string v1, string v2)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + @"ReportTemplates\Реестр  за арендованных и  собственных вагон-цистерн компании.xlsx";

            DataTable dt, dataTable, Itog_Rep;
            string Name;

            using (SLDocument sl = new SLDocument(path))
            {

                sl.SelectWorksheet("Batys");
                sl.RenameWorksheet("Batys", "Реестр");
                sl.AddWorksheet("Temp");

                string RefreshAllCount = "exec dbo.GetReportAllRenderedService_v1 '" + v1 + "','" + v2 + "'," + "@Type = " + 2;
                dt = DbConnection.DBConnect(RefreshAllCount);

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (j == 1)
                        {
                            Name = dt.Rows[i][j].ToString();
                            sl.CopyWorksheet("Реестр", Name);
                        }
                    }
                }

                sl.SelectWorksheet("Реестр");
                sl.DeleteWorksheet("Temp");

                int k = 0;

                foreach (var name in sl.GetWorksheetNames())
                {
                    if (name == "Реестр")
                    {
                        string RefreshAll = "exec dbo.GetReportAllRenderedService_v1 '" + v1 + "','" + v2 + "'," + "@Type = " + 1;
                        dataTable = DbConnection.DBConnect(RefreshAll);

                        string GetCountServiceCost = "exec dbo.Itog_All_Report '" + v1 + "','" + v2 + "'";
                        Itog_Rep = DbConnection.DBConnect(GetCountServiceCost);

                        sl.SelectWorksheet(name);
                        sl.SetCellValue("F12", "в ТОО \"Batys Petroleum\" " + v1 + " по " + v2);

                        var val = dataTable.Rows.Count + 18;
                        sl.CopyCell("B18", "G24", "B" + val, true);

                        sl.ImportDataTable(16, 1, dataTable, false);
                        sl.CopyCell("W16", "W" + Convert.ToString(dataTable.Rows.Count + 16), "X16", true);

                        for (int i = 0; i < dataTable.Rows.Count; i++)
                        {
                            for (int j = 1; j <= dataTable.Columns.Count+1; j++)
                            {
                                sl.SetCellStyle(i + 16, j, FormattingExcelCells(sl, true));
                            }
                        }

                        double EndSum = 0;
                        int rowcount = 0;
                        for (int i = 0; i < Itog_Rep.Rows.Count; i++)
                        {
                            rowcount++;
                            double val1 = 1;
                            for (int j = 0; j < Itog_Rep.Columns.Count; j++)
                            {
                                if (j == 0)
                                {
                                    sl.SetCellValue(i + dataTable.Rows.Count + 19 + rowcount, j + 2, Itog_Rep.Rows[i][j].ToString());
                                    sl.SetCellStyle(i + dataTable.Rows.Count + 19 + rowcount, j + 2, FormattingExcelCells(sl, false));
                                }
                                else
                                {
                                    val1 = val1 * double.Parse(Itog_Rep.Rows[i][j].ToString());
                                    sl.SetCellValue(i + dataTable.Rows.Count + 19 + rowcount, j + 10, Convert.ToDecimal(Itog_Rep.Rows[i][j].ToString()));
                                    sl.SetCellStyle(i + dataTable.Rows.Count + 19 + rowcount, j + 10, FormattingExcelCells(sl, false));
                                }
                            }
                            EndSum += val1;
                        }

                        //Итоговая сумма
                        sl.SetCellValue(val, 4, EndSum);
                        sl.SetCellStyle(val, 4, FormattingExcelCells(sl, false));

                        dataTable.Clear();
                        EndSum = 0;
                    }
                    else
                    {
                        int cs = Convert.ToInt32(dt.Rows[k][0].ToString());
                        string Refresh = "dbo.GetReportRenderedServices_v1 '" + v1 + "','" + v2 + "','" + cs + "'";
                        dataTable = DbConnection.DBConnect(Refresh);

                        string GetCountServiceCost = "exec dbo.Itog_Report  '" + v1 + "','" + v2 + "','" + cs + "'";
                        Itog_Rep = DbConnection.DBConnect(GetCountServiceCost);

                        Name = dt.Rows[k][1].ToString();
                        sl.SelectWorksheet(Name);

                        sl.SetCellValue("F10", dt.Rows[k][1].ToString());

                        sl.SetCellValue("F12", "в ТОО \"Batys Petroleum\"" + v1 + " по " + v2);

                        var val = dataTable.Rows.Count + 18;
                        sl.CopyCell("B18", "G24", "B" + val, true);

                        sl.ImportDataTable(16, 1, dataTable, false);
                        sl.CopyCell("W16", "W" + Convert.ToString(dataTable.Rows.Count + 16), "X16", true);

                        for (int i = 0; i < dataTable.Rows.Count; i++)
                        {
                            for (int j = 1; j <= dataTable.Columns.Count + 1; j++)
                            {
                                sl.SetCellStyle(i + 16, j, FormattingExcelCells(sl, true));
                            }
                        }

                        double EndSum = 0;
                        int rowcount = 0;
                        for (int i = 0; i < Itog_Rep.Rows.Count; i++)
                        {
                            rowcount++;
                            double val1 = 1;
                            for (int j = 0; j < Itog_Rep.Columns.Count; j++)
                            {
                                if (j == 0)
                                {
                                    sl.SetCellValue(i + dataTable.Rows.Count + 19 + rowcount, j + 2, Itog_Rep.Rows[i][j].ToString());
                                    sl.SetCellStyle(i + dataTable.Rows.Count + 19 + rowcount, j + 2, FormattingExcelCells(sl, false));
                                }
                                else
                                {
                                    val1 = val1 * double.Parse(Itog_Rep.Rows[i][j].ToString());
                                    sl.SetCellValue(i + dataTable.Rows.Count + 19 + rowcount, j + 10, Convert.ToDecimal(Itog_Rep.Rows[i][j].ToString()));
                                    sl.SetCellStyle(i + dataTable.Rows.Count + 19 + rowcount, j + 10, FormattingExcelCells(sl, false));
                                }
                            }
                            EndSum += val1;
                        }

                        //Итоговая сумма
                        sl.SetCellValue(val, 4, EndSum); 
                        sl.SetCellStyle(val, 4, FormattingExcelCells(sl, false));

                        k++;
                    }
                }

                sl.SaveAs(AppDomain.CurrentDomain.BaseDirectory + @"Report\Общий Реестр  за арендованных и  собственных вагон-цистерн компании.xlsx");
            }
            Process.Start(AppDomain.CurrentDomain.BaseDirectory + @"Report\Общий Реестр  за арендованных и  собственных вагон-цистерн компании.xlsx");
        }


        public SLStyle FormattingExcelCells(SLDocument sl, bool val)
        {
            SLStyle style1 = sl.CreateStyle();

            if (val == true) 
            {
                style1.SetBottomBorder(BorderStyleValues.Thin, System.Drawing.Color.Black);
                style1.SetTopBorder(BorderStyleValues.Thin, System.Drawing.Color.Black);
                style1.SetLeftBorder(BorderStyleValues.Thin, System.Drawing.Color.Black);
                style1.SetRightBorder(BorderStyleValues.Thin, System.Drawing.Color.Black);
                style1.Font.FontName = "Arial";
                style1.Font.FontSize = 9;
                style1.Font.Bold = true;
                style1.Alignment.Horizontal = HorizontalAlignmentValues.Center;

                return style1;
            }
            else
            {
                SLStyle style2 = sl.CreateStyle();
                style2.Font.FontName = "Arial Cyr";
                style2.Font.FontSize = 9;
                style2.Font.Bold = true;

                return style2;
            }
        }
    }
}
