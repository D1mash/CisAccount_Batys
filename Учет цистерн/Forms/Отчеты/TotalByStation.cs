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
    class TotalByStation
    {
        public void TBS(string date_1, string date_2, System.Data.DataTable dt)
        {
            var GR = new General_Reestr();

            string path = AppDomain.CurrentDomain.BaseDirectory + @"ReportTemplates\Итог по станции.xlsx";

            using (SLDocument sl = new SLDocument(path))
            {
                sl.SelectWorksheet("Итоговая  справка");

                sl.SetCellValue("B6", "c " + date_1 + " по " + date_2);

                var val = dt.Rows.Count * 2 + 11;
                sl.CopyCell("B13", "H24", "B" + val, true);
                //worksheet.Range["B13:H24"].Cut(worksheet.Cells[gridView1.RowCount * 2 + 11, 2]);

                int item = 0;

                //Кол.во услуг
                int total = 0;

                //Сумм * Кол.во
                double final_sum = 0;

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (i % 2 == 0)
                    {
                        var k = 11 + item;
                        //Excel.Range range = worksheet.Range[worksheet.Cells[i + k, 2], worksheet.Cells[i + k, 8]];
                        //range.Merge();

                        sl.MergeWorksheetCells(i + k, 2, i + k, 8);

                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            if (j == 0)
                            {
                                sl.SetCellValue(i + k, j + 2, dt.Rows[i][j].ToString());
                                sl.SetCellStyle(i + k, j + 2, GR.FormattingExcelCells(sl, false));
                            }
                            else
                            {
                                sl.SetCellValue(i + k, j + 8, Convert.ToDecimal(dt.Rows[i][j].ToString()));
                                sl.SetCellStyle(i + k, j + 8, GR.FormattingExcelCells(sl, false));
                            }
                        }
                    }
                    else
                    {
                        var k = 11 + item;
                        //Excel.Range range = worksheet.Range[worksheet.Cells[i + k, 2], worksheet.Cells[i + k, 8]];
                        //range.Merge();

                        sl.MergeWorksheetCells(i + k, 2, i + k, 8);

                        //FormattingExcelCells(range, false, false);
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            if (j == 0)
                            {
                                sl.SetCellValue(i + k, j + 2, dt.Rows[i][j].ToString());
                                sl.SetCellStyle(i + k, j + 2, GR.FormattingExcelCells(sl, false));
                            }
                            else
                            {
                                sl.SetCellValue(i + k, j + 8, Convert.ToDecimal(dt.Rows[i][j].ToString()));
                                sl.SetCellStyle(i + k, j + 8, GR.FormattingExcelCells(sl, false));
                            }
                        }

                    }

                    final_sum += int.Parse(dt.Rows[i][1].ToString()) * double.Parse(dt.Rows[i][2].ToString()); ;

                    if (i < dt.Rows.Count && dt.Rows[i][0].ToString() != "Текущий отцепочный ремонт горячей обработкой" && dt.Rows[i][0].ToString() != "Текущий отцепочный ремонт")
                    {
                        total += int.Parse(dt.Rows[i][1].ToString());
                    }
                    else
                    {
                        continue;
                    }

                    //backgroundWorker.ReportProgress(i);
                    item++;
                }
                //Кол.во обработанных
                sl.SetCellValue("I8", total);

                //Итоговая сумма
                sl.SetCellValue(dt.Rows.Count + 12 + item, 10, final_sum);
                sl.SetCellStyle(dt.Rows.Count + 12 + item, 10, GR.FormattingExcelCells(sl, false));

                sl.SaveAs(AppDomain.CurrentDomain.BaseDirectory + @"Report\Итог по станции.xlsx");
            }

            dt.Clear();

            Process.Start(AppDomain.CurrentDomain.BaseDirectory + @"Report\Итог по станции.xlsx");
        }


        public void TotalByCompany(string date_1, string date_2, object comapany, string company_name, string dates_1, string dates_2)
        {
            var GR = new General_Reestr();

            DataTable dataTable, Itog_Rep;

            string path = AppDomain.CurrentDomain.BaseDirectory + @"ReportTemplates\Реестр  за арендованных и  собственных вагон-цистерн компании.xlsx";

            using (SLDocument sl = new SLDocument(path))
            {
                sl.SelectWorksheet("ТОО Казыкурт");
                sl.RenameWorksheet("ТОО Казыкурт", "Реестр");

                string Refresh = "dbo.GetReportRenderedServices_v1 '" + date_1 + "','" + date_2 + "','" + comapany + "'";
                dataTable = DbConnection.DBConnect(Refresh);

                string GetCountServiceCost = "exec dbo.Itog_Report  '" + date_1 + "','" + date_2 + "','" + comapany + "'";
                Itog_Rep = DbConnection.DBConnect(GetCountServiceCost);

                sl.SetCellValue("C4", company_name);

                sl.SetCellValue("C6", "в ТОО Казыгурт-Юг c " + dates_1 + " по " + dates_2);

                var val = dataTable.Rows.Count + 16 + Itog_Rep.Rows.Count * 2;
                sl.CopyCell("B13", "K20", "B" + val, true);

                sl.ImportDataTable(10, 1, dataTable, false);
                sl.CopyCell("E10", "R" + Convert.ToString(dataTable.Rows.Count + 10), "G10", true);

                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    for (int j = 1; j < dataTable.Columns.Count + 3; j++)
                    {
                        sl.SetCellStyle(i + 10, j, GR.FormattingExcelCells(sl, true));
                    }
                }

                sl.SetCellValue(dataTable.Rows.Count + 12, 2, "=C6");
                sl.SetCellStyle(dataTable.Rows.Count + 12, 2, GR.FormattingExcelCells(sl, false));

                sl.SetCellValue(dataTable.Rows.Count + 14, 2, "Всего обработано вагонов - цистерн " + company_name + " по видам операций:");
                sl.SetCellStyle(dataTable.Rows.Count + 14, 2, GR.FormattingExcelCells(sl, false));

                //Итоговая сводка
                int rowcount = 0;
                int total = 0;
                double EndSum = 0;

                for (int i = 0; i < Itog_Rep.Rows.Count; i++)
                {
                    rowcount++;
                    double val1 = 1;
                    for (int j = 0; j < Itog_Rep.Columns.Count; j++)
                    {
                        if (j == 0)
                        {
                            sl.SetCellValue(i + dataTable.Rows.Count + 15 + rowcount, j + 2, Itog_Rep.Rows[i][j].ToString());
                            sl.SetCellStyle(i + dataTable.Rows.Count + 15 + rowcount, j + 2, GR.FormattingExcelCells(sl, false));
                        }
                        else
                        {
                            val1 = val1 * double.Parse(Itog_Rep.Rows[i][j].ToString());
                            sl.SetCellValue(i + dataTable.Rows.Count + 15 + rowcount, j + 12, Convert.ToDecimal(Itog_Rep.Rows[i][j].ToString()));
                            sl.SetCellStyle(i + dataTable.Rows.Count + 15 + rowcount, j + 12, GR.FormattingExcelCells(sl, false));
                        }
                    }
                    EndSum += val1;
                    if (i < Itog_Rep.Rows.Count && Itog_Rep.Rows[i][0].ToString() != "Текущий отцепочный ремонт горячей обработкой" && Itog_Rep.Rows[i][0].ToString() != "Текущий отцепочный ремонт")
                    {
                        total += int.Parse(Itog_Rep.Rows[i][1].ToString());
                    }
                    else
                    {
                        continue;
                    }
                }

                sl.SetCellValue(dataTable.Rows.Count + 14, 13, total);
                sl.SetCellStyle(dataTable.Rows.Count + 14, 13, GR.FormattingExcelCells(sl, false));


                sl.SetCellValue(dataTable.Rows.Count + 14, 13, dataTable.Rows.Count);

                //Итоговая сумма
                sl.SetCellValue(dataTable.Rows.Count + Itog_Rep.Rows.Count * 2 + 16, 14, EndSum);
                sl.SetCellStyle(dataTable.Rows.Count + Itog_Rep.Rows.Count * 2 + 16, 14, GR.FormattingExcelCells(sl, false));

                sl.SaveAs(AppDomain.CurrentDomain.BaseDirectory + @"Report\" + company_name + " Реестр  за арендованных и  собственных вагон-цистерн компании.xlsx");
            }

            dataTable.Clear();
            Itog_Rep.Clear();
            Process.Start(AppDomain.CurrentDomain.BaseDirectory + @"Report\" + company_name + " Реестр  за арендованных и  собственных вагон-цистерн компании.xlsx");
        }
    }
}
