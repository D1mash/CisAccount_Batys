using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Учет_цистерн.Forms.Отчеты
{
    class SNO
    {
        public void SNO_IN(string date_1, string date_2, System.Data.DataTable dt)
        {
            var GR = new General_Reestr();

            string path = AppDomain.CurrentDomain.BaseDirectory + @"ReportTemplates\СНО Приход.xlsx";

            using (SLDocument sl = new SLDocument(path))
            {
                sl.SelectWorksheet("СНО Приход");

                sl.SetCellValue("B1", "Приход СНО в ТОО Казыгурт-Юг за период с " + date_1 + " по " + date_2);

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int j = 1; j < dt.Columns.Count; j++)
                    {
                        if (j == 1)
                        {
                            sl.SetCellValue(i + 4, j, Convert.ToDecimal(dt.Rows[i][j].ToString()));
                            sl.SetCellStyle(i + 4, j, GR.FormattingExcelCells(sl, true));
                        }
                        else
                        {
                            sl.SetCellStyle(i + 4, j + 1, GR.FormattingExcelCells(sl, true));
                            if (j > 1 && j < 4)
                            {
                                sl.SetCellValue(i + 4, j + 1, Convert.ToDecimal(dt.Rows[i][j].ToString()));
                            }
                            else
                            {
                                sl.SetCellValue(i + 4, j + 1, dt.Rows[i][j].ToString());
                            }
                        }
                    }
                    sl.SetCellValue(i + 4, 2, $"=C{i + 4} + D{i + 4}");
                    sl.SetCellStyle(i + 4, 2, GR.FormattingExcelCells(sl, true));
                    //backgroundWorker1.ReportProgress(i);
                }

                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    sl.SetCellStyle(dt.Rows.Count + 4, j + 1, GR.FormattingExcelCells(sl, true));
                }

                sl.SetCellValue(dt.Rows.Count + 4, 1, "Итог");

                sl.SetCellValue(dt.Rows.Count + 4, 2, $"=SUM(B{4}:C{dt.Rows.Count + 3})");

                sl.SetCellValue(dt.Rows.Count + 4, 3, $"=SUM(C{4}:C{dt.Rows.Count + 3})");

                sl.SetCellValue(dt.Rows.Count + 4, 4, $"=SUM(D{4}:D{dt.Rows.Count + 3})");

                sl.SaveAs(AppDomain.CurrentDomain.BaseDirectory + @"Report\СНО Приход.xlsx");
            }

            dt.Clear();
            Process.Start(AppDomain.CurrentDomain.BaseDirectory + @"Report\СНО Приход.xlsx");
        }

        public void SNO_OUT(string date_1, string date_2, System.Data.DataTable dt)
        {
            var GR = new General_Reestr();
            string path = AppDomain.CurrentDomain.BaseDirectory + @"ReportTemplates\СНО Реализ.xlsx";

            using (SLDocument sl = new SLDocument(path))
            {
                sl.SelectWorksheet("СНО Реализация");

                sl.SetCellValue("B3", "в ТОО Казыгурт-Юг реализация СНО за период с " + date_1 + " по " + date_2);

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int j = 2; j < dt.Columns.Count; j++)
                    {
                        if (j >= 2 && j <= 3)
                        {
                            sl.SetCellValue(i + 6, j - 1, dt.Rows[i][j].ToString());
                            sl.SetCellStyle(i + 6, j - 1, GR.FormattingExcelCells(sl, true));
                        }
                        else
                        {
                            if (j > 3 && j <= 6)
                            {
                                sl.SetCellValue(i + 6, j - 1, Convert.ToDecimal(dt.Rows[i][j].ToString()));
                                sl.SetCellStyle(i + 6, j - 1, GR.FormattingExcelCells(sl, true));
                            }
                            else
                            if (j == 8)
                            {
                                sl.SetCellValue(i + 6, j - 2, Convert.ToDecimal(dt.Rows[i][j].ToString()));
                                sl.SetCellStyle(i + 6, j - 2, GR.FormattingExcelCells(sl, true));
                            }
                            else
                            if (j == 9)
                            {
                                sl.SetCellValue(i + 6, j - 2, dt.Rows[i][j].ToString());
                                sl.SetCellStyle(i + 6, j - 2, GR.FormattingExcelCells(sl, true));
                            }
                        }
                    }

                    //backgroundWorker1.ReportProgress(i);
                }

                for (int j = 0; j < dt.Columns.Count - 2; j++)
                {
                    sl.SetCellStyle(dt.Rows.Count + 6, j, GR.FormattingExcelCells(sl, true));
                }

                sl.SetCellValue(dt.Rows.Count + 6, 1, "Итог");

                sl.SetCellValue(dt.Rows.Count + 6, 3, $"=SUM(C{6}:C{dt.Rows.Count + 5})");

                sl.SetCellValue(dt.Rows.Count + 6, 5, $"=SUM(E{6}:E{dt.Rows.Count + 5})");

                sl.SetCellValue(dt.Rows.Count + 6, 6, $"=SUM(F{6}:F{dt.Rows.Count + 5})");

                sl.SaveAs(AppDomain.CurrentDomain.BaseDirectory + @"Report\СНО Реализация.xlsx");
            }

            dt.Clear();
            Process.Start(AppDomain.CurrentDomain.BaseDirectory + @"Report\СНО Реализация.xlsx");
        }
    }
}
