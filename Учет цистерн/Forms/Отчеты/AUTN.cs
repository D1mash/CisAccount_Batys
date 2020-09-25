using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Учет_цистерн.Forms.Отчеты
{
    class AUTN
    {
        public void AUTN_Report(System.Data.DataTable dt)
        {
            var GR = new General_Reestr();

            string path = AppDomain.CurrentDomain.BaseDirectory + @"ReportTemplates\Реестр АУТН.xlsx";

            using (SLDocument sl = new SLDocument(path))
            {
                sl.SelectWorksheet("АУТН");

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        sl.SetCellValue(i + 3, 1, i);
                        sl.SetCellStyle(i + 3, 1, GR.FormattingExcelCells(sl, true));
                        if (j == 1)
                        {
                            sl.SetCellValue(i + 3, j + 2, Convert.ToInt32(dt.Rows[i][j].ToString()));
                            sl.SetCellStyle(i + 3, j + 2, GR.FormattingExcelCells(sl, true));
                        }
                        else
                        {
                            sl.SetCellValue(i + 3, j + 2, dt.Rows[i][j].ToString());
                            sl.SetCellStyle(i + 3, j + 2, GR.FormattingExcelCells(sl, true));
                        }
                    }

                    //backgroundWorker.ReportProgress(i);
                }
                sl.AutoFitColumn(1, 9);


                sl.SaveAs(AppDomain.CurrentDomain.BaseDirectory + @"Report\Реестр АУТН.xlsx");

            }

            dt.Clear();

            Process.Start(AppDomain.CurrentDomain.BaseDirectory + @"Report\Реестр АУТН.xlsx");
        }
    }
}
