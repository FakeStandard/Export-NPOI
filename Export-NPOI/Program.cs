using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using NPOI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;

namespace Export_NPOI
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("使用 NPOI 匯出");

            // 取得 Orders 資料
            DataTable dt = Connect();

            string path = @"D:\Temp.xlsx";

            // 匯出 Excel
            SimpleExport(dt, path);

            Console.WriteLine("匯出成功");

            Console.ReadLine();
        }

        /// <summary>
        /// 匯出 Excel 簡易方法
        /// </summary>
        /// <param name="dt">要匯出的資料</param>
        /// <param name="path">要匯出的目的地路徑</param>
        static void SimpleExport(DataTable dt, string path)
        {
            // 建立工作簿
            IWorkbook wb = new XSSFWorkbook();

            // .xls
            // IWorkbook wb = new HSSFWorkbook();

            using (FileStream fs = File.Create(path))
            {
                // 建立名為 Simple 的工作表
                ISheet sheet = wb.CreateSheet("Simple");

                int i = 0;
                int j = 0;

                #region Title

                // 建立新的 row 給標題用
                sheet.CreateRow(0);

                for (i = 0; i < dt.Columns.Count; i++)
                {
                    // 取得已建立的 row ，再建立 cell，最後再配置值給 cell
                    sheet.GetRow(0).CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
                }

                #endregion

                #region Content

                // 走訪 dt
                for (i = 1; i <= dt.Rows.Count; i++)
                {                 
                    sheet.CreateRow(i);

                    for (j = 0; j < dt.Columns.Count; j++)
                    {
                        // 因為建立過的 row 或 cell 再次建立會覆蓋原有的值，所以建立過的物件使用 Get 取得再設置 Value
                        sheet.GetRow(i).CreateCell(j).SetCellValue(dt.Rows[i-1][j].ToString());
                    }
                }
                #endregion

                // 將 wb 寫出到檔案流
                wb.Write(fs);

                // 釋放資源
                wb.Close();
                wb = null;
                sheet = null;               
            }
        }

        /// <summary>
        /// 資料庫連線方法
        /// </summary>
        /// <returns>返回 DataTable 物件</returns>
        static DataTable Connect()
        {
            SqlConnectionStringBuilder cnsb = new SqlConnectionStringBuilder();
            cnsb.DataSource = ".";
            cnsb.InitialCatalog = "Northwind";
            cnsb.IntegratedSecurity = true;

            SqlConnection cn = new SqlConnection(cnsb.ConnectionString);

            string sql = "SELECT * FROM [Northwind].[dbo].[Orders]";
            SqlDataAdapter da = new SqlDataAdapter(sql, cn);
            DataSet ds = new DataSet();
            da.Fill(ds);

            DataTable dt = ds.Tables[0];

            return dt;
        }
    }
}
