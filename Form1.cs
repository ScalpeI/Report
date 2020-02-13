using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;

namespace Report_on_issued_policies
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        //Екземпляр приложения Excel
        Excel.Application xlApp;
        //Лист
        Excel.Worksheet xlSheet;
        //Выделеная область
        Excel.Range xlSheetRange;

        private DataTable GetData()
        {
            //строка соединения
            string connString = @"Network Library=DBMSSOCN;Data Source=192.168.1.101;Initial Catalog=Oms_Buryatiya;User ID=sa;Password=Kmsadmin_403;Persist Security Info=True;";

            SqlConnection con = new SqlConnection(connString);
            DataTable dt = new DataTable();
            try
            {
                string query =
                  @"declare @DR as datetime
                    declare @DB as datetime
                    declare @DE as datetime

                    set @DB = '" + dtpStart.Value + @"'
                    set @DR = '" + dtpNewborn.Value + @"'
                    set @DE = '" + dtpEnd.Value + @"'

                    Select
                        r.Name as Офис,
	                    p.IDPRZ as ПВП,
	                    sum(case when p.PolisType = '2' and s.StatementType = '1' and s.StatementWork = '1' then 1 else 0 end) as Выбор,
	                    sum(case when p.PolisType = '2' and s.StatementType = '2' and s.StatementWork = '1' then 1 else 0 end) as Переоформление,
	                    sum(case when p.PolisType = '2' and s.StatementType = '2' and s.StatementWork = '2' then 1 else 0 end) as Дубликат,
	                    sum(case when p.PolisType = '2' and (s.StatementType = '1') and (s.StatementWork = '4' or s.StatementWork = '2' or s.StatementWork = '3') and s.StatementBase != '0' then 1 else 0 end) as  С_переизготовлением,
	                    sum(case when p.PolisType = '3' and (s.StatementType = '1') and (s.StatementWork = '4' or s.StatementWork = '2' or s.StatementWork = '3') and s.StatementBase = '0' then 1 else 0 end) as Без_переизготовлением,
	                    sum(case when (s.Birthday>= @DR and s.Birthday< @DE) and p.PolisType = '2' and (s.StatementWork = '1' or s.StatementWork = '4' or s.StatementWork = '2' or s.StatementWork = '3') then 1 else 0 end) as Новорожденные,
	                    sum(case when (p.PolisType = '2' and s.StatementType = '1' and s.StatementWork = '1') or 
	                        (p.PolisType = '2' and (s.StatementType = '1') and (s.StatementWork = '4' or s.StatementWork = '2' or s.StatementWork = '3') and s.StatementBase != '0') or 
	                        (p.PolisType = '3' and (s.StatementType = '1') and (s.StatementWork = '4' or s.StatementWork = '2' or s.StatementWork = '3') and s.StatementBase = '0')	then 1 else 0 end) as ВСЕГО
	                from Polis as p
	                    left join
	                        Statement as s
	                    on
	                        p.IDPolis = s.IdPolis 
	                    left join
	                        PRZ as r
	                    on
	                        p.IDPRZ = r.IDPRZ
	                    left join
	                        PravaPLZ as w
	                    on
	                        p.IDOperator = w.PrUserID
	                where 
		                p.PolisDate >= @DB
			            and
			            p.PolisDate <= @DE

	                group by p.IDPRZ,r.Name
	                order by p.IDPRZ";

                SqlCommand comm = new SqlCommand(query, con);
                con.Open();
                SqlDataAdapter da = new SqlDataAdapter(comm);
                DataSet ds = new DataSet();
                da.Fill(ds);
                dt = ds.Tables[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                con.Close();
                con.Dispose();
            }
            return dt;
        }

        private DataTable GetSum()
        {
            //строка соединения
            string connString = @"Network Library=DBMSSOCN;Data Source=192.168.1.101;Initial Catalog=Oms_Buryatiya;User ID=sa;Password=Kmsadmin_403;Persist Security Info=True;";

            SqlConnection con = new SqlConnection(connString);
            DataTable dtsum = new DataTable();
            try
            {
                string query =
                  @"declare @DR as datetime
                    declare @DB as datetime
                    declare @DE as datetime

                    set @DB = '" + dtpStart.Value + @"'
                    set @DR = '" + dtpNewborn.Value + @"'
                    set @DE = '" + dtpEnd.Value + @"'

                    Select
                        'ИТОГО',
	                    '',
	                    sum(case when p.PolisType = '2' and s.StatementType = '1' and s.StatementWork = '1' then 1 else 0 end) as Выбор,
	                    sum(case when p.PolisType = '2' and s.StatementType = '2' and s.StatementWork = '1' then 1 else 0 end) as Переоформление,
	                    sum(case when p.PolisType = '2' and s.StatementType = '2' and s.StatementWork = '2' then 1 else 0 end) as Дубликат,
	                    sum(case when p.PolisType = '2' and (s.StatementType = '1') and (s.StatementWork = '4' or s.StatementWork = '2' or s.StatementWork = '3') and s.StatementBase != '0' then 1 else 0 end) as  С_переизготовлением,
	                    sum(case when p.PolisType = '3' and (s.StatementType = '1') and (s.StatementWork = '4' or s.StatementWork = '2' or s.StatementWork = '3') and s.StatementBase = '0' then 1 else 0 end) as Без_переизготовлением,
	                    sum(case when (s.Birthday>= @DR and s.Birthday< @DE) and p.PolisType = '2' and (s.StatementWork = '1' or s.StatementWork = '4' or s.StatementWork = '2' or s.StatementWork = '3') then 1 else 0 end) as Новорожденные,
	                    sum(case when (p.PolisType = '2' and s.StatementType = '1' and s.StatementWork = '1') or 
	                        (p.PolisType = '2' and (s.StatementType = '1') and (s.StatementWork = '4' or s.StatementWork = '2' or s.StatementWork = '3') and s.StatementBase != '0') or 
	                        (p.PolisType = '3' and (s.StatementType = '1') and (s.StatementWork = '4' or s.StatementWork = '2' or s.StatementWork = '3') and s.StatementBase = '0')	then 1 else 0 end) as ВСЕГО
	                from Polis as p
	                    left join
	                        Statement as s
	                    on
	                        p.IDPolis = s.IdPolis 
	                    left join
	                        PRZ as r
	                    on
	                        p.IDPRZ = r.IDPRZ
	                    left join
	                        PravaPLZ as w
	                    on
	                        p.IDOperator = w.PrUserID
	                where 
		                p.PolisDate >= @DB
			            and
			            p.PolisDate <= @DE";

                SqlCommand comm = new SqlCommand(query, con);
                con.Open();
                SqlDataAdapter da = new SqlDataAdapter(comm);
                DataSet ds = new DataSet();
                da.Fill(ds);
                dtsum = ds.Tables[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                con.Close();
                con.Dispose();
            }
            return dtsum;
        }

        void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show(ex.ToString(), "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            finally
            {
                GC.Collect();
            }
        }

        private void btnCreate_Click(object sender, EventArgs e)
        {
            xlApp = new Excel.Application();

            try
            {
                //добавляем книгу
                xlApp.Workbooks.Add(Type.Missing);

                //делаем временно неактивным документ
                xlApp.Interactive = false;
                xlApp.EnableEvents = false;

                //выбираем лист на котором будем работать (Лист 1)
                xlSheet = (Excel.Worksheet)xlApp.Sheets[1];
                //Название листа
                xlSheet.Name = "Данные";

                //Выгрузка данных
                DataTable dt = GetData();
                DataTable dtsum = GetSum();

                int collInd = 0;
                int rowInd = 0;
                string data = "";

                //называем колонки
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    data = dt.Columns[i].ColumnName.ToString();
                    xlSheet.Cells[1, i + 1] = data;

                    //выделяем первую строку
                    xlSheetRange = xlSheet.get_Range("A1:Z1", Type.Missing);

                    //делаем полужирный текст и перенос слов
                    xlSheetRange.WrapText = true;
                    xlSheetRange.Font.Bold = true;
                }

                //заполняем строки
                for (rowInd = 0; rowInd < dt.Rows.Count; rowInd++)
                {
                    for (collInd = 0; collInd < dt.Columns.Count; collInd++)
                    {
                        data = dt.Rows[rowInd].ItemArray[collInd].ToString();
                        xlSheet.Cells[rowInd + 2, collInd + 1] = data;
                    }
                }
                for (rowInd = 0; rowInd < dtsum.Rows.Count; rowInd++)
                {
                    for (collInd = 0; collInd < dtsum.Columns.Count; collInd++)
                    {
                        data = dtsum.Rows[rowInd].ItemArray[collInd].ToString();
                        xlSheet.Cells[rowInd + dt.Rows.Count + 2, collInd +  1] = data;
                    }
                }

                //выбираем всю область данных
                xlSheetRange = xlSheet.UsedRange;

                //выравниваем строки и колонки по их содержимому
                xlSheetRange.Columns.AutoFit();
                xlSheetRange.Rows.AutoFit();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                //Показываем ексель
                xlApp.Visible = true;

                xlApp.Interactive = true;
                xlApp.ScreenUpdating = true;
                xlApp.UserControl = true;

                //Отсоединяемся от Excel
                releaseObject(xlSheetRange);
                releaseObject(xlSheet);
                releaseObject(xlApp);
            }
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            var year = DateTime.Now.Year;
            var month = DateTime.Now.Month;
            dtpStart.Value = new DateTime(year, month, 1);
            dtpEnd.Value = new DateTime(year, month, 31);
            dtpNewborn.Value = new DateTime(year, 1, 1);
        }
    }
}
