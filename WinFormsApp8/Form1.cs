using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace WinFormsApp8
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OleDbConnection excel_con = null;
            string xls_filename;

            try
            {
                //엑셀 불러오기
                FileDialog file_dlg = new OpenFileDialog();
                file_dlg.InitialDirectory = "c\\";
                file_dlg.Filter = "모든 파일 (*.*)|*.*";
                file_dlg.RestoreDirectory = true;

                //엑셀문서를 보여주기
                if (file_dlg.ShowDialog() == DialogResult.OK)
                {
                    xls_filename = file_dlg.FileName;

                    string str_con = "Provider = Microsoft.ACE.OLEDB.12.0.0;Data Source=" + xls_filename + ";Extended Properties='Excel 12.0;HDR=YES'";
                    excel_con = new OleDbConnection(str_con);

                    excel_con.Open();
                    string excel_sql = @"select * from[Sheet1$]";

                    OleDbDataAdapter excel_adapter = new OleDbDataAdapter(excel_sql, excel_con);
                    DataSet excel_DS = new DataSet();
                    excel_adapter.Fill(excel_DS);

                    DataTable excel_table = excel_DS.Tables[0];

                    dataGridView1.DataSource = excel_table;
                    lbl_flename.Text = xls_filename;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("파일 가져오기 실패 : " + ex.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (lbl_flename.Text == "filename")
            {
                MessageBox.Show("엑셀파일을 불러와 주세요.");
            }
            else if (txt_table_name.Text == "")
            {
                MessageBox.Show("저장할 테이블 명을 입력해 주세요. ");
            }
            else
            {
                OleDbConnection excel_con = null;

                string server = txt_server.Text.ToString().Trim();
                string db = txt_db.Text.ToString().Trim();
                string id = txt_id.Text.ToString().Trim();
                string pw = txt_pw.Text.ToString().Trim();
                //string filename = lbl_flename.Text;

                string xls_filename = lbl_flename.Text;
                try
                {
                    string str_con = "Provider = Microsoft.ACE.OLEDB.12.0.0;Data Source=" + xls_filename + ";Extended Properties='Excel 12.0;HDR=YES'";
                    excel_con = new OleDbConnection(str_con);

                    excel_con.Open();
                    string excel_sql = @"select * from[Sheet1$]";

                    OleDbDataAdapter excel_adapter = new OleDbDataAdapter(excel_sql, excel_con);
                    DataSet excel_DS = new DataSet();
                    excel_adapter.Fill(excel_DS);

                    DataTable excel_table = excel_DS.Tables[0];

                    //sql에 저장하기
                    string cnString = "Data Source=" + server + ";Initial Catalog=" + db + ";Persist Security Info=True;User ID=" + id + ";Password=" + pw;
                    SqlConnection con = new SqlConnection(cnString);

                    con.Open();
                    SqlCommand SqlCommand = new SqlCommand();
                    SqlCommand.Connection = con;

                    //table 유무확인
                    SqlCommand.CommandText = "SELECT COUNT(*) FROM information_schema.tables where table_name = '" + txt_table_name.Text + "' ";
                    string tbl_exist = SqlCommand.ExecuteScalar().ToString();

                    //이미 DB에 table이 있을경우
                    if (tbl_exist == "1")
                    {
                        MessageBox.Show("디비에 업로드 중입니다. ");
                        SqlCommand.CommandText = "update [uniLITE].[BPR200T] SET SAFE_STOCK_Q = 0";
                        SqlCommand.ExecuteNonQuery();
                    }

                    for(int i = 0; i < dataGridView1.RowCount; i++)
                    {
                        SqlCommand.CommandText = "update [uniLITE].[BPR200T] SET SAFE_STOCK_Q = " + dataGridView1.Rows[i].Cells[1].Value.ToString() + " WHERE ITEM_CODE = '" + dataGridView1.Rows[i].Cells[0].Value.ToString() + "'";
                        SqlCommand.ExecuteNonQuery();
                        MessageBox.Show("업로드가 완료되었습니다..");
                    }

                    ////DB에 table이 없는 경우  data_table의 컬럼값을 읽어와 sql_table생성
                    //else
                    //{
                    //    int column_count = excel_table.Columns.Count;//칼럼명
                    //    string[] sql_column = new string[column_count];
                    //    string sql_create = "create table [" + txt_table_name.Text.Trim() + "] ( ";
                    //    for (int i = 0; i <= column_count - 1; i++)
                    //    {
                    //        string column_name = excel_table.Columns[i].ColumnName.ToString();
                    //        sql_create = sql_create + "[" + column_name + "] varchar(50)";
                    //        if (i != column_count - 1)
                    //        { sql_create = sql_create + ","; }
                    //    }
                    //    sql_create = sql_create + ")";
                    //    SqlCommand.CommandText = sql_create;
                    //    SqlCommand.ExecuteNonQuery();

                    //    //레코드 불러와서 insert
                    //    if (excel_table.Rows.Count > 0)
                    //    {
                    //        //레코드 갯수 불러와서 그 수만큼 반복
                    //        for (int j = 0; j <= excel_table.Rows.Count - 1; j++)
                    //        {
                    //            string sql_insert = " insert [" + txt_table_name.Text.Trim() + "] (";
                    //            //data_talbe 컬럼값가져와 sql 필드명으로 대입
                    //            for (int i = 0; i <= column_count - 1; i++)
                    //            {
                    //                string column_name = excel_table.Columns[i].ColumnName.ToString();
                    //                sql_insert = sql_insert + "[" + column_name + "]";
                    //                if (i != column_count - 1)
                    //                { sql_insert = sql_insert + ","; }
                    //            }
                    //            sql_insert = sql_insert + ") values (";
                    //            //data_table에서 입력할 필드값가져오기
                    //            for (int i = 0; i <= column_count - 1; i++)
                    //            {
                    //                string value = Convert.ToString(excel_table.Rows[j][i]);
                    //                sql_insert = sql_insert + "'" + value + "'";
                    //                if (i != column_count - 1)
                    //                { sql_insert = sql_insert + ","; }
                    //            }
                    //            sql_insert = sql_insert + ")";
                    //            SqlCommand.CommandText = sql_insert;
                    //            SqlCommand.ExecuteNonQuery();
                    //        }
                    //    }
                    //    MessageBox.Show("업로드 완료되었습니다");
                    //}

                }
                catch (Exception ex)
                {
                    MessageBox.Show("파일 저장에 실패하였습니다. :" + ex.Message);
                }
            }

        }
    }
}

