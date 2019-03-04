using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Data;
using System.Data.OleDb;


namespace MDBDataHandle
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileName1=null;
            string fileName2=null;
            string fileName3 = null;
            string fileName4 = null;
            string fileName5 = null;






        }



      public   static DataSet MDBtoDT(string fileName,string tableName)
        {
            DataSet ds = new DataSet();

           string sql = "select * from " + tableName;
           string _connectionString;  
           OleDbConnection _odcConnection;  

             _connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";";  
          
            try  
               {  
                    // 建立连接  
                   _odcConnection = new OleDbConnection(_connectionString);  
 
  
                    // 打开连接  
                   _odcConnection.Open();  
                }  
               catch (Exception)  
                {  
                    throw new Exception("尝试打开 " + fileName + " 失败，确认文件是否存在！");  
                }  


              try  
               {  
                   OleDbDataAdapter adapter = new OleDbDataAdapter(sql, _odcConnection);
                   adapter.Fill(ds,tableName);  
               }                 
              catch (Exception)  
                {  
                    throw new Exception("sql语句： " + sql + " 执行失败！");  
                }  
  
            
            return ds;
        
        }

   
      public   static void DTtoMDB(DataSet ds, string fileName,string tableName)
       {
        
        string _connectionString;
        OleDbConnection _odcConnection;
     
        _connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";";
    

         string[] fields= GetTableColumn(fileName, tableName);


        try
        {
            // 建立连接  
            _odcConnection = new OleDbConnection(_connectionString);

            // 打开连接  
            _odcConnection.Open();

        }
        catch (Exception)
        {
            throw new Exception("尝试打开 " + fileName + " 失败，确认文件是否存在！");
        }


        try
        {
            //string strSQL = "INSERT INTO user (names,pwd) VALUES（'" + textBox1.Text + " ',' " + textBox2.Text + "'）";

          


            foreach (DataTable dt in ds.Tables)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    string sql = " INSERT INTO " + tableName ;
                    sql += string.Format("( {0} ", fields[0]);
                    for (int i = 1; i < fields.Length;i++ ) 
                    {
                        sql += string.Format(" ,{0} ", fields[i]);
                    }
                    sql += " ) VALUES ( ";
                    sql += string.Format("( {0} ", dr[0]);
                    for (int i = 1; i < fields.Length; i++)
                    {
                        sql += string.Format(" ,{0} ", dr[i]);
                    }
                    sql += ")";


                    OleDbCommand cmd = _odcConnection.CreateCommand(); 
                    cmd.CommandText = sql;
                    //OracleCommand cmd1 = new OracleCommand(sqlStr1, conn);
                    int result = cmd.ExecuteNonQuery();
                    if (result > 0)
                    {
                        MessageBox.Show("");
                        //输出                    
                    }
                }
            }
                 

          
        }
        catch (Exception)
        {
            throw new Exception("插入数据执行失败！");
        }
              
      
    
    
    }

        /// <summary>  
        /// 返回某一表的所有字段名  
        /// </summary>  
        public static string[] GetTableColumn(string database_path, string varTableName)
        {
            DataTable dt = new DataTable();

            OleDbConnection conn;
            string _connectionString = "Provider = Microsoft.Jet.OleDb.4.0;Data Source=" + database_path;
           

            try
            {
                // 建立连接  
                conn = new OleDbConnection(_connectionString);

                // 打开连接  
                conn.Open();

            }
            catch (Exception)
            {
                throw new Exception("尝试打开 " + database_path + " 失败，确认文件是否存在！");
            }


            try
            {              
               
               
                dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, new object[] { null, null, varTableName, null });
                int n = dt.Rows.Count;
                string[] strTable = new string[n];
                int m = dt.Columns.IndexOf("COLUMN_NAME");
                for (int i = 0; i < n; i++)
                {
                    DataRow m_DataRow = dt.Rows[i];
                    strTable[i] = m_DataRow.ItemArray.GetValue(m).ToString();
                }
                return strTable;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                conn.Close();
            }
        }  
    
    }
}
