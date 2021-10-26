using Npgsql;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Json2ExcelAndXml.SQL
{
    class AddComment
    {
        NpgsqlConnection conn;
        public AddComment()
        {
            //var conString = "Host=localhost;Port=5432;Username=postgres;Password=PgSQL2021; Database=StandardShps;Pooling=true;Minimum Pool Size=1";

            //using (conn = new NpgsqlConnection(conString))
            //{
            //    try
            //    {
            //        conn.Open();
            //        conn.Close();
            //    }
            //    catch (Exception e)
            //    {

            //        throw e;
            //    }

            //}
        }


        public int ExecuteNonQuery(string sqrstr)
        {
            var conString = "Host=localhost;Port=5432;Username=postgres;Password=PgSQL2021; Database=StandardShps;Pooling=true;Minimum Pool Size=1";

            using (conn = new NpgsqlConnection(conString))
            {
                try
                {
                    conn.Open();
                    using (NpgsqlCommand SqlCommand = new NpgsqlCommand(sqrstr, conn))
                    {
                        int r = SqlCommand.ExecuteNonQuery();  //执行查询并返回受影响的行数
                        conn.Close();
                        return r; //r如果是>0操作成功！ 
                    }
                }
                catch (System.Exception ex)
                {
                    CloseConnection();
                    throw ex;
                }

            }

        }

        private void CloseConnection()
        {
            conn.Close();
        }
    }
}
