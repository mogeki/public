using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Configuration;

namespace WindowsFormsApp1
{
    class DbConnect
    {
        private static SqlConnection conn = null;

        public static SqlConnection GetCOnnection()
        {
            if (conn == null)
            {
                //DB接続文字列情報の取得
                var connectionStringSettings = ConfigurationManager.ConnectionStrings["DBCON"];

                ////DbProficerFactoryインスタンスの生成、取得
                //var dbProviderFactory = DbProviderFactories.GetFactory(connectionStringSettings.ProviderName);

                ////DB接続オブジェクトを作成
                //conn = dbProviderFactory.CreateConnection();

                ////DB接続文字列の設定
                //conn.ConnectionString = connectionStringSettings.ConnectionString;

                conn = new SqlConnection(connectionStringSettings.ConnectionString);

                //DB接続を開く
                conn.Open();
            }

            return conn;
        }
    }
}
