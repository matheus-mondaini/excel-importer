using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace DesafioImportaExcel
{ 
    public class GerenciadorConexaoBancoDados
    {
        private readonly string _server;
        private readonly string _databaseName;
        private readonly string _id;
        private readonly string _password;

        public GerenciadorConexaoBancoDados(string server, string databaseName, string id, string password)
        {
            _server = server;
            _databaseName = databaseName;
            _id = id;
            _password = password;

            GetConnection();
    }

        public SqlConnection GetConnection()
        {
            string connectionString = $"Server={_server};Database={_databaseName};User Id = {_id}; Password = {_password};;";
            return new SqlConnection(connectionString);
        }

    }

}
