using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Configuration;
using System.Drawing.Text;
using System.Data;

namespace DesafioImportaExcel.Controllers
{
    public class GerenciadorConexaoBancoDados
    {

        public static String connectionString = ConfigurationManager.ConnectionStrings["MyDBConnectionString"].ConnectionString.ToString();
        private static SqlConnection sqlConnection = new SqlConnection(connectionString);
        
        public static SqlConnection Conectar()
        {
            try
            {
                sqlConnection.Open();
                return sqlConnection;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao acessar o banco de dados: {ex.Message}");
                throw;
            }
        }
    }
}
