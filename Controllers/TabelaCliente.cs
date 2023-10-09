using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DesafioImportaExcel.Controllers
{
    public class TabelaCliente
    {
        private readonly string _connectionString;

        public TabelaCliente(string connectionString)
        {
            _connectionString = connectionString;
        }

        public void CriarTabelaSeNaoExistir()
        {
            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                try
                {
                    connection.Open();

                    string createTableQuery = @"
                        IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'Cliente')
                        BEGIN
                            CREATE TABLE Cliente
                            (
                                ID INT IDENTITY(1,1) NOT NULL,
                                Nome NVARCHAR(255),
                                Cidade NVARCHAR(255),
                                UF NVARCHAR(2),
                                CEP NVARCHAR(9),
                                CPF NVARCHAR(14)
                            )
                        END";

                    using (SqlCommand command = new SqlCommand(createTableQuery, connection))
                    {
                        var NumberOfRows = command.ExecuteNonQuery();
                        if (NumberOfRows > 0)
                            MessageBox.Show("Tabela de Cliente Criada");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ocorreu um erro: " + ex.Message);
                }
                finally
                {
                    connection.Close();
                }
            }
        }
    }
}
