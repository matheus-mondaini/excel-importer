using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DesafioImportaExcel
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
                                ID INT PRIMARY KEY IDENTITY(1,1),
                                Nome NVARCHAR(255),
                                Cidade NVARCHAR(255),
                                UF NVARCHAR(2),
                                CEP NVARCHAR(8),
                                CPF NVARCHAR(11)
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
