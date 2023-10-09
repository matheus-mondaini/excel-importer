using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;

namespace DesafioImportaExcel.Controllers
{
    public class TabelaDebitos
    {
        private readonly string _connectionString;

        public TabelaDebitos(string connectionString)
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
                        IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'Debitos')
                        BEGIN
                            CREATE TABLE Debitos
                            (
                                ID INT PRIMARY KEY IDENTITY(1,1),
                                Fatura NVARCHAR(255),
                                Cliente INT,
                                Emissao DATETIME,
                                Vencimento DATETIME,
                                Valor DECIMAL(18, 2),
                                Juros DECIMAL(18, 2),
                                Descontos DECIMAL(18, 2),
                                Pagamento DATETIME,
                                ValorPago DECIMAL(18, 2)
                            )
                        END";

                    using (SqlCommand command = new SqlCommand(createTableQuery, connection))
                    {
                        var NumeroDeLinhas = command.ExecuteNonQuery();
                        if (NumeroDeLinhas > 0)
                            MessageBox.Show("Tabela de Debitos Criada");
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
