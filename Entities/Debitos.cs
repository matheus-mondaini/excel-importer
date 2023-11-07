using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DesafioImportaExcel.Models
{
    public class Debitos
    {
        public string Fatura { get; set; } = string.Empty;
        public int Cliente { get; set; }
        public DateTime Emissao { get; set; }
        public DateTime Vencimento { get; set; }
        public decimal Valor { get; set; }
        public decimal Juros { get; set; }
        public decimal Descontos { get; set; }
        public DateTime? Pagamento { get; set; }
        public decimal ValorPago { get; set; }
    }

}
