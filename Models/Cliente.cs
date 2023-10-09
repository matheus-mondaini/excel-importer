using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DesafioImportaExcel.Models
{
    public class Cliente
    {
        public int ID { get; set; }
        public string Nome { get; set; }
        public string Cidade { get; set; }
        public string UF { get; set; }
        public string CEP { get; set; }
        public string CPF { get; set; }
    }
}
