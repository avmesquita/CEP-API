using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace api.CEP.Entity
{
	public class Cep
	{
		public long Id { get; set; }
		public string Codigo { get; set; }
		public string CidadeUf { get; set; }
		public string Bairro { get; set; }
		public string Localidade { get; set; }
	}
}
