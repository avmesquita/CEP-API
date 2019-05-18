using api.CEP.Entity;
using api.CEP.Infrastructure;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace api.CEP.Data
{
	public class CepDAO : SQLiteConnect
	{
		private IConfiguration _configuration;

		public CepDAO(IConfiguration configuration)
		{
			_configuration = configuration;
		}

		public dynamic Search(string s)
		{
			string sql = " SELECT * FROM CEP WHERE TXT_CEP like '%{0}%' or TXT_CIDADE_UF like '%{0}%' or TXT_BAIRRO like '%{0}%' or TXT_LOCALIDADE like '%{0}%' ";

			var data = this.ExecuteQuery(string.Format(sql, s));

			List<Cep> list = new List<Cep>();
			
			
			foreach (DataRow row in data.Rows)
			{
				Cep _cep = new Cep
				{
					Id = Convert.ToInt64(row["ID"]),
					Codigo = Convert.ToString(row["TXT_CEP"]),
					CidadeUf = Convert.ToString(row["TXT_CIDADE_UF"]),
					Bairro = Convert.ToString(row["TXT_BAIRRO"]),
					Localidade = Convert.ToString(row["TXT_LOCALIDADE"])
				};
				list.Add(_cep);
			}
			
			return list;
		}

		public dynamic Search(long i)
		{
			// FIND DIRECT FIRST
			string sql = " SELECT * FROM CEP WHERE TXT_CEP = '{0}' ";

			var data = this.ExecuteQuery(string.Format(sql, i.ToString()));

			// IF HAS NO ROWS, LOOK BETWEEN RANGES
			if (data.Rows.Count == 0)
			{
				string cepBefore = i.ToString();
				string cepAfter = i.ToString();

				cepBefore = cepBefore.Substring(0, 4) + "000";
				cepAfter = cepAfter.Substring(0, 4) + "999";

				string sqlInteligence = " SELECT * from CEP WHERE CAST(TXT_CEP AS BIGINT) BETWEEN {0} AND {1}";

				sqlInteligence = string.Format(sqlInteligence, cepBefore, cepAfter);

				data = this.ExecuteQuery(string.Format(sqlInteligence, i.ToString()));
			}

			// CREATE A LIST STRUCTURE
			List<Cep> list = new List<Cep>();
			foreach (DataRow row in data.Rows)
			{
				Cep _cep = new Cep
				{
					Id = Convert.ToInt64(row["ID"]),
					Codigo = Convert.ToString(row["TXT_CEP"]),
					CidadeUf = Convert.ToString(row["TXT_CIDADE_UF"]),
					Bairro = Convert.ToString(row["TXT_BAIRRO"]),
					Localidade = Convert.ToString(row["TXT_LOCALIDADE"])
				};
				list.Add(_cep);
			}

			// RETURN RESULTS
			return list;
		}

	}
}
