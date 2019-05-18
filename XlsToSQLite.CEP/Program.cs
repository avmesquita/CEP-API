using System;
using System.Linq;
using System.Data;
using System.Data.SQLite;

namespace ImportacaoCEP
{
	class Program
	{
		static void Main(string[] args)
		{
			#region [ LOAD EXCEL FILE ]

			var arquivo = ".\\Data\\CEPs.xlsx";
			var tabela = new ImportarCEP().Importar(arquivo);
			var linhas = tabela.AsEnumerable().ToArray();

			Console.WriteLine("XLS LINES FOUND = " + linhas.Count().ToString());

			#endregion

			#region [ CONFIGURE SQLITE ]
			string sqlitedb = @".\cep.db3";

			SQLiteConnectionStringBuilder connSB = new SQLiteConnectionStringBuilder();
			connSB.DataSource = sqlitedb;
			connSB.FailIfMissing = false;
			connSB.Version = 3;
			connSB.LegacyFormat = true;
			connSB.Pooling = true;
			connSB.JournalMode = SQLiteJournalModeEnum.Persist;

			SQLiteConnection sqlite = new SQLiteConnection(connSB.ConnectionString);			
			#endregion

			#region [ SQLite INSERT Routine ]

			string sql = string.Empty;

			long execucaoLite = 0;
			sqlite.Open();
			foreach (var row in linhas)
			{
				sql = string.Format("INSERT INTO cep (TXT_CEP, TXT_CIDADE_UF, TXT_BAIRRO, TXT_LOCALIDADE) VALUES ('{0}','{1}','{2}','{3}') ", clean(row[0].ToString()), clean(row[1].ToString()), clean(row[2].ToString()), clean(row[3].ToString()));

				var cmdLite = new SQLiteCommand(sql, sqlite);

				execucaoLite += cmdLite.ExecuteNonQuery();
			}			

			sqlite.Close();

			Console.WriteLine("SQLite Imported = " + execucaoLite.ToString());

			#endregion
		}

		private static string clean(string texto)
		{
			texto = texto.Trim().Replace("'", "´");

			return texto;
		}

	}
}
