using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Linq;
using System.Threading.Tasks;

namespace api.CEP.Infrastructure
{
	/// <summary>
	/// SQLite Connection class
	/// </summary>
	public class SQLiteConnect : IDisposable
	{
		#region [ PROPERTIES ]

		public string DataFile { get; set; }

		protected SQLiteConnection _connection;
		protected string _connectionString = string.Empty;

		private SQLiteConnection Connection { get { return GetConnection(); } }
		private string ConnectionString { get { return GetConnectionString(); } }

		#endregion

		#region [ CONSTRUCTORS / DESTRUCTORS ]

		public SQLiteConnect()
		{
			this.DataFile = "./Database/cep.db3";
		}

		public SQLiteConnect(string DataFile)
		{
			this.DataFile = DataFile;
		}

		~SQLiteConnect()
		{
			_connection.Close();
		}

		public void Dispose()
		{
			_connection.Dispose();
		}

		#endregion

		#region [ PUBLIC METHODS ]

		/// <summary>
		/// EXECUTE NON QUERY COMMANDS (INSERT,UPDATE,DELETE)
		/// </summary>
		/// <param name="sql"></param>
		/// <returns></returns>
		public int ExecuteNonQuery(string sql)
		{
			var cmdLite = new SQLiteCommand(sql, Connection);

			return cmdLite.ExecuteNonQuery();
		}

		/// <summary>
		/// EXECUTE QUERY COMMAND (SELECT)
		/// </summary>
		/// <param name="sql"></param>
		/// <returns></returns>
		public DataTable ExecuteQuery(string sql)
		{
			DataTable result = new DataTable();

			SQLiteCommand cmd = new SQLiteCommand(sql, Connection);

			SQLiteDataAdapter da = new SQLiteDataAdapter(cmd);
			da.Fill(result);

			return result;
		}
		
		#endregion

		#region [ PRIVATE METHODS ]

		private SQLiteConnection GetConnection()
		{
			_connection = new SQLiteConnection(ConnectionString);
			_connection.Open();
			return _connection;
		}

		private string GetConnectionString()
		{
			// VALIDATION
			if (string.IsNullOrEmpty(this.DataFile))
			{
				throw new Exception("DataFile was not configured. Set DataFile first!");
			}

			// CREATE A CONNECTION STRING
			var connSB = new SQLiteConnectionStringBuilder();
			connSB.DataSource = this.DataFile;
			_connectionString = connSB.ConnectionString;
			return _connectionString;
		}

		#endregion
	}
}
