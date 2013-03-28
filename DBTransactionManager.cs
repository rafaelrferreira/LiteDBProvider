using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.Common;
using System.Data;

using LiteDBProvider.DataAccessLayer;

namespace LiteDBProvider.DataAccessLayer
{
    /// <summary>
    /// 
    /// </summary>
    public sealed class DBTransactionManager
    {
        private DBConnectionProvider dbConnectionProvider;
        private DbConnection dbConnection;
        private DbProviderFactory factory;

        private DbCommand command;
        private DbTransaction transaction = null;
        private DbParameter parameter = null;
        private bool transHold = false;

        /// <summary>
        /// 
        /// </summary>
        public DbConnection DbConnection
        {
            get { return dbConnection; }
            set { dbConnection = value; }
        }
        /// <summary>
        /// 
        /// </summary>
        public DbCommand Command
        {
            get { return command; }
            set { command = value; }
        }
        /// <summary>
        /// 
        /// </summary>
        public DbTransaction Transaction
        {
            get { return transaction; }
            set { transaction = value; }
        }
        /// <summary>
        /// 
        /// </summary>
        public DbProviderFactory Factory
        {
            get { return factory; }
            set { factory = value; }
        }
        /// <summary>
        /// 
        /// </summary>
        public bool TransHold
        {
            get { return transHold; }
            set { transHold = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public enum TransactionState
        {
            Begin,
            End,
            ContinueExecution,
            StandAlone
        };

        /// <summary>
        /// 
        /// </summary>
        public DBTransactionManager()
        {
            //Cria uma conexao na base padrão do AppSettings (Web.config)
            dbConnectionProvider = new DBConnectionProvider();
            dbConnection = dbConnectionProvider.GetConnection();

            //Cria o objeto command a partir do tipo do Provider do AppSettings (Web.config)
            factory = DBConnectionProvider.GetDbFactory();
            command = factory.CreateCommand();
        }
       
        /// <summary>
        /// 
        /// </summary>
        public DBTransactionManager(string _connectionKey)
        {
            //Cria uma conexao na base padrão do AppSettings (Web.config)
            dbConnectionProvider = new DBConnectionProvider();
            dbConnection = dbConnectionProvider.GetConnection(_connectionKey);

            //Cria o objeto command a partir do tipo do Provider do AppSettings (Web.config)
            factory = DBConnectionProvider.GetDbFactory();
            command = factory.CreateCommand();
        }

        /// <summary>
        /// 
        /// </summary>
        public void EndTransaction()
        {
            if (command != null)
            {
                command.Dispose();
            }
            if (dbConnection.State == ConnectionState.Open)
            {
                dbConnection.Close();
            }
            if (dbConnection != null)
            {
                dbConnection.Dispose();
            }
        }
    }
}
