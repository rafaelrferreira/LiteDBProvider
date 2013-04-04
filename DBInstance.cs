using System;
using System.Collections.Generic;
using System.Text;
using System.Data.Common;

using LiteDBProvider.DataAccessLayer;

namespace LiteDBProvider.DataAccessLayer
{
    class DBInstance
    {
        //Variáveis Globais da Classe
        #region {...}

        private DBConnectionProvider dbConnectionProvider;
        private DbConnection dbConnection;

        public DBConnectionProvider DbConnectionProvider { get; set; }
        public DbConnection DbConnection { get; set; }
        public DbProviderFactory Factory { get; set; }
        public DbCommand Command { get; set; }

        #endregion

        /// <summary>
        /// 
        /// </summary>
        public DBInstance()
        {
            //Cria uma conexao na base padrão do AppSettings (Web.config)
            this.dbConnectionProvider = new DBConnectionProvider();
        }
    }
}
