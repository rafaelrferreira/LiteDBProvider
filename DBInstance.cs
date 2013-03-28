using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.Common;

using TechShift.DataAccessLayer;

namespace DataLayer
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
