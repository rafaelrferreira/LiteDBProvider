using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.Common;

namespace TechShift.DataAccessLayer
{
    /// <summary>
    /// Struct de encapsulamento de itens de um bloco de execução (query/procedure + msgErro + lista de parâmetros)
    /// </summary>
    public struct DBTransactionBlock
    {
        private string inQuery;
        private string _errorMessage;
        private DbParameter[] parametro;

        /// <summary>
        /// Query a ser executada.
        /// </summary>
        public string InQuery
        {
            get { return inQuery; }
            set { inQuery = value; }
        }

        /// <summary>
        /// Mensagem de erro no retorno.
        /// </summary>
        public string ErrorMessage
        {
            get { return _errorMessage; }
            set { _errorMessage = value; }
        }

        /// <summary>
        /// Lista de Parâmetros do tipo da Struct DBStructParamentro.
        /// </summary>
        public DbParameter[] Parametro
        {
            get { return parametro; }
            set { parametro = value; }
        }
    }
}
