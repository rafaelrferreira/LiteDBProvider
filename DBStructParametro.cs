using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.Common;
using System.Data;
using System.Collections;

namespace LiteDBProvider.DataAccessLayer
{
    /// <summary>
    /// 
    /// </summary>
    public struct DbParametro
    {
        private string nome;
        private object valor;
        private DbType tipo;
        private DbParameter inout;
        private int tam;
        private static DbProviderFactory factory;

        public string Nome
        {
            get { return nome; }
            set { nome = value; }
        }
        public object Valor
        {
            get { return valor; }
            set { valor = value; }
        }
        public DbType Tipo
        {
            get { return tipo; }
            set { tipo = value; }
        }
        public DbParameter Inout
        {
            get { return inout; }
            set { inout = value; }
        }
        public int Tam
        {
            get { return tam; }
            set { tam = value; }
        }


        /// <summary>
        /// Método que cria e retorna um parâmetro DbParameter do tipo da DbFactory da tag AppSettings do web.config.
        /// </summary>
        /// <param name="_nome">Nome do parâmetro.</param>
        /// <param name="_valor">Valor do parâmetro.</param>
        /// <param name="_tipo">Tipo do parâmetro (DbType).</param>
        public static DbParameter Parametro(string _nome, object _valor, DbType _tipo)
        {
            factory = DBConnectionProvider.GetDbFactory();
            DbParameter parameter = factory.CreateParameter();

            parameter.ParameterName = _nome;
            parameter.Value  = _valor;
            parameter.DbType = _tipo;

            return parameter;
        }

        /// <summary>
        /// Método que cria e retorna um parâmetro DbParameter do tipo da DbFactory da tag AppSettings do web.config.
        /// </summary>
        /// <param name="_nome">Nome do parâmetro.</param>
        /// <param name="_valor">Valor do parâmetro.</param>
        /// <param name="_tipo">Tipo do parâmetro (DbType).</param>
        /// <param name="_direction">Direção do parâmetro (ParameterDirection).</param>
        public static DbParameter Parametro(string _nome, object _valor, DbType _tipo, ParameterDirection _direction)
        {
            factory = DBConnectionProvider.GetDbFactory();
            DbParameter parameter = factory.CreateParameter();

            parameter.ParameterName = _nome;
            parameter.Value  = _valor;
            parameter.DbType = _tipo;
            parameter.Direction = _direction;

            return parameter;
        }
    }
}
