using System;
using System.Collections.Generic;
using System.Web;
using System.Data.Common;
using System.Data;
using System.Collections;
using System.Data.SqlClient;

using LiteDBProvider.DataAccessLayer;

namespace LiteDBProvider.DataAccessLayer
{
    /// <summary>
    /// Classe base com os métodos de execução dos comandos SQL
    /// </summary>
    public sealed class DBProvider
    {

        /* 
        * 
        * SQLClient, OleDb, Odbc & Oracle Database Provider
        * 
        */

        #region { Instance }

        //Query Instance

        /// <summary>
        /// Executa uma query sem parâmetros, retornando um Int32, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <returns>Int32 indicando a quantidade de linhas afetadas (INSERT, UPDATE ou DELETE).</returns>
        public static int Instance(string _inQuery, out string _errorMessage)
        {
            DBInstance dbInstance = new DBInstance();

            _errorMessage = String.Empty;
            int recordsAffected = 0;

            if (_inQuery.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return recordsAffected = -1;
            }

            try
            {
                using (dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection())
                using (dbInstance.Command = dbInstance.Factory.CreateCommand())
                {
                    dbInstance.Factory = DBConnectionProvider.GetDbFactory();

                    dbInstance.Command.CommandText = _inQuery;
                    dbInstance.Command.CommandType = CommandType.Text;
                    dbInstance.Command.Connection = dbInstance.DbConnection;

                    dbInstance.DbConnection.Open();

                    recordsAffected = dbInstance.Command.ExecuteNonQuery();
                }
            }
            catch(DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }

            return recordsAffected;
        }

        /// <summary>
        /// Executa uma query sem parâmetros, retornando um Int32, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <returns>Int32 indicando a quantidade de linhas afetadas (INSERT, UPDATE ou DELETE).</returns>
        public static int Instance(string _connectionKey, string _inQuery, out string _errorMessage)
        {
            DBInstance dbInstance = new DBInstance();

            _errorMessage = String.Empty;
            int recordsAffected = 0;

            if (_inQuery.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return recordsAffected = -1;
            }

            try
            {
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inQuery;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                dbInstance.DbConnection.Open();

                recordsAffected = dbInstance.Command.ExecuteNonQuery();

                dbInstance.DbConnection.Close();
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return recordsAffected;
        }

        /// <summary>
        /// Executa uma query sem parâmetros, retornando um Int32, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_providerKey">Chave do Provider do AppSettings do web.config.</param>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <returns>Int32 indicando a quantidade de linhas afetadas (INSERT, UPDATE ou DELETE).</returns>
        public static int Instance(string _providerKey, string _connectionKey, string _inQuery, out string _errorMessage)
        {
            DBInstance dbInstance = new DBInstance();

            _errorMessage = String.Empty;
            int recordsAffected = 0;

            if (_inQuery.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return recordsAffected = -1;
            }

            try
            {
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_providerKey, _connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory(_providerKey);
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inQuery;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                dbInstance.DbConnection.Open();

                recordsAffected = dbInstance.Command.ExecuteNonQuery();

                dbInstance.DbConnection.Close();
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return recordsAffected;
        }

        /// <summary>
        /// Executa uma query com parâmetros, retornando um Int32, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>Int32 indicando a quantidade de linhas afetadas (INSERT, UPDATE ou DELETE).</returns>
        public static int Instance(string _inQuery, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();

            DbParameter parameter = null;
            int recordsAffected = 0;

            _errorMessage = String.Empty;

            if (_inQuery.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return recordsAffected = -1;
            }

            try
            {
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection();

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inQuery;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                parameter = dbInstance.Factory.CreateParameter();

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }      

                dbInstance.DbConnection.Open();

                recordsAffected = dbInstance.Command.ExecuteNonQuery();

                dbInstance.DbConnection.Close();
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return recordsAffected;

        }

        /// <summary>
        /// Executa uma query com parâmetros, retornando um Int32, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>Int32 indicando a quantidade de linhas afetadas (INSERT, UPDATE ou DELETE).</returns>
        public static int Instance(string _connectionKey, string _inQuery, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();

            DbParameter parameter = null;
            int recordsAffected = 0;

            _errorMessage = String.Empty;

            if (_inQuery.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return recordsAffected = -1;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inQuery;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                parameter = dbInstance.Factory.CreateParameter();

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dbInstance.DbConnection.Open();

                recordsAffected = dbInstance.Command.ExecuteNonQuery();

                dbInstance.DbConnection.Close();
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return recordsAffected;
        }

        /// <summary>
        /// Executa uma query com parâmetros, retornando um Int32, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_providerKey">Chave do Provider do AppSettings do web.config.</param>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>Int32 indicando a quantidade de linhas afetadas (INSERT, UPDATE ou DELETE).</returns>
        public static int Instance(string _providerKey, string _connectionKey, string _inQuery, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();

            DbParameter parameter = null;
            int recordsAffected = 0;

            _errorMessage = String.Empty;

            if (_inQuery.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return recordsAffected = -1;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_providerKey, _connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory(_providerKey);
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inQuery;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                parameter = dbInstance.Factory.CreateParameter();

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dbInstance.DbConnection.Open();

                recordsAffected = dbInstance.Command.ExecuteNonQuery();

                dbInstance.DbConnection.Close();
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return recordsAffected;

        }

        //Query Instance Transaction - Execução Independente de Comandos de uma Transação (blocos isolados)

        /// <summary>
        ///  Executa uma query sem parâmetros, retornando um Int32, utilizando a conexão do banco e transações.
        /// </summary>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_transManager">Objeto DBTransactionManager da transação.</param>
        /// <param name="_transState">Estado de execução do bloco da transação. (Begin, End, ContinueExecution, StandAlone)</param>
        /// <returns>Int32 indicando a quantidade de linhas afetadas (INSERT, UPDATE ou DELETE).</returns>
        public static int Instance(string _inQuery, out string _errorMessage, ref DBTransactionManager _transManager, DBTransactionManager.TransactionState _transState)
        {
            _errorMessage = String.Empty;
            int recordsAffected = 0;

            if (_inQuery.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return recordsAffected = -1;
            }

            try
            {
                //Inicia uma transação isolada
                if (_transState == DBTransactionManager.TransactionState.StandAlone)
                {
                    //Abre a conexão e inicia a transação
                    if (_transManager.DbConnection.State == ConnectionState.Closed)
                        _transManager.DbConnection.Open();

                    _transManager.Transaction = _transManager.DbConnection.BeginTransaction();

                    QueryTransactionExecute(_inQuery, out recordsAffected, _transManager);

                    //Comita a transação
                    _transManager.Command.Transaction.Commit();

                    //Fecha a conexão da transação
                    _transManager.EndTransaction();
                }
                else
                {
                    //Bloqueio de execução depois da detecção de um erro
                    if (!_transManager.TransHold)
                    {
                        //Inicia a transação.DBTransactionManager(true)
                        if (_transState == DBTransactionManager.TransactionState.Begin)
                        {
                            //Abre a conexão e inicia a transação
                            if (_transManager.DbConnection.State == ConnectionState.Closed)
                                _transManager.DbConnection.Open();

                            _transManager.Transaction = _transManager.DbConnection.BeginTransaction();

                            QueryTransactionExecute(_inQuery, out recordsAffected, _transManager);
                        }
                        //Permite a execução de + de três comandos em uma transação
                        else if (_transState == DBTransactionManager.TransactionState.ContinueExecution)
                        {
                            QueryTransactionExecute(_inQuery, out recordsAffected, _transManager);
                        }
                        //Finaliza a transação.DBTransactionManager(false)
                        else
                        {
                            QueryTransactionExecute(_inQuery, out recordsAffected, _transManager);

                            //Comita a transação
                            _transManager.Command.Transaction.Commit();

                            //Fecha a conexão da transação
                            _transManager.EndTransaction();
                        }
                    }
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção de banco ocorreu durante a execução dos comandos. Nenhuma das operações foi executada. <BR>";
                _errorMessage = _errorMessage + exp.Message;

                if (_transManager.Command.Transaction != null)
                    _transManager.Command.Transaction.Rollback();

                //Bloqueia a execução de qualquer chamada posterior dentro da transação em execução
                _transManager.TransHold = true;

                //Fecha a conexão da transação
                _transManager.EndTransaction();
            }
            catch (Exception exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução dos comandos. Nenhuma das operações foi executada. <BR>";
                _errorMessage = _errorMessage + exp.Message;

                if (_transManager.Command.Transaction != null)
                    _transManager.Command.Transaction.Rollback();

                //Bloqueia a execução de qualquer chamada posterior dentro da transação em execução
                _transManager.TransHold = true;

                //Fecha a conexão da transação
                _transManager.EndTransaction();
            }

            return recordsAffected;
        }

        /// <summary>
        ///  Executa uma query sem parâmetros, retornando um Int32, utilizando a conexão do banco e transações.
        /// </summary>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_transManager">Objeto DBTransactionManager da transação.</param>
        /// <param name="_transState">Estado de execução do bloco da transação. (Begin, End, ContinueExecution, StandAlone)</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>Int32 indicando a quantidade de linhas afetadas (INSERT, UPDATE ou DELETE).</returns>
        public static int Instance(string _inQuery, out string _errorMessage, ref DBTransactionManager _transManager, DBTransactionManager.TransactionState _transState, params DbParameter[] _params)
        {
            _errorMessage = String.Empty;
            int recordsAffected = 0;

            if (_inQuery.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return recordsAffected = -1;
            }

            try
            {
                //Inicia uma transação isolada
                if (_transState == DBTransactionManager.TransactionState.StandAlone)
                {
                    //Abre a conexão e inicia a transação
                    if (_transManager.DbConnection.State == ConnectionState.Closed)
                        _transManager.DbConnection.Open();

                    _transManager.Transaction = _transManager.DbConnection.BeginTransaction();

                    QueryTransactionExecute(_inQuery, out recordsAffected, _transManager, _params);

                    //Comita a transação
                    _transManager.Command.Transaction.Commit();

                    //Fecha a conexão da transação
                    _transManager.EndTransaction();
                }
                else
                {
                    //Bloqueio de execução depois da detecção de um erro
                    if (!_transManager.TransHold)
                    {
                        //Inicia a transação.DBTransactionManager(true)
                        if (_transState == DBTransactionManager.TransactionState.Begin)
                        {
                            //Abre a conexão e inicia a transação
                            if (_transManager.DbConnection.State == ConnectionState.Closed)
                                _transManager.DbConnection.Open();

                            _transManager.Transaction = _transManager.DbConnection.BeginTransaction();

                            QueryTransactionExecute(_inQuery, out recordsAffected, _transManager, _params);
                        }
                        //Permite a execução de + de três comandos em uma transação
                        else if (_transState == DBTransactionManager.TransactionState.ContinueExecution)
                        {
                            QueryTransactionExecute(_inQuery, out recordsAffected, _transManager, _params);
                        }
                        //Finaliza a transação.DBTransactionManager(false)
                        else
                        {
                            QueryTransactionExecute(_inQuery, out recordsAffected, _transManager, _params);

                            //Comita a transação
                            _transManager.Command.Transaction.Commit();

                            //Fecha a conexão da transação
                            _transManager.EndTransaction();
                        }
                    }
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção de banco ocorreu durante a execução dos comandos. Nenhuma das operações foi executada. <BR>";
                _errorMessage = _errorMessage + exp.Message;

                if (_transManager.Command.Transaction != null)
                    _transManager.Command.Transaction.Rollback();

                //Bloqueia a execução de qualquer chamada posterior dentro da transação em execução
                _transManager.TransHold = true;

                //Fecha a conexão da transação
                _transManager.EndTransaction();
            }
            catch (Exception exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução dos comandos. Nenhuma das operações foi executada. <BR>";
                _errorMessage = _errorMessage + exp.Message;

                if (_transManager.Command.Transaction != null)
                    _transManager.Command.Transaction.Rollback();

                //Bloqueia a execução de qualquer chamada posterior dentro da transação em execução
                _transManager.TransHold = true;

                //Fecha a conexão da transação
                _transManager.EndTransaction();
            }

            return recordsAffected;
        }

        /// <summary>
        /// Executa uma ou várias querys sem parâmetros, retornando um DataSet com o resultado de cada query, utilizando a conexão do banco e transações.
        /// </summary>
        /// <param name="_inQuerySet">Querys de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <returns>DataSet com vários DataTables, cada DataTable com o resultado de uma das querys.</returns>
        public static DataSet InstanceResultSet(string _inQuerySet, out string _errorMessage)
        {
            DBInstance dbInstance = new DBInstance();

            DataSet dataSet = new DataSet();
            DbDataReader dataReader = null;

            _errorMessage = String.Empty;

            if (_inQuerySet.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataSet;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection();

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();

                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inQuerySet;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                dbInstance.Command.Connection.Open();

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);

                if (!dataReader.HasRows)
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    dbInstance.Command.Connection.Close();

                    return null;
                }
                else
                {
                    int count = 0;
                    while (!dataReader.IsClosed)
                    {
                        dataSet.Tables.Add(new DataTable());
                        dataSet.Tables[count].Load(dataReader, LoadOption.OverwriteChanges);

                        count++;
                    }
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataSet);
        }

        /// <summary>
        /// Executa uma ou várias querys sem parâmetros, retornando um DataSet com o resultado de cada query, utilizando a conexão do banco e transações.
        /// </summary>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inQuerySet">Querys de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <returns>DataSet com vários DataTables, cada DataTable com o resultado de uma das querys.</returns>
        public static DataSet InstanceResultSet(string _connectionKey, string _inQuerySet, out string _errorMessage)
        {
            DBInstance dbInstance = new DBInstance();
            DataSet dataSet = new DataSet();
            DbDataReader dataReader = null;

            _errorMessage = String.Empty;

            if (_inQuerySet.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataSet;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();

                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inQuerySet;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                dbInstance.Command.Connection.Open();

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);

                if (!dataReader.HasRows)
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    dbInstance.Command.Connection.Close();

                    return null;
                }
                else
                {
                    int count = 0;
                    while (!dataReader.IsClosed)
                    {
                        dataSet.Tables.Add(new DataTable());
                        dataSet.Tables[count].Load(dataReader, LoadOption.OverwriteChanges);

                        count++;
                    }
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataSet);
        }

        /// <summary>
        /// Executa uma ou várias querys sem parâmetros, retornando um DataSet com o resultado de cada query, utilizando a conexão do banco e transações.
        /// </summary>
        /// <param name="_providerKey">Chave do Provider do AppSettings do web.config.</param>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inQuerySet">Querys de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <returns>DataSet com vários DataTables, cada DataTable com o resultado de uma das querys.</returns>
        public static DataSet InstanceResultSet(string _providerKey, string _connectionKey, string _inQuerySet, out string _errorMessage)
        {
            DBInstance dbInstance = new DBInstance();
            DataSet dataSet = new DataSet();
            DbDataReader dataReader = null;

            _errorMessage = String.Empty;

            if (_inQuerySet.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataSet;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_providerKey, _connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory(_providerKey);

                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inQuerySet;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                dbInstance.Command.Connection.Open();

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);

                if (!dataReader.HasRows)
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    dbInstance.Command.Connection.Close();

                    return null;
                }
                else
                {
                    int count = 0;
                    while (!dataReader.IsClosed)
                    {
                        dataSet.Tables.Add(new DataTable());
                        dataSet.Tables[count].Load(dataReader, LoadOption.OverwriteChanges);

                        count++;
                    }
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataSet);
        }

        /// <summary>
        /// Executa uma ou várias querys com parâmetros, retornando um DataSet com o resultado de cada query, utilizando a conexão do banco e transações.
        /// </summary>
        /// <param name="_inQuerySet">Querys de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>DataSet com vários DataTables, cada DataTable com o resultado de uma das querys.</returns>
        public static DataSet InstanceResultSet(string _inQuerySet, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();
            DataSet dataSet = new DataSet();
            DbParameter parameter = null;
            DbDataReader dataReader = null;

            _errorMessage = String.Empty;

            if (_inQuerySet.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataSet;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection();

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();

                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inQuerySet;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                parameter = dbInstance.Factory.CreateParameter();

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dbInstance.Command.Connection.Open();

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);

                if (!dataReader.HasRows)
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    dbInstance.Command.Connection.Close();

                    return null;
                }
                else
                {
                    int count = 0;
                    while (!dataReader.IsClosed)
                    {
                        dataSet.Tables.Add(new DataTable());
                        dataSet.Tables[count].Load(dataReader, LoadOption.OverwriteChanges);

                        count++;
                    }
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataSet);
        }

        /// <summary>
        /// Executa uma ou várias querys com parâmetros, retornando um DataSet com o resultado de cada query, utilizando a conexão do banco e transações.
        /// </summary>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inQuerySet">Querys de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>DataSet com vários DataTables, cada DataTable com o resultado de uma das querys.</returns>
        public static DataSet InstanceResultSet(string _connectionKey, string _inQuerySet, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();
            DataSet dataSet = new DataSet();
            DbParameter parameter = null;
            DbDataReader dataReader = null;

            _errorMessage = String.Empty;

            if (_inQuerySet.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataSet;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();

                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inQuerySet;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                parameter = dbInstance.Factory.CreateParameter();

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dbInstance.Command.Connection.Open();

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);

                if (!dataReader.HasRows)
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    dbInstance.Command.Connection.Close();

                    return null;
                }
                else
                {
                    int count = 0;
                    while (!dataReader.IsClosed)
                    {
                        dataSet.Tables.Add(new DataTable());
                        dataSet.Tables[count].Load(dataReader, LoadOption.OverwriteChanges);

                        count++;
                    }
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataSet);
        }

        /// <summary>
        /// Executa uma ou várias querys com parâmetros, retornando um DataSet com o resultado de cada query, utilizando a conexão do banco e transações.
        /// </summary>
        /// <param name="_providerKey">Chave do Provider do AppSettings do web.config.</param>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inQuerySet">Querys de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>DataSet com vários DataTables, cada DataTable com o resultado de uma das querys.</returns>
        public static DataSet InstanceResultSet(string _providerKey, string _connectionKey, string _inQuerySet, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();
            DataSet dataSet = new DataSet();
            DbParameter parameter = null;
            DbDataReader dataReader = null;

            _errorMessage = String.Empty;

            if (_inQuerySet.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataSet;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_providerKey, _connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory(_providerKey);

                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inQuerySet;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                parameter = dbInstance.Factory.CreateParameter();

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dbInstance.Command.Connection.Open();

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);

                if (!dataReader.HasRows)
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    dbInstance.Command.Connection.Close();

                    return null;
                }
                else
                {
                    int count = 0;
                    while (!dataReader.IsClosed)
                    {
                        dataSet.Tables.Add(new DataTable());
                        dataSet.Tables[count].Load(dataReader, LoadOption.OverwriteChanges);

                        count++;
                    }
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataSet);
        }

        //Métodos Privados para DBProvider.Instance

        /// <summary>
        /// Método privado para execução de uma query usando uma transação.
        /// </summary>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_recordsAffected">Int32 indicando a quantidade de linhas afetadas (INSERT, UPDATE ou DELETE).</param>
        /// <param name="_transManager">Objeto DBTransactionManager da transação.</param>
        private static void QueryTransactionExecute(string _inQuery, out int _recordsAffected, DBTransactionManager _transManager)
        {
            //Associa a transação ao comando
            _transManager.Command.Transaction = _transManager.Transaction;

            _transManager.Command.CommandText = _inQuery;
            _transManager.Command.CommandType = CommandType.Text;
            _transManager.Command.Connection = _transManager.DbConnection;

            //Executa o comando
            _recordsAffected = _transManager.Command.ExecuteNonQuery();
        }

        /// <summary>
        /// Método privado para execução de uma query usando uma transação.
        /// </summary>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_recordsAffected">Int32 indicando a quantidade de linhas afetadas (INSERT, UPDATE ou DELETE).</param>
        /// <param name="_transManager">Objeto DBTransactionManager da transação.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        private static void QueryTransactionExecute(string _inQuery, out int _recordsAffected, DBTransactionManager _transManager, params DbParameter[] _params)
        {
            DbParameter parameter = null;

            //Associa a transação ao comando
            _transManager.Command.Transaction = _transManager.Transaction;

            _transManager.Command.CommandText = _inQuery;
            _transManager.Command.CommandType = CommandType.Text;
            _transManager.Command.Connection = _transManager.DbConnection;

            _transManager.Command.Parameters.Clear();

            //Adiciona os parâmetros de execução ao comando da transação
            parameter = _transManager.Factory.CreateParameter();
            for (int i = 0; i < _params.Length; i++)
            {
                parameter = _params[i];

                _transManager.Command.Parameters.Add(parameter);
            }

            //Executa o comando
            _recordsAffected = _transManager.Command.ExecuteNonQuery();
        }

        //Procedure InstanceProcedure

        /// <summary>
        /// Executa uma procedure sem parâmetros, retornando um Int32, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_inProcedure">Procedure de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <returns>Int32 indicando a quantidade de linhas afetadas (INSERT, UPDATE ou DELETE).</returns>
        public static int InstanceProcedure(string _inProcedure, out string _errorMessage)
        {
            DBInstance dbInstance = new DBInstance();

            _errorMessage = String.Empty;
            int recordsAffected = 0;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return recordsAffected = -1;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection();

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inProcedure;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                dbInstance.DbConnection.Open();

                recordsAffected = dbInstance.Command.ExecuteNonQuery();

                dbInstance.DbConnection.Close();
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return recordsAffected;
        }

        /// <summary>
        /// Executa uma procedure sem parâmetros, retornando um Int32, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inProcedure">Procedure de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <returns>Int32 indicando a quantidade de linhas afetadas (INSERT, UPDATE ou DELETE).</returns>
        public static int InstanceProcedure(string _connectionKey, string _inProcedure, out string _errorMessage)
        {
            DBInstance dbInstance = new DBInstance();

            _errorMessage = String.Empty;
            int recordsAffected = 0;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return recordsAffected = -1;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inProcedure;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                dbInstance.DbConnection.Open();

                recordsAffected = dbInstance.Command.ExecuteNonQuery();

                dbInstance.DbConnection.Close();
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return recordsAffected;
        }

        /// <summary>
        /// Executa uma procedure com parâmetros, retornando um Int32, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_inProcedure">Procedure de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>Int32 indicando a quantidade de linhas afetadas (INSERT, UPDATE ou DELETE).</returns>
        public static int InstanceProcedure(string _inProcedure, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();

            DbParameter parameter = null;
            int recordsAffected = 0;

            _errorMessage = String.Empty;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return recordsAffected = -1;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection();

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inProcedure;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                parameter = dbInstance.Factory.CreateParameter();

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dbInstance.DbConnection.Open();

                recordsAffected = dbInstance.Command.ExecuteNonQuery();

                dbInstance.DbConnection.Close();
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return recordsAffected;

        }

        /// <summary>
        /// Executa uma procedure com parâmetros, retornando um Int32, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inProcedure">Procedure de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>Int32 indicando a quantidade de linhas afetadas (INSERT, UPDATE ou DELETE).</returns>
        public static int InstanceProcedure(string _connectionKey, string _inProcedure, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();

            DbParameter parameter = null;
            int recordsAffected = 0;

            _errorMessage = String.Empty;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return recordsAffected = -1;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inProcedure;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                parameter = dbInstance.Factory.CreateParameter();

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dbInstance.DbConnection.Open();

                recordsAffected = dbInstance.Command.ExecuteNonQuery();

                dbInstance.DbConnection.Close();
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return recordsAffected;

        }

        /// <summary>
        /// Executa uma procedure sem parâmetros, retornando um Int32, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_inProcedureSet">Procedure de Entrada com os ResultSets.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <returns>DataSet com vários DataTables, cada DataTable com o resultado de uma das querys da procedure.</returns>
        public static DataSet InstanceProcedureResultSet(string _inProcedureSet, out string _errorMessage)
        {
            DBInstance dbInstance = new DBInstance();

            DataSet dataSet = new DataSet();
            DbDataReader dataReader = null;

            _errorMessage = String.Empty;

            if (_inProcedureSet.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataSet;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection();

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();

                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inProcedureSet;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                dbInstance.Command.Connection.Open();

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);

                if (!dataReader.HasRows)
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    dbInstance.Command.Connection.Close();

                    return null;
                }
                else
                {
                    int count = 0;
                    while (!dataReader.IsClosed)
                    {
                        dataSet.Tables.Add(new DataTable());
                        dataSet.Tables[count].Load(dataReader, LoadOption.OverwriteChanges);

                        count++;
                    }
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataSet);
        }

        /// <summary>
        /// Executa uma procedure sem parâmetros, retornando um Int32, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inProcedureSet">Procedure de Entrada com os ResultSets.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <returns>DataSet com vários DataTables, cada DataTable com o resultado de uma das querys da procedure.</returns>
        public static DataSet InstanceProcedureResultSet(string _connectionKey, string _inProcedureSet, out string _errorMessage)
        {
            DBInstance dbInstance = new DBInstance();

            DataSet dataSet = new DataSet();
            DbDataReader dataReader = null;

            _errorMessage = String.Empty;

            if (_inProcedureSet.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataSet;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();

                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inProcedureSet;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                dbInstance.Command.Connection.Open();

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);

                if (!dataReader.HasRows)
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    dbInstance.Command.Connection.Close();

                    return null;
                }
                else
                {
                    int count = 0;
                    while (!dataReader.IsClosed)
                    {
                        dataSet.Tables.Add(new DataTable());
                        dataSet.Tables[count].Load(dataReader, LoadOption.OverwriteChanges);

                        count++;
                    }
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataSet);
        }

        /// <summary>
        /// Executa uma procedure sem parâmetros, retornando um Int32, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_providerKey">Chave do Provider do AppSettings do web.config.</param>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inProcedureSet">Procedure de Entrada com os ResultSets.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <returns>DataSet com vários DataTables, cada DataTable com o resultado de uma das querys da procedure.</returns>
        public static DataSet InstanceProcedureResultSet(string _providerKey, string _connectionKey, string _inProcedureSet, out string _errorMessage)
        {
            DBInstance dbInstance = new DBInstance();

            DataSet dataSet = new DataSet();
            DbDataReader dataReader = null;

            _errorMessage = String.Empty;

            if (_inProcedureSet.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataSet;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_providerKey, _connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory(_providerKey);

                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inProcedureSet;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                dbInstance.Command.Connection.Open();

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);

                if (!dataReader.HasRows)
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    dbInstance.Command.Connection.Close();

                    return null;
                }
                else
                {
                    int count = 0;
                    while (!dataReader.IsClosed)
                    {
                        dataSet.Tables.Add(new DataTable());
                        dataSet.Tables[count].Load(dataReader, LoadOption.OverwriteChanges);

                        count++;
                    }
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataSet);
        }

        /// <summary>
        /// Executa uma procedure sem parâmetros, retornando um Int32, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_inProcedureSet">Procedure de Entrada com os ResultSets.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>DataSet com vários DataTables, cada DataTable com o resultado de uma das querys da procedure.</returns>
        public static DataSet InstanceProcedureResultSet(string _inProcedureSet, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();

            DataSet dataSet = new DataSet();
            DbParameter parameter = null;
            DbDataReader dataReader = null;

            _errorMessage = String.Empty;

            if (_inProcedureSet.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataSet;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection();

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();

                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inProcedureSet;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                parameter = dbInstance.Factory.CreateParameter();

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dbInstance.Command.Connection.Open();

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);

                if (!dataReader.HasRows)
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    dbInstance.Command.Connection.Close();

                    return null;
                }
                else
                {
                    int count = 0;
                    while (!dataReader.IsClosed)
                    {
                        dataSet.Tables.Add(new DataTable());
                        dataSet.Tables[count].Load(dataReader, LoadOption.OverwriteChanges);

                        count++;
                    }
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataSet);
        }

        /// <summary>
        /// Executa uma procedure sem parâmetros, retornando um Int32, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inProcedureSet">Procedure de Entrada com os ResultSets.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>DataSet com vários DataTables, cada DataTable com o resultado de uma das querys da procedure.</returns>
        public static DataSet InstanceProcedureResultSet(string _connectionKey, string _inProcedureSet, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();

            DataSet dataSet = new DataSet();
            DbParameter parameter = null;
            DbDataReader dataReader = null;

            _errorMessage = String.Empty;

            if (_inProcedureSet.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataSet;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();

                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inProcedureSet;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                parameter = dbInstance.Factory.CreateParameter();

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dbInstance.Command.Connection.Open();

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);

                if (!dataReader.HasRows)
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    dbInstance.Command.Connection.Close();

                    return null;
                }
                else
                {
                    int count = 0;
                    while (!dataReader.IsClosed)
                    {
                        dataSet.Tables.Add(new DataTable());
                        dataSet.Tables[count].Load(dataReader, LoadOption.OverwriteChanges);

                        count++;
                    }
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataSet);
        }

        /// <summary>
        /// Executa uma procedure sem parâmetros, retornando um Int32, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_providerKey">Chave do Provider do AppSettings do web.config.</param>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inProcedureSet">Procedure de Entrada com os ResultSets.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>DataSet com vários DataTables, cada DataTable com o resultado de uma das querys da procedure.</returns>
        public static DataSet InstanceProcedureResultSet(string _providerKey, string _connectionKey, string _inProcedureSet, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();
            DataSet dataSet = new DataSet();
            DbParameter parameter = null;
            DbDataReader dataReader = null;

            _errorMessage = String.Empty;

            if (_inProcedureSet.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataSet;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_providerKey, _connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory(_providerKey);

                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inProcedureSet;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                parameter = dbInstance.Factory.CreateParameter();

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dbInstance.Command.Connection.Open();

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);

                if (!dataReader.HasRows)
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    dbInstance.Command.Connection.Close();

                    return null;
                }
                else
                {
                    int count = 0;
                    while (!dataReader.IsClosed)
                    {
                        dataSet.Tables.Add(new DataTable());
                        dataSet.Tables[count].Load(dataReader, LoadOption.OverwriteChanges);

                        count++;
                    }
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataSet);
        }

        //Procedure Instance Transaction - Execução Idependente de Comandos de uma Transação (blocos isolados)

        /// <summary>
        /// Executa uma procedure sem parâmetros, retornando um Int32, utilizando a conexão do banco e transações.
        /// </summary>
        /// <param name="_inProcedure">Procedure de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_transManager">Objeto DBTransactionManager da transação.</param>
        /// <param name="_transState">Estado de execução do bloco da transação. (Begin, End, ContinueExecution, StandAlone)</param>
        /// <returns>Int32 indicando a quantidade de linhas afetadas (INSERT, UPDATE ou DELETE).</returns>
        public static int InstanceProcedure(string _inProcedure, out string _errorMessage, ref DBTransactionManager _transManager, DBTransactionManager.TransactionState _transState)
        {
            DBInstance dbInstance = new DBInstance();

            _errorMessage = String.Empty;
            int recordsAffected = 0;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return recordsAffected = -1;
            }

            try
            {
                //Inicia uma transação isolada
                if (_transState == DBTransactionManager.TransactionState.StandAlone)
                {
                    //Abre a conexão e inicia a transação
                    if (_transManager.DbConnection.State == ConnectionState.Closed)
                        _transManager.DbConnection.Open();

                    _transManager.Transaction = _transManager.DbConnection.BeginTransaction();

                    ProcedureTransactionExecute(_inProcedure, out recordsAffected, _transManager);

                    //Comita a transação
                    _transManager.Command.Transaction.Commit();

                    //Fecha a conexão da transação
                    _transManager.EndTransaction();
                }

                //Bloqueio de execução depois da detecção de um erro
                if (!_transManager.TransHold)
                {
                    //Inicia a transação.DBTransactionManager(true)
                    if (_transState == DBTransactionManager.TransactionState.Begin)
                    {
                        //Abre a conexão e inicia a transação
                        if (_transManager.DbConnection.State == ConnectionState.Closed)
                            _transManager.DbConnection.Open();

                        _transManager.Transaction = _transManager.DbConnection.BeginTransaction();

                        ProcedureTransactionExecute(_inProcedure, out recordsAffected, _transManager);
                    }
                    //Permite a execução de + de três comandos em uma transação
                    else if (_transState == DBTransactionManager.TransactionState.ContinueExecution)
                    {
                        ProcedureTransactionExecute(_inProcedure, out recordsAffected, _transManager);
                    }
                    //Finaliza a transação.DBTransactionManager(false)
                    else
                    {
                        ProcedureTransactionExecute(_inProcedure, out recordsAffected, _transManager);

                        //Comita a transação
                        _transManager.Command.Transaction.Commit();

                        //Fecha a conexão da transação
                        _transManager.EndTransaction();
                    }
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção de banco ocorreu durante a execução dos comandos. Nenhuma das operações foi executada. <BR>";
                _errorMessage = _errorMessage + exp.Message;

                _transManager.Command.Transaction.Rollback();

                //Bloqueia a execução de qualquer chamada posterior dentro da transação em execução
                _transManager.TransHold = true;

                //Fecha a conexão da transação
                _transManager.EndTransaction();
            }
            catch (Exception exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução dos comandos. Nenhuma das operações foi executada. <BR>";
                _errorMessage = _errorMessage + exp.Message;

                _transManager.Command.Transaction.Rollback();

                //Bloqueia a execução de qualquer chamada posterior dentro da transação em execução
                _transManager.TransHold = true;

                //Fecha a conexão da transação
                _transManager.EndTransaction();
            }

            return recordsAffected;
        }

        /// <summary>
        /// Executa uma procedure com parâmetros, retornando um Int32, utilizando a conexão do banco e transações.
        /// </summary>
        /// <param name="_inProcedure">Procedure de Entrada.</param>
        /// <param name="_errorMessage">>Mensagem de Erro.</param>
        /// <param name="_transManager">Objeto DBTransactionManager da transação.</param>
        /// <param name="_transState">Estado de execução do bloco da transação. (Begin, End, ContinueExecution, StandAlone)</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>Int32 indicando a quantidade de linhas afetadas (INSERT, UPDATE ou DELETE).</returns>
        public static int InstanceProcedure(string _inProcedure, out string _errorMessage, ref DBTransactionManager _transManager, DBTransactionManager.TransactionState _transState, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();

            _errorMessage = String.Empty;
            int recordsAffected = 0;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return recordsAffected = -1;
            }

            try
            {
                //Executa uma transação isolada
                if (_transState == DBTransactionManager.TransactionState.StandAlone)
                {
                    //Abre a conexão
                    if (_transManager.DbConnection.State == ConnectionState.Closed)
                        _transManager.DbConnection.Open();

                    //Inicia a transação
                    _transManager.Transaction = _transManager.DbConnection.BeginTransaction();

                        ProcedureTransactionExecute(_inProcedure, out recordsAffected, _transManager, _params);

                    //Comita a transação
                    _transManager.Command.Transaction.Commit();

                    //Fecha a conexão da transação
                    _transManager.EndTransaction();
                }
                else
                {
                    //Bloqueio de execução depois da detecção de um erro
                    if (!_transManager.TransHold)
                    {
                        //Inicia a transação.DBTransactionManager(true)
                        if (_transState == DBTransactionManager.TransactionState.Begin)
                        {
                            //Abre a conexão e inicia a transação
                            if (_transManager.DbConnection.State == ConnectionState.Closed)
                                _transManager.DbConnection.Open();

                            _transManager.Transaction = _transManager.DbConnection.BeginTransaction();

                                ProcedureTransactionExecute(_inProcedure, out recordsAffected, _transManager, _params);
                        }
                        //Permite a execução de + de três comandos em uma transação
                        else if (_transState == DBTransactionManager.TransactionState.ContinueExecution)
                        {
                                ProcedureTransactionExecute(_inProcedure, out recordsAffected, _transManager, _params);
                        }
                        //Finaliza a transação.DBTransactionManager(false)
                        else if (_transState == DBTransactionManager.TransactionState.End)
                        {
                                ProcedureTransactionExecute(_inProcedure, out recordsAffected, _transManager, _params);

                            //Comita a transação
                            _transManager.Command.Transaction.Commit();

                            //Fecha a conexão da transação
                            _transManager.EndTransaction();
                        }
                    }
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção de banco ocorreu durante a execução dos comandos. Nenhuma das operações foi executada. <BR>";
                _errorMessage = _errorMessage + exp.Message;

                if (_transManager.Command.Transaction != null)
                    _transManager.Command.Transaction.Rollback();

                //Bloqueia a execução de qualquer chamada posterior dentro da transação em execução
                _transManager.TransHold = true;

                //Fecha a conexão da transação
                _transManager.EndTransaction();
            }
            catch (Exception exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução dos comandos. Nenhuma das operações foi executada. <BR>";
                _errorMessage = _errorMessage + exp.Message;

                if (_transManager.Command.Transaction != null)
                    _transManager.Command.Transaction.Rollback();

                //Bloqueia a execução de qualquer chamada posterior dentro da transação em execução
                _transManager.TransHold = true;

                //Fecha a conexão da transação
                _transManager.EndTransaction();
            }

            return recordsAffected;
        }

        //Métodos Privados para DBProvider.Instance

        /// <summary>
        ///  Método privado para execução de uma procedure usando uma transação.
        /// </summary>
        /// <param name="_inProcedure">Procedure de Entrada.</param>
        /// <param name="_recordsAffected">Int32 indicando a quantidade de linhas afetadas (INSERT, UPDATE ou DELETE).</param>
        /// <param name="_transManager">>Objeto DBTransactionManager da transação.</param>
        private static void ProcedureTransactionExecute(string _inProcedure, out int _recordsAffected, DBTransactionManager _transManager)
        {
            //Associa a transação ao comando
            _transManager.Command.Transaction = _transManager.Transaction;

            _transManager.Command.CommandText = _inProcedure;
            _transManager.Command.CommandType = CommandType.StoredProcedure;
            _transManager.Command.Connection = _transManager.DbConnection;

            //Executa o comando
            _recordsAffected = _transManager.Command.ExecuteNonQuery();
        }

        /// <summary>
        /// Método privado para execução de uma procedure usando uma transação.
        /// </summary>
        /// <param name="_inProcedure">Procedure de Entrada.</param>
        /// <param name="_recordsAffected">Int32 indicando a quantidade de linhas afetadas (INSERT, UPDATE ou DELETE).</param>
        /// <param name="_transManager">>Objeto DBTransactionManager da transação.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        private static void ProcedureTransactionExecute(string _inProcedure, out int _recordsAffected, DBTransactionManager _transManager, params DbParameter[] _params)
        {
            DbParameter parameter = null;

            //Associa a transação ao comando
            _transManager.Command.Transaction = _transManager.Transaction;

            _transManager.Command.CommandText = _inProcedure;
            _transManager.Command.CommandType = CommandType.StoredProcedure;
            _transManager.Command.Connection = _transManager.DbConnection;

            _transManager.Command.Parameters.Clear();

            //Adiciona os parâmetros de execução ao comando da transação
            parameter = _transManager.Factory.CreateParameter();
            for (int i = 0; i < _params.Length; i++)
            {
                parameter = _params[i];

                _transManager.Command.Parameters.Add(parameter);
            }

            //Executa o comando
            _recordsAffected = _transManager.Command.ExecuteNonQuery();
        }

        #endregion

        #region { InstanceDataSet }

        //Query InstanceDataSet

        /// <summary>
        /// Executa uma query sem parâmetros, retornando um DataSet, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <returns>System.Data.DataSet com as linhas da Query.</returns>
        public static DataSet InstanceDataSet(string _inQuery, out string _errorMessage)
        {
            DBInstance dbInstance = new DBInstance();
            DataSet dataSet = new DataSet();
            DbDataAdapter dataAdapter = null;

            _errorMessage = String.Empty;

            if (_inQuery.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataSet;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection();

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();
                dataAdapter = dbInstance.Factory.CreateDataAdapter();

                dbInstance.Command.CommandText = _inQuery;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection  = dbInstance.DbConnection;

                dataAdapter.SelectCommand = dbInstance.Command;
                dataAdapter.Fill(dataSet);
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dataAdapter != null)
                {
                    dataAdapter.Dispose();
                }
            }

            return (dataSet);
        }

        /// <summary>
        /// Executa uma query sem parâmetros, retornando um DataSet, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <returns>System.Data.DataSet com as linhas da Query.</returns>
        public static DataSet InstanceDataSet(string _connectionKey, string _inQuery, out string _errorMessage)
        {
            DBInstance dbInstance = new DBInstance();
            DataSet dataSet = new DataSet();
            DbDataAdapter dataAdapter = null;

            _errorMessage = String.Empty;

            if (_inQuery.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataSet;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();
                dataAdapter = dbInstance.Factory.CreateDataAdapter();

                dbInstance.Command.CommandText = _inQuery;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                dataAdapter.SelectCommand = dbInstance.Command;
                dataAdapter.Fill(dataSet);
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dataAdapter != null)
                {
                    dataAdapter.Dispose();
                }
            }

            return (dataSet);
        }

        /// <summary>
        /// Executa uma query com parâmetros, retornando um DataSet, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>System.Data.DataSet com as linhas da Query.</returns>
        public static DataSet InstanceDataSet(string _inQuery, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();
            DataSet dataSet = new DataSet();
            DbParameter parameter = null;
            DbDataAdapter dataAdapter = null;
            
            _errorMessage = String.Empty;

            if (_inQuery.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataSet;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection();

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();

                dbInstance.Command     = dbInstance.Factory.CreateCommand();
                dataAdapter = dbInstance.Factory.CreateDataAdapter();
                parameter   = dbInstance.Factory.CreateParameter();

                dbInstance.Command.CommandText = _inQuery;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection  = dbInstance.DbConnection;

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dataAdapter.SelectCommand = dbInstance.Command;
                dataAdapter.Fill(dataSet);
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dataAdapter != null)
                {
                    dataAdapter.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataSet);
        }

        /// <summary>
        /// Executa uma query com parâmetros, retornando um DataSet, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>System.Data.DataSet com as linhas da Query.</returns>
        public static DataSet InstanceDataSet(string _connectionKey, string _inQuery, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();
            DataSet dataSet = new DataSet();
            DbParameter parameter = null;
            DbDataAdapter dataAdapter = null;

            _errorMessage = String.Empty;

            if (_inQuery.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataSet;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();

                dbInstance.Command = dbInstance.Factory.CreateCommand();
                dataAdapter = dbInstance.Factory.CreateDataAdapter();
                parameter = dbInstance.Factory.CreateParameter();

                dbInstance.Command.CommandText = _inQuery;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dataAdapter.SelectCommand = dbInstance.Command;
                dataAdapter.Fill(dataSet);
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dataAdapter != null)
                {
                    dataAdapter.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataSet);
        }

        /// <summary>
        /// Executa uma query com parâmetros, retornando um DataSet, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Hashtable com o parâmetro.</param>
        /// <returns>System.Data.DataSet com as linhas da Query.</returns>
        public static DataSet InstanceDataSet(string _inQuery, out string _errorMessage, Hashtable _params)
        {
            DBInstance dbInstance = new DBInstance();
            DataSet dataSet = new DataSet();
            DbParameter parameter     = null;
            DbDataAdapter dataAdapter = null;

            _errorMessage = String.Empty;

            if (_inQuery.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataSet;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection();

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();

                dbInstance.Command = dbInstance.Factory.CreateCommand();
                dataAdapter = dbInstance.Factory.CreateDataAdapter();
                parameter = dbInstance.Factory.CreateParameter();

                dbInstance.Command.CommandText = _inQuery;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                foreach (DictionaryEntry de in _params)
                {
                    parameter.ParameterName = de.Key.ToString();
                    parameter.Value = de.Value;

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dataAdapter.SelectCommand = dbInstance.Command;
                dataAdapter.Fill(dataSet);
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dataAdapter != null)
                {
                    dataAdapter.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }
            
            return (dataSet);
        }

        /// <summary>
        /// Executa uma query com parâmetros, retornando um DataSet, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Hashtable com o parâmetro.</param>
        /// <returns>System.Data.DataSet com as linhas da Query.</returns>
        public static DataSet InstanceDataSet(string _connectionKey, string _inQuery, out string _errorMessage, Hashtable _params)
        {
            DBInstance dbInstance = new DBInstance();
            DataSet dataSet = new DataSet();
            DbParameter parameter = null;
            DbDataAdapter dataAdapter = null;

            _errorMessage = String.Empty;

            if (_inQuery.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataSet;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection();

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();

                dbInstance.Command = dbInstance.Factory.CreateCommand();
                dataAdapter = dbInstance.Factory.CreateDataAdapter();
                parameter = dbInstance.Factory.CreateParameter();

                dbInstance.Command.CommandText = _inQuery;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                foreach (DictionaryEntry de in _params)
                {
                    parameter.ParameterName = de.Key.ToString();
                    parameter.Value = de.Value;

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dataAdapter.SelectCommand = dbInstance.Command;
                dataAdapter.Fill(dataSet);
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dataAdapter != null)
                {
                    dataAdapter.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataSet);
        }

        //Procedure InstanceProcedureDataSet

        /// <summary>
        /// Executa uma procedure sem parâmetros, retornando um DataSet, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_inProcedure">Nome da StoredProcedure.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <returns>System.Data.DataSet com as linhas da StoredProcedure.</returns>
        public static DataSet InstanceProcedureDataSet(string _inProcedure, out string _errorMessage)
        {
            DBInstance dbInstance = new DBInstance();
            DataSet dataSet = new DataSet();
            DbDataAdapter dataAdapter = null;

            _errorMessage = String.Empty;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataSet;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection();

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();
                dataAdapter = dbInstance.Factory.CreateDataAdapter();

                dbInstance.Command.CommandText = _inProcedure;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                dataAdapter.SelectCommand = dbInstance.Command;
                dataAdapter.Fill(dataSet);
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dataAdapter != null)
                {
                    dataAdapter.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataSet);
        }

        /// <summary>
        /// Executa uma procedure sem parâmetros, retornando um DataSet, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inProcedure">Nome da StoredProcedure.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <returns>System.Data.DataSet com as linhas da StoredProcedure.</returns>
        public static DataSet InstanceProcedureDataSet(string _connectionKey, string _inProcedure, out string _errorMessage)
        {
            DBInstance dbInstance = new DBInstance();
            DataSet dataSet = new DataSet();
            DbDataAdapter dataAdapter = null;

            _errorMessage = String.Empty;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataSet;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();
                dataAdapter = dbInstance.Factory.CreateDataAdapter();

                dbInstance.Command.CommandText = _inProcedure;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                dataAdapter.SelectCommand = dbInstance.Command;
                dataAdapter.Fill(dataSet);
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dataAdapter != null)
                {
                    dataAdapter.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataSet);
        }

        /// <summary>
        /// Executa uma query com parâmetros, retornando um DataSet, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_inProcedure">Nome da StoredProcedure.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>System.Data.DataSet com as linhas da StoredProcedure.</returns>
        public static DataSet InstanceProcedureDataSet(string _inProcedure, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();
            DataSet dataSet = new DataSet();
            DbParameter parameter     = null;
            DbDataAdapter dataAdapter = null;

            _errorMessage = String.Empty;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataSet;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection();

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();

                dbInstance.Command = dbInstance.Factory.CreateCommand();
                dataAdapter = dbInstance.Factory.CreateDataAdapter();
                parameter = dbInstance.Factory.CreateParameter();

                dbInstance.Command.CommandText = _inProcedure;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dataAdapter.SelectCommand = dbInstance.Command;
                dataAdapter.Fill(dataSet);
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dataAdapter != null)
                {
                    dataAdapter.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataSet);
        }

        /// <summary>
        /// Executa uma query com parâmetros, retornando um DataSet, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inProcedure">Nome da StoredProcedure.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>System.Data.DataSet com as linhas da StoredProcedure.</returns>
        public static DataSet InstanceProcedureDataSet(string _connectionKey, string _inProcedure, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();
            DataSet dataSet = new DataSet();
            DbParameter parameter = null;
            DbDataAdapter dataAdapter = null;

            _errorMessage = String.Empty;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataSet;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();

                dbInstance.Command = dbInstance.Factory.CreateCommand();
                dataAdapter = dbInstance.Factory.CreateDataAdapter();
                parameter = dbInstance.Factory.CreateParameter();

                dbInstance.Command.CommandText = _inProcedure;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dataAdapter.SelectCommand = dbInstance.Command;
                dataAdapter.Fill(dataSet);
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dataAdapter != null)
                {
                    dataAdapter.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataSet);
        }

        /// <summary>
        /// Executa uma query com parâmetros, retornando um DataSet, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_inProcedure">Nome da StoredProcedure.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_hashTable">HashTable com os parâmetros de saída.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>System.Data.DataSet com as linhas da StoredProcedure.</returns>
        public static DataSet InstanceProcedureDataSet(string _inProcedure, out string _errorMessage, out Hashtable _hashTable , params DbParameter[] _params)
        {
            _hashTable = new Hashtable();
            DBInstance dbInstance = new DBInstance();
            DataSet dataSet = new DataSet();
            DbParameter parameter = null;
            DbDataAdapter dataAdapter = null;

            _errorMessage = String.Empty;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataSet;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection();

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();

                dbInstance.Command = dbInstance.Factory.CreateCommand();
                dataAdapter = dbInstance.Factory.CreateDataAdapter();
                parameter   = dbInstance.Factory.CreateParameter();

                dbInstance.Command.CommandText = _inProcedure;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection  = dbInstance.DbConnection;

                //Adiciona os parâmetros de entrada
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);

                }

                dataAdapter.SelectCommand = dbInstance.Command;
                dataAdapter.Fill(dataSet);

                //Recupera os valores de saída e insere numa HashTable
                for (int i = 0; i < _params.Length; i++)
                {
                    _hashTable.Add(_params[i].ParameterName, _params[i].Value); //Valores de saída
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dataAdapter != null)
                {
                    dataAdapter.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataSet);
        }

        /// <summary>
        /// Executa uma query com parâmetros, retornando um DataSet, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inProcedure">Nome da StoredProcedure.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_hashTable">HashTable com os parâmetros de saída.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>System.Data.DataSet com as linhas da StoredProcedure.</returns>
        public static DataSet InstanceProcedureDataSet(string _connectionKey, string _inProcedure, out string _errorMessage, out Hashtable _hashTable, params DbParameter[] _params)
        {
            _hashTable = new Hashtable();
            DBInstance dbInstance = new DBInstance();
            DataSet dataSet = new DataSet();
            DbParameter parameter = null;
            DbDataAdapter dataAdapter = null;

            _errorMessage = String.Empty;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataSet;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();

                dbInstance.Command = dbInstance.Factory.CreateCommand();
                dataAdapter = dbInstance.Factory.CreateDataAdapter();
                parameter = dbInstance.Factory.CreateParameter();

                dbInstance.Command.CommandText = _inProcedure;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                //Adiciona os parâmetros de entrada
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);

                }

                dataAdapter.SelectCommand = dbInstance.Command;
                dataAdapter.Fill(dataSet);

                //Recupera os valores de saída e insere numa HashTable
                for (int i = 0; i < _params.Length; i++)
                {
                    _hashTable.Add(_params[i].ParameterName, _params[i].Value); //Valores de saída
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dataAdapter != null)
                {
                    dataAdapter.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataSet);
        }

        /// <summary>
        /// Executa uma query com parâmetros, retornando um DataSet, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_inProcedure">Nome da StoredProcedure.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Hashtable com o parâmetro.</param>
        /// <returns>System.Data.DataSet com as linhas da StoredProcedure.</returns>
        public static DataSet InstanceProcedureDataSet(string _inProcedure, out string _errorMessage, Hashtable _params)
        {
            DBInstance dbInstance = new DBInstance();
            DataSet dataSet = new DataSet();
            DbParameter parameter     = null;
            DbDataAdapter dataAdapter = null;

            _errorMessage = String.Empty;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataSet;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection();

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();

                dbInstance.Command = dbInstance.Factory.CreateCommand();
                dataAdapter = dbInstance.Factory.CreateDataAdapter();
                parameter = dbInstance.Factory.CreateParameter();

                dbInstance.Command.CommandText = _inProcedure;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                foreach (DictionaryEntry de in _params)
                {
                    parameter.ParameterName = de.Key.ToString();
                    parameter.Value = de.Value;

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dataAdapter.SelectCommand = dbInstance.Command;
                dataAdapter.Fill(dataSet);
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dataAdapter != null)
                {
                    dataAdapter.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataSet);
        }

        /// <summary>
        /// Executa uma query com parâmetros, retornando um DataSet, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inProcedure">Nome da StoredProcedure.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Hashtable com o parâmetro.</param>
        /// <returns>System.Data.DataSet com as linhas da StoredProcedure.</returns>
        public static DataSet InstanceProcedureDataSet(string _connectionKey, string _inProcedure, out string _errorMessage, Hashtable _params)
        {
            DBInstance dbInstance = new DBInstance();
            DataSet dataSet = new DataSet();
            DbParameter parameter = null;
            DbDataAdapter dataAdapter = null;

            _errorMessage = String.Empty;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataSet;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();

                dbInstance.Command = dbInstance.Factory.CreateCommand();
                dataAdapter = dbInstance.Factory.CreateDataAdapter();
                parameter = dbInstance.Factory.CreateParameter();

                dbInstance.Command.CommandText = _inProcedure;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                foreach (DictionaryEntry de in _params)
                {
                    parameter.ParameterName = de.Key.ToString();
                    parameter.Value = de.Value;

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dataAdapter.SelectCommand = dbInstance.Command;
                dataAdapter.Fill(dataSet);
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dataAdapter != null)
                {
                    dataAdapter.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataSet);
        }

        #endregion

        #region { InstanceDataTable }

        //Query InstanceDataTable

        /// <summary>
        /// Executa uma query sem parâmetros, retornando uma DataTable, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <returns>System.Data.DataTable com as linhas da Query.</returns>
        public static DataTable InstanceDataTable(string _inQuery, out string _errorMessage)
        {
            DBInstance dbInstance = new DBInstance();
            DataTable dataTable = new DataTable();
            DbDataAdapter dataAdapter = null;
            DbDataReader dataReader   = null;

            _errorMessage = String.Empty;

            if (_inQuery.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataTable;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection();

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();
                dataAdapter = dbInstance.Factory.CreateDataAdapter();

                dbInstance.Command.CommandText = _inQuery;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection  = dbInstance.DbConnection;

                dbInstance.Command.Connection.Open();

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);

                if (dataReader.HasRows)
                {
                    dataTable.Load(dataReader);
                }
                else
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    return null;
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dataAdapter != null)
                {
                    dataAdapter.Dispose();
                }
                if (dataReader != null)
                {
                    dataReader.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataTable);
        }

        /// <summary>
        /// Executa uma query sem parâmetros, retornando uma DataTable, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <returns>System.Data.DataTable com as linhas da Query.</returns>
        public static DataTable InstanceDataTable(string _connectionKey, string _inQuery, out string _errorMessage)
        {
            DBInstance dbInstance = new DBInstance();
            DataTable dataTable = new DataTable();
            DbDataAdapter dataAdapter = null;
            DbDataReader dataReader = null;

            _errorMessage = String.Empty;

            if (_inQuery.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataTable;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();
                dataAdapter = dbInstance.Factory.CreateDataAdapter();

                dbInstance.Command.CommandText = _inQuery;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                dbInstance.Command.Connection.Open();

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);

                if (dataReader.HasRows)
                {
                    dataTable.Load(dataReader);
                }
                else
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    return null;
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dataAdapter != null)
                {
                    dataAdapter.Dispose();
                }
                if (dataReader != null)
                {
                    dataReader.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataTable);
        }

        /// <summary>
        /// Executa uma query sem parâmetros, retornando uma DataTable, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_providerKey">Chave do Provider do AppSettings do web.config.</param>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <returns>System.Data.DataTable com as linhas da Query.</returns>
        public static DataTable InstanceDataTable(string _providerKey, string _connectionKey, string _inQuery, out string _errorMessage)
        {
            DBInstance dbInstance = new DBInstance();
            DataTable dataTable = new DataTable();
            DbDataReader dataReader = null;

            _errorMessage = String.Empty;

            if (_inQuery.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataTable;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_providerKey, _connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory(_providerKey);
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inQuery;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                dbInstance.Command.Connection.Open();

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);

                if (dataReader.HasRows)
                {
                    dataTable.Load(dataReader);
                }
                else
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    return null;
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dataReader != null)
                {
                    dataReader.Dispose();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataTable);
        }

        /// <summary>
        /// Executa uma query sem parâmetros, retornando uma DataTable, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>System.Data.DataTable com as linhas da Query.</returns>
        public static DataTable InstanceDataTable(string _inQuery, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();
            DataTable dataTable = new DataTable();
            DbDataAdapter dataAdapter = null;
            DbDataReader dataReader   = null;
            DbParameter parameter     = null;

            _errorMessage = String.Empty;

            if (_inQuery.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataTable;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection();

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();
                dataAdapter = dbInstance.Factory.CreateDataAdapter();

                dbInstance.Command.CommandText = _inQuery;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                dbInstance.Command.Connection.Open();

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);

                if (dataReader.HasRows)
                {
                    dataTable.Load(dataReader);
                }
                else
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    return null;
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dataAdapter != null)
                {
                    dataAdapter.Dispose();
                }
                if (dataReader != null)
                {
                    dataReader.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataTable);
        }

        /// <summary>
        /// Executa uma query sem parâmetros, retornando uma DataTable, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>System.Data.DataTable com as linhas da Query.</returns>
        public static DataTable InstanceDataTable(string _connectionKey, string _inQuery, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();
            DataTable dataTable = new DataTable();
            DbDataAdapter dataAdapter = null;
            DbDataReader dataReader = null;
            DbParameter parameter = null;

            _errorMessage = String.Empty;

            if (_inQuery.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataTable;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();
                dataAdapter = dbInstance.Factory.CreateDataAdapter();

                dbInstance.Command.CommandText = _inQuery;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                dbInstance.Command.Connection.Open();

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);

                if (dataReader.HasRows)
                {
                    dataTable.Load(dataReader);
                }
                else
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    return null;
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dataAdapter != null)
                {
                    dataAdapter.Dispose();
                }
                if (dataReader != null)
                {
                    dataReader.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataTable);
        }

        /// <summary>
        /// Executa uma query sem parâmetros, retornando uma DataTable, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_providerKey">Chave do Provider do AppSettings do web.config.</param>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>System.Data.DataTable com as linhas da Query.</returns>
        public static DataTable InstanceDataTable(string _providerKey, string _connectionKey, string _inQuery, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();
            DataTable dataTable = new DataTable();
            DbDataReader dataReader = null;
            DbParameter parameter = null;

            _errorMessage = String.Empty;

            if (_inQuery.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataTable;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_providerKey, _connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory(_providerKey);
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inQuery;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                dbInstance.Command.Connection.Open();

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);

                if (dataReader.HasRows)
                {
                    dataTable.Load(dataReader);
                }
                else
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    return null;
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dataReader != null)
                {
                    dataReader.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataTable);
        }

        /// <summary>
        /// Executa uma query sem parâmetros, retornando uma DataTable, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>System.Data.DataTable com as linhas da Query.</returns>
        public static DataTable InstanceDataTable(string _inQuery, out string _errorMessage, Hashtable _params)
        {
            DBInstance dbInstance = new DBInstance();
            DataTable dataTable = new DataTable();
            DbDataAdapter dataAdapter = null;
            DbDataReader dataReader   = null;
            DbParameter parameter     = null;

            _errorMessage = String.Empty;

            if (_inQuery.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataTable;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection();

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();
                dataAdapter = dbInstance.Factory.CreateDataAdapter();

                dbInstance.Command.CommandText = _inQuery;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                foreach (DictionaryEntry de in _params)
                {
                    parameter.ParameterName = de.Key.ToString();
                    parameter.Value = de.Value;

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dbInstance.Command.Connection.Open();

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);

                if (dataReader.HasRows)
                {
                    dataTable.Load(dataReader);
                }
                else
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    return null;
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dataAdapter != null)
                {
                    dataAdapter.Dispose();
                }
                if (dataReader != null)
                {
                    dataReader.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataTable);
        }

        /// <summary>
        /// Executa uma query sem parâmetros, retornando uma DataTable, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>System.Data.DataTable com as linhas da Query.</returns>
        public static DataTable InstanceDataTable(string _connectionKey, string _inQuery, out string _errorMessage, Hashtable _params)
        {
            DBInstance dbInstance = new DBInstance();
            DataTable dataTable = new DataTable();
            DbDataAdapter dataAdapter = null;
            DbDataReader dataReader = null;
            DbParameter parameter = null;

            _errorMessage = String.Empty;

            if (_inQuery.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataTable;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();
                dataAdapter = dbInstance.Factory.CreateDataAdapter();

                dbInstance.Command.CommandText = _inQuery;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                foreach (DictionaryEntry de in _params)
                {
                    parameter.ParameterName = de.Key.ToString();
                    parameter.Value = de.Value;

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dbInstance.Command.Connection.Open();

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);

                if (dataReader.HasRows)
                {
                    dataTable.Load(dataReader);
                }
                else
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    return null;
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dataAdapter != null)
                {
                    dataAdapter.Dispose();
                }
                if (dataReader != null)
                {
                    dataReader.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataTable);
        }

        //Procedure InstanceDataTable

        /// <summary>
        /// Executa uma procedure com parâmetros, retornando uma DataTable, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_inProcedure">Procedure de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>System.Data.DataTable com as linhas da Query.</returns>
        public static DataTable InstanceProcedureDataTable(string _inProcedure, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();
            DataTable dataTable = new DataTable();
            DbDataAdapter dataAdapter = null;
            DbDataReader dataReader = null;
            DbParameter parameter = null;

            _errorMessage = String.Empty;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataTable;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection();

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();
                dataAdapter = dbInstance.Factory.CreateDataAdapter();

                dbInstance.Command.CommandText = _inProcedure;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dbInstance.Command.Connection.Open();

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);

                if (dataReader.HasRows)
                {
                    dataTable.Load(dataReader);
                }
                else
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    return null;
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dataAdapter != null)
                {
                    dataAdapter.Dispose();
                }
                if (dataReader != null)
                {
                    dataReader.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataTable);
        }

        /// <summary>
        /// Executa uma procedure com parâmetros, retornando uma DataTable, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inProcedure">Procedure de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>System.Data.DataTable com as linhas da Query.</returns>
        public static DataTable InstanceProcedureDataTable(string _connectionKey, string _inProcedure, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();
            DataTable dataTable = new DataTable();
            DbDataAdapter dataAdapter = null;
            DbDataReader dataReader = null;
            DbParameter parameter = null;

            _errorMessage = String.Empty;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataTable;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();
                dataAdapter = dbInstance.Factory.CreateDataAdapter();

                dbInstance.Command.CommandText = _inProcedure;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dbInstance.Command.Connection.Open();

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);

                if (dataReader.HasRows)
                {
                    dataTable.Load(dataReader);
                }
                else
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    return null;
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dataAdapter != null)
                {
                    dataAdapter.Dispose();
                }
                if (dataReader != null)
                {
                    dataReader.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataTable);
        }
  
        /// <summary>
        /// Executa uma procedure com parâmetros, retornando uma DataTable, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_providerKey">Chave do Provider do AppSettings do web.config.</param>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inProcedure">Procedure de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>System.Data.DataTable com as linhas da Query.</returns>
        public static DataTable InstanceProcedureDataTable(string _providerKey, string _connectionKey, string _inProcedure, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();
            DataTable dataTable = new DataTable();
            DbDataReader dataReader = null;
            DbParameter parameter = null;

            _errorMessage = String.Empty;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataTable;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_providerKey,_connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory(_providerKey);
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inProcedure;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dbInstance.Command.Connection.Open();

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);

                if (dataReader.HasRows)
                {
                    dataTable.Load(dataReader);
                }
                else
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    return null;
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dataReader != null)
                {
                    dataReader.Dispose();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataTable);
        }

        /// <summary>
        /// Executa uma procedure com parâmetros, retornando uma DataTable, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_inProcedure">Procedure de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>System.Data.DataTable com as linhas da Query.</returns>
        public static DataTable InstanceProcedureDataTable(string _inProcedure, out string _errorMessage, Hashtable _params)
        {
            DBInstance dbInstance = new DBInstance();
            DataTable dataTable = new DataTable();
            DbDataAdapter dataAdapter = null;
            DbDataReader dataReader = null;
            DbParameter parameter = null;

            _errorMessage = String.Empty;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataTable;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection();

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();
                dataAdapter = dbInstance.Factory.CreateDataAdapter();

                dbInstance.Command.CommandText = _inProcedure;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                foreach (DictionaryEntry de in _params)
                {
                    parameter.ParameterName = de.Key.ToString();
                    parameter.Value = de.Value;

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dbInstance.Command.Connection.Open();

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);

                if (dataReader.HasRows)
                {
                    dataTable.Load(dataReader);
                }
                else
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    return null;
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dataAdapter != null)
                {
                    dataAdapter.Dispose();
                }
                if (dataReader != null)
                {
                    dataReader.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataTable);
        }

        /// <summary>
        /// Executa uma procedure com parâmetros, retornando uma DataTable, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inProcedure">Procedure de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>System.Data.DataTable com as linhas da Query.</returns>
        public static DataTable InstanceProcedureDataTable(string _connectionKey, string _inProcedure, out string _errorMessage, Hashtable _params)
        {
            DBInstance dbInstance = new DBInstance();
            DataTable dataTable = new DataTable();
            DbDataAdapter dataAdapter = null;
            DbDataReader dataReader = null;
            DbParameter parameter = null;

            _errorMessage = String.Empty;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataTable;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();
                dataAdapter = dbInstance.Factory.CreateDataAdapter();

                dbInstance.Command.CommandText = _inProcedure;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                foreach (DictionaryEntry de in _params)
                {
                    parameter.ParameterName = de.Key.ToString();
                    parameter.Value = de.Value;

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dbInstance.Command.Connection.Open();

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);

                if (dataReader.HasRows)
                {
                    dataTable.Load(dataReader);
                }
                else
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    return null;
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dataAdapter != null)
                {
                    dataAdapter.Dispose();
                }
                if (dataReader != null)
                {
                    dataReader.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataTable);
        }

        /// <summary>
        /// Executa uma procedure com parâmetros, retornando uma DataTable, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_providerKey">Chave do Provider do AppSettings do web.config.</param>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inProcedure">Procedure de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>System.Data.DataTable com as linhas da Query.</returns>
        public static DataTable InstanceProcedureDataTable(string _providerKey, string _connectionKey, string _inProcedure, out string _errorMessage, Hashtable _params)
        {
            DBInstance dbInstance = new DBInstance();
            DataTable dataTable = new DataTable();
            DbDataReader dataReader = null;
            DbParameter parameter = null;

            _errorMessage = String.Empty;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataTable;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inProcedure;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                foreach (DictionaryEntry de in _params)
                {
                    parameter.ParameterName = de.Key.ToString();
                    parameter.Value = de.Value;

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dbInstance.Command.Connection.Open();

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);

                if (dataReader.HasRows)
                {
                    dataTable.Load(dataReader);
                }
                else
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    return null;
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dataReader != null)
                {
                    dataReader.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (dataTable);
        }

        //Procedure InstanceProcedureDataTable

        /// <summary>
        /// Executa uma procedure com parâmetros, retornando uma List, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_inProcedure">Procedure de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>System.Collections.Generic.List(DataRow) com as linhas da Procedure.</returns>
        public static List<DataRow> InstanceProcedureDataTableOutList(string _inProcedure, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();
            DataTable dataTable = new DataTable();
            DbDataAdapter dataAdapter = null;
            DbDataReader dataReader = null;
            DbParameter parameter = null;
            List<DataRow> list = null;

            _errorMessage = String.Empty;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return list;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection();

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();
                dataAdapter = dbInstance.Factory.CreateDataAdapter();

                dbInstance.Command.CommandText = _inProcedure;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dbInstance.Command.Connection.Open();

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);

                if (dataReader.HasRows)
                {
                    dataTable.Load(dataReader);
                    list = new List<DataRow>(dataTable.Select());
                }
                else
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    return null;
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dataAdapter != null)
                {
                    dataAdapter.Dispose();
                }
                if (dataReader != null)
                {
                    dataReader.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (list);
        }

        /// <summary>
        /// Executa uma procedure com parâmetros, retornando uma List, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inProcedure">Procedure de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>System.Collections.Generic.List(DataRow) com as linhas da Procedure.</returns>
        public static List<DataRow> InstanceProcedureDataTableOutList(string _connectionKey, string _inProcedure, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();
            DataTable dataTable = new DataTable();
            DbDataAdapter dataAdapter = null;
            DbDataReader dataReader = null;
            DbParameter parameter = null;
            List<DataRow> list = null;

            _errorMessage = String.Empty;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return list;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();
                dataAdapter = dbInstance.Factory.CreateDataAdapter();

                dbInstance.Command.CommandText = _inProcedure;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dbInstance.Command.Connection.Open();

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);

                if (dataReader.HasRows)
                {
                    dataTable.Load(dataReader);
                    list = new List<DataRow>(dataTable.Select());
                }
                else
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    return null;
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dataAdapter != null)
                {
                    dataAdapter.Dispose();
                }
                if (dataReader != null)
                {
                    dataReader.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (list);
        }

        /// <summary>
        /// Executa uma procedure com parâmetros, retornando uma List, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_providerKey">Chave do Provider do AppSettings do web.config.</param>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inProcedure">Procedure de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>System.Collections.Generic.List(DataRow) com as linhas da Procedure.</returns>
        public static List<DataRow> InstanceProcedureDataTableOutList(string _providerKey, string _connectionKey, string _inProcedure, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();
            DataTable dataTable = new DataTable();
            DbDataReader dataReader = null;
            DbParameter parameter = null;
            List<DataRow> list = null;

            _errorMessage = String.Empty;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return list;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_providerKey, _connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory(_providerKey);
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inProcedure;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dbInstance.Command.Connection.Open();

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);

                if (dataReader.HasRows)
                {
                    dataTable.Load(dataReader);
                    list = new List<DataRow>(dataTable.Select());
                }
                else
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    return null;
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dataReader != null)
                {
                    dataReader.Dispose();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return (list);
        }

        #endregion

        #region { InstanceDataReader }
        //IMPORTANTE:O objeto DataReader deve ser fechado explicitamente pelo programador depois de utilizado, seja na camada de apresentação ou na camada de negócio.

        //Query InstanceDataReader

        /// <summary>
        /// Executa uma query sem parâmetros, retornando uma DataReader, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <returns>System.Data.DataReader com as linhas da Query.</returns>
        public static DbDataReader InstanceDataReader(string _inQuery, out string _errorMessage)
        {
            DBInstance dbInstance = new DBInstance();
            DbDataReader dataReader = null;

            _errorMessage = String.Empty;

            if (_inQuery.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataReader;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection();

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inQuery;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection  = dbInstance.DbConnection;
                
                dbInstance.Command.Connection.Open();

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);

                if (!dataReader.HasRows)
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    dbInstance.Command.Connection.Close();

                    return null;
                }

            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
            }

            return (dataReader);
        }

        /// <summary>
        /// Executa uma query sem parâmetros, retornando uma DataReader, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <returns>System.Data.DataReader com as linhas da Query.</returns>
        public static DbDataReader InstanceDataReader(string _connectionKey, string _inQuery, out string _errorMessage)
        {
            DBInstance dbInstance = new DBInstance();
            DbDataReader dataReader = null;

            _errorMessage = String.Empty;

            if (_inQuery.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataReader;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inQuery;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                dbInstance.Command.Connection.Open();

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);

                if (!dataReader.HasRows)
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    dbInstance.Command.Connection.Close();

                    return null;
                }

            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
            }

            return (dataReader);
        }

        /// <summary>
        /// Executa uma query sem parâmetros, retornando uma DataReader, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_providerKey">Chave do Provider do AppSettings do web.config.</param>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <returns>System.Data.DataReader com as linhas da Query.</returns>
        public static DbDataReader InstanceDataReader(string _providerKey, string _connectionKey, string _inQuery, out string _errorMessage)
        {
            DBInstance dbInstance = new DBInstance();
            DbDataReader dataReader = null;

            _errorMessage = String.Empty;

            if (_inQuery.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataReader;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_providerKey, _connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory(_providerKey);
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inQuery;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                dbInstance.Command.Connection.Open();

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);

                if (!dataReader.HasRows)
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    dbInstance.Command.Connection.Close();

                    return null;
                }

            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
            }

            return (dataReader);
        }

        /// <summary>
        /// Executa uma query com parâmetros, retornando uma DataReader, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <returns>System.Data.DataReader com as linhas da Query.</returns>
        public static DbDataReader InstanceDataReader(string _inQuery, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();
            DbDataReader dataReader = null;
            DbParameter parameter = null;

            _errorMessage = String.Empty;

            if (_inQuery.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataReader;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection();

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inQuery;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                dbInstance.Command.Connection.Open();

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);
                if (!dataReader.HasRows)
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    dbInstance.Command.Connection.Close();

                    return null;
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
            }

            return (dataReader);
        }

        /// <summary>
        /// Executa uma query com parâmetros, retornando uma DataReader, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <returns>System.Data.DataReader com as linhas da Query.</returns>
        public static DbDataReader InstanceDataReader(string _connectionKey, string _inQuery, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();
            DbDataReader dataReader = null;
            DbParameter parameter = null;

            _errorMessage = String.Empty;

            if (_inQuery.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataReader;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inQuery;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                dbInstance.Command.Connection.Open();

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);
                if (!dataReader.HasRows)
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    dbInstance.Command.Connection.Close();

                    return null;
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
            }

            return (dataReader);
        }

        /// <summary>
        /// Executa uma query com parâmetros, retornando uma DataReader, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_providerKey">Chave do Provider do AppSettings do web.config.</param>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <returns>System.Data.DataReader com as linhas da Query.</returns>
        public static DbDataReader InstanceDataReader(string _providerKey, string _connectionKey, string _inQuery, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();
            DbDataReader dataReader = null;
            DbParameter parameter = null;

            _errorMessage = String.Empty;

            if (_inQuery.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataReader;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_providerKey, _connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory(_providerKey);
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inQuery;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                dbInstance.Command.Connection.Open();

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);
                if (!dataReader.HasRows)
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    dbInstance.Command.Connection.Close();

                    return null;
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
            }

            return (dataReader);
        }

        //Procedure InstanceProcedureDataReader

        /// <summary>
        /// Executa uma procedure sem parâmetros, retornando uma DataReader, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_inProcedure">Nome da StoredProcedure.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <returns>System.Data.DataReader com as linhas da StoredProcedure.</returns>
        public static DbDataReader InstanceProcedureDataReader(string _inProcedure, out string _errorMessage)
        {
            DBInstance dbInstance = new DBInstance();
            DbDataReader dataReader = null;

            _errorMessage = String.Empty;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataReader;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection();

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inProcedure;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                dbInstance.Command.Connection.Open();

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);
                if (!dataReader.HasRows)
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";
                    
                    dbInstance.Command.Connection.Close();

                    return null;
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
            }

            return (dataReader);
        }

        /// <summary>
        /// Executa uma procedure sem parâmetros, retornando uma DataReader, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inProcedure">Nome da StoredProcedure.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <returns>System.Data.DataReader com as linhas da StoredProcedure.</returns>
        public static DbDataReader InstanceProcedureDataReader(string _connectionKey, string _inProcedure, out string _errorMessage)
        {
            DBInstance dbInstance = new DBInstance();
            DbDataReader dataReader = null;

            _errorMessage = String.Empty;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataReader;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inProcedure;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                dbInstance.Command.Connection.Open();

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);
                if (!dataReader.HasRows)
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    dbInstance.Command.Connection.Close();

                    return null;
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
            }

            return (dataReader);
        }

        /// <summary>
        /// Executa uma procedure com parâmetros, retornando uma DataReader, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_inProcedure">Nome da StoredProcedure.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <returns>System.Data.DataReader com as linhas da StoredProcedure.</returns>
        public static DbDataReader InstanceProcedureDataReader(string _inProcedure, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();
            DbDataReader dataReader = null;
            DbParameter parameter = null;

            _errorMessage = String.Empty;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataReader;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection();

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inProcedure;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dbInstance.Command.Connection.Open();

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);
                if (!dataReader.HasRows)
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    dbInstance.Command.Connection.Close();

                    return null;
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
            }

            return (dataReader);
        }

        /// <summary>
        /// Executa uma procedure com parâmetros, retornando uma DataReader, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_providerKey">Chave do Provider do AppSettings do web.config.</param>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inProcedure">Nome da StoredProcedure.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <returns>System.Data.DataReader com as linhas da StoredProcedure.</returns>
        public static DbDataReader InstanceProcedureDataReader(string _providerKey, string _connectionKey, string _inProcedure, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();
            DbDataReader dataReader = null;
            DbParameter parameter = null;

            _errorMessage = String.Empty;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataReader;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_providerKey, _connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory(_providerKey);
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inProcedure;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dbInstance.Command.Connection.Open();

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);
                if (!dataReader.HasRows)
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    dbInstance.Command.Connection.Close();

                    return null;
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
            }

            return (dataReader);
        }

        /// <summary>
        /// Executa uma procedure com parâmetros, retornando uma DataReader, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inProcedure">Nome da StoredProcedure.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <returns>System.Data.DataReader com as linhas da StoredProcedure.</returns>
        public static DbDataReader InstanceProcedureDataReader(string _connectionKey, string _inProcedure, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();
            DbDataReader dataReader = null;
            DbParameter parameter = null;

            _errorMessage = String.Empty;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return dataReader;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inProcedure;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dbInstance.Command.Connection.Open();

                dataReader = dbInstance.Command.ExecuteReader(CommandBehavior.CloseConnection);
                if (!dataReader.HasRows)
                {
                    _errorMessage = "Nenhum registro encontrado.  <BR>";

                    dbInstance.Command.Connection.Close();

                    return null;
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
            }

            return (dataReader);
        }

        #endregion

        #region { InstanceQueryOutData }

        //Query InstanceQueryOutData

        /// <summary>
        /// Executa uma query com parâmetros, retornando um Int32, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>Hashtable.</returns>
        public static Hashtable InstanceQueryOutData(string _inQuery, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();
            Hashtable ht = new Hashtable();
            DbParameter parameter = null;

            _errorMessage = String.Empty;

            if (_inQuery.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return null;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection();

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inQuery;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                parameter = dbInstance.Factory.CreateParameter();

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dbInstance.DbConnection.Open();

                dbInstance.Command.ExecuteNonQuery();

                dbInstance.DbConnection.Close();

                //Recupera os valores de saída e insere numa HashTable
                for (int i = 0; i < _params.Length; i++)
                {
                    ht.Add(_params[i].ParameterName, _params[i].Value);
                }

                return ht;
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return ht;

        }

        /// <summary>
        /// Executa uma query com parâmetros, retornando um Int32, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>Hashtable.</returns>
        public static Hashtable InstanceQueryOutData(string _connectionKey, string _inQuery, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();
            Hashtable ht = new Hashtable();
            DbParameter parameter = null;

            _errorMessage = String.Empty;

            if (_inQuery.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return null;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inQuery;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                parameter = dbInstance.Factory.CreateParameter();

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dbInstance.DbConnection.Open();

                dbInstance.Command.ExecuteNonQuery();

                dbInstance.DbConnection.Close();

                //Recupera os valores de saída e insere numa HashTable
                for (int i = 0; i < _params.Length; i++)
                {
                    ht.Add(_params[i].ParameterName, _params[i].Value);
                }

                return ht;
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return ht;

        }

        //Procedure InstanceProcedureOutData

        /// <summary>
        /// Executa uma procedure com parâmetros, retornando um Int32, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_inProcedure">Nome da StoredProcedure.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>Hashtable.</returns>
        public static Hashtable InstanceProcedureOutData(string _inProcedure, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();
            Hashtable ht = new Hashtable();
            DbParameter parameter = null;

            _errorMessage = String.Empty;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return null;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection();

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inProcedure;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                parameter = dbInstance.Factory.CreateParameter();

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dbInstance.DbConnection.Open();

                dbInstance.Command.ExecuteNonQuery();

                dbInstance.DbConnection.Close();

                //Recupera os valores de saída e insere numa HashTable
                for (int i = 0; i < _params.Length; i++)
                {
                    ht.Add(_params[i].ParameterName, _params[i].Value);
                }

                return ht;
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return ht;

        }

        /// <summary>
        /// Executa uma procedure com parâmetros, retornando um Int32, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inProcedure">Nome da StoredProcedure.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>Hashtable.</returns>
        public static Hashtable InstanceProcedureOutData(string _connectionKey, string _inProcedure, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();
            Hashtable ht = new Hashtable();
            DbParameter parameter = null;

            _errorMessage = String.Empty;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return null;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inProcedure;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                parameter = dbInstance.Factory.CreateParameter();

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dbInstance.DbConnection.Open();

                dbInstance.Command.ExecuteNonQuery();

                dbInstance.DbConnection.Close();

                //Recupera os valores de saída e insere numa HashTable
                for (int i = 0; i < _params.Length; i++)
                {
                    ht.Add(_params[i].ParameterName, _params[i].Value);
                }

                return ht;
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return ht;

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="_inProcedure"></param>
        /// <param name="_errorMessage"></param>
        /// <param name="_transManager"></param>
        /// <param name="_transState"></param>
        /// <param name="_params"></param>
        /// <returns></returns>
        public static Hashtable InstanceProcedureOutData(string _inProcedure, out string _errorMessage, ref DBTransactionManager _transManager, DBTransactionManager.TransactionState _transState, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();
            Hashtable ht = new Hashtable();
            int recordsAffected = 0;

            _errorMessage = String.Empty;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return null;
            }

            try
            {
                //Executa uma transação isolada
                if (_transState == DBTransactionManager.TransactionState.StandAlone)
                {
                    //Abre a conexão
                    if (_transManager.DbConnection.State == ConnectionState.Closed)
                        _transManager.DbConnection.Open();

                    //Inicia a transação
                    _transManager.Transaction = _transManager.DbConnection.BeginTransaction();

                    ProcedureTransactionExecute(_inProcedure, out recordsAffected, _transManager, _params);

                    //Comita a transação
                    _transManager.Command.Transaction.Commit();

                    //Fecha a conexão da transação
                    _transManager.EndTransaction();
                }
                else
                {
                    //Bloqueio de execução depois da detecção de um erro
                    if (!_transManager.TransHold)
                    {
                        //Inicia a transação.DBTransactionManager(true)
                        if (_transState == DBTransactionManager.TransactionState.Begin)
                        {
                            //Abre a conexão e inicia a transação
                            if (_transManager.DbConnection.State == ConnectionState.Closed)
                                _transManager.DbConnection.Open();

                            _transManager.Transaction = _transManager.DbConnection.BeginTransaction();

                                ProcedureTransactionExecute(_inProcedure, out recordsAffected, _transManager, _params);
                        }
                        //Permite a execução de + de três comandos em uma transação
                        else if (_transState == DBTransactionManager.TransactionState.ContinueExecution)
                        {
                                ProcedureTransactionExecute(_inProcedure, out recordsAffected, _transManager, _params);
                        }
                        //Finaliza a transação.DBTransactionManager(false)
                        else if (_transState == DBTransactionManager.TransactionState.End)
                        {
                                ProcedureTransactionExecute(_inProcedure, out recordsAffected, _transManager, _params);

                            //Comita a transação
                            _transManager.Command.Transaction.Commit();

                            //Fecha a conexão da transação
                            _transManager.EndTransaction();
                        }
                    }
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção de banco ocorreu durante a execução dos comandos. Nenhuma das operações foi executada. <BR>";
                _errorMessage = _errorMessage + exp.Message;

                if (_transManager.Command.Transaction != null)
                    _transManager.Command.Transaction.Rollback();

                //Bloqueia a execução de qualquer chamada posterior dentro da transação em execução
                _transManager.TransHold = true;

                //Fecha a conexão da transação
                _transManager.EndTransaction();
            }
            catch (Exception exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução dos comandos. Nenhuma das operações foi executada. <BR>";
                _errorMessage = _errorMessage + exp.Message;

                if (_transManager.Command.Transaction != null)
                    _transManager.Command.Transaction.Rollback();

                //Bloqueia a execução de qualquer chamada posterior dentro da transação em execução
                _transManager.TransHold = true;

                //Fecha a conexão da transação
                _transManager.EndTransaction();
            }

            //Recupera os valores de saída e insere numa HashTable
            for (int i = 0; i < _params.Length; i++)
            {
                ht.Add(_params[i].ParameterName, _params[i].Value);
            }

            return ht;
        }

        #endregion

        #region { InstanceExecuteScalar }

        //Query InstanceExecuteScalar

        /// <summary>
        /// Executa uma query sem parâmetros, retornando um System.Object, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <returns>System.Object</returns>
        public static object InstanceExecuteScalar(string _inQuery, out string _errorMessage)
        {
            DBInstance dbInstance = new DBInstance();
            object scalar = null;
            _errorMessage = String.Empty;

            if (_inQuery.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return scalar;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection();

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inQuery;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                dbInstance.DbConnection.Open();

                scalar = dbInstance.Command.ExecuteScalar();

                dbInstance.DbConnection.Close();
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return scalar;
        }

        /// <summary>
        /// Executa uma query sem parâmetros, retornando um System.Object, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <returns>System.Object</returns>
        public static object InstanceExecuteScalar(string _connectionKey, string _inQuery, out string _errorMessage)
        {
            DBInstance dbInstance = new DBInstance();
            object scalar = null;
            _errorMessage = String.Empty;

            if (_inQuery.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return scalar;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inQuery;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                dbInstance.DbConnection.Open();

                scalar = dbInstance.Command.ExecuteScalar();

                dbInstance.DbConnection.Close();
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return scalar;
        }

        /// <summary>
        /// Executa uma query com parâmetros, retornando um System.Object, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>System.Object</returns>
        public static object InstanceExecuteScalar(string _inQuery, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();
            object scalar = null;
            DbParameter parameter = null;

            _errorMessage = String.Empty;

            if (_inQuery.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return scalar = null;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection();

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inQuery;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                parameter = dbInstance.Factory.CreateParameter();

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dbInstance.DbConnection.Open();

                scalar = dbInstance.Command.ExecuteScalar();

                dbInstance.DbConnection.Close();
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return scalar;
        }

        /// <summary>
        /// Executa uma query com parâmetros, retornando um System.Object, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inQuery">Query de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>System.Object</returns>
        public static object InstanceExecuteScalar(string _connectionKey, string _inQuery, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();
            object scalar = null;
            DbParameter parameter = null;

            _errorMessage = String.Empty;

            if (_inQuery.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return scalar = null;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inQuery;
                dbInstance.Command.CommandType = CommandType.Text;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                parameter = dbInstance.Factory.CreateParameter();

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dbInstance.DbConnection.Open();

                scalar = dbInstance.Command.ExecuteScalar();

                dbInstance.DbConnection.Close();
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return scalar;
        }

        //Procedure InstanceProcedureExecuteScalar

        /// <summary>
        /// Executa uma procedure sem parâmetros, retornando um System.Object, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_inProcedure">Nome da StoredProcedure.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <returns>System.Object</returns>
        public static object InstanceProcedureExecuteScalar(string _inProcedure, out string _errorMessage)
        {
            DBInstance dbInstance = new DBInstance();
            object scalar = null;
            _errorMessage = String.Empty;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return scalar;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection();

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inProcedure;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                dbInstance.DbConnection.Open();

                scalar = dbInstance.Command.ExecuteScalar();

                dbInstance.DbConnection.Close();
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return scalar;
        }

        /// <summary>
        /// Executa uma procedure sem parâmetros, retornando um System.Object, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inProcedure">Nome da StoredProcedure.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <returns>System.Object</returns>
        public static object InstanceProcedureExecuteScalar(string _connectionKey, string _inProcedure, out string _errorMessage)
        {
            DBInstance dbInstance = new DBInstance();
            object scalar = null;
            _errorMessage = String.Empty;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return scalar;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inProcedure;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                dbInstance.DbConnection.Open();

                scalar = dbInstance.Command.ExecuteScalar();

                dbInstance.DbConnection.Close();
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return scalar;
        }

        /// <summary>
        /// Executa uma procedure com parâmetros, retornando um System.Object, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_inProcedure">Nome da StoredProcedure.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>System.Object</returns>
        public static object InstanceProcedureExecuteScalar(string _inProcedure, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();
            object scalar = null;
            DbParameter parameter = null;

            _errorMessage = String.Empty;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return scalar = null;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection();

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inProcedure;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                parameter = dbInstance.Factory.CreateParameter();

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dbInstance.DbConnection.Open();

                scalar = dbInstance.Command.ExecuteScalar();

                dbInstance.DbConnection.Close();
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return scalar;
        }

        /// <summary>
        /// Executa uma procedure com parâmetros, retornando um System.Object, utilizando a conexão do banco.
        /// </summary>
        /// <param name="_connectionKey">Chave de Conexão do AppSettings do web.config.</param>
        /// <param name="_inProcedure">Nome da StoredProcedure.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>System.Object</returns>
        public static object InstanceProcedureExecuteScalar(string _connectionKey, string _inProcedure, out string _errorMessage, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();
            object scalar = null;
            DbParameter parameter = null;

            _errorMessage = String.Empty;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return scalar = null;
            }

            try
            {
                //Cria uma conexao na base padrão do AppSettings (Web.config)
                dbInstance.DbConnectionProvider = new DBConnectionProvider();
                dbInstance.DbConnection = dbInstance.DbConnectionProvider.GetConnection(_connectionKey);

                dbInstance.Factory = DBConnectionProvider.GetDbFactory();
                dbInstance.Command = dbInstance.Factory.CreateCommand();

                dbInstance.Command.CommandText = _inProcedure;
                dbInstance.Command.CommandType = CommandType.StoredProcedure;
                dbInstance.Command.Connection = dbInstance.DbConnection;

                parameter = dbInstance.Factory.CreateParameter();

                //
                for (int i = 0; i < _params.Length; i++)
                {
                    parameter = _params[i];

                    dbInstance.Command.Parameters.Add(parameter);
                }

                dbInstance.DbConnection.Open();

                scalar = dbInstance.Command.ExecuteScalar();

                dbInstance.DbConnection.Close();
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução do comando.  <BR>";
                _errorMessage = _errorMessage + exp.Message;
            }
            finally
            {
                if (dbInstance.Command != null)
                {
                    dbInstance.Command.Dispose();
                }
                if (dbInstance.DbConnection.State == ConnectionState.Open)
                {
                    dbInstance.DbConnection.Close();
                }
                if (dbInstance.DbConnection != null)
                {
                    dbInstance.DbConnection.Dispose();
                }
            }

            return scalar;
        }

        //Procedure Instance ExecuteEscalar Transaction - Execução Idependente de Comandos de uma Transação (blocos isolados)

        /// <summary>
        /// Executa uma procedure sem parâmetros, retornando um System.Object, utilizando a conexão do banco e transações.
        /// </summary>
        /// <param name="_inProcedure">Procedure de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_transManager">Objeto DBTransactionManager da transação.</param>
        /// <param name="_transState">Estado de execução do bloco da transação. (Begin, End, ContinueExecution, StandAlone)</param>
        /// <returns>System.Object</returns>
        public static object InstanceProcedureExecuteScalar(string _inProcedure, out string _errorMessage, ref DBTransactionManager _transManager, DBTransactionManager.TransactionState _transState)
        {
            DBInstance dbInstance = new DBInstance();

            _errorMessage = String.Empty;
            object result = null;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return null;
            }

            try
            {
                //Inicia uma transação isolada
                if (_transState == DBTransactionManager.TransactionState.StandAlone)
                {
                    //Abre a conexão e inicia a transação
                    if (_transManager.DbConnection.State == ConnectionState.Closed)
                        _transManager.DbConnection.Open();

                    _transManager.Transaction = _transManager.DbConnection.BeginTransaction();

                    ProcedureTransactionExecuteEscalar(_inProcedure, out result, _transManager);

                    //Comita a transação
                    _transManager.Command.Transaction.Commit();

                    //Fecha a conexão da transação
                    _transManager.EndTransaction();
                }

                //Bloqueio de execução depois da detecção de um erro
                if (!_transManager.TransHold)
                {
                    //Inicia a transação.DBTransactionManager(true)
                    if (_transState == DBTransactionManager.TransactionState.Begin)
                    {
                        //Abre a conexão e inicia a transação
                        if (_transManager.DbConnection.State == ConnectionState.Closed)
                            _transManager.DbConnection.Open();

                        _transManager.Transaction = _transManager.DbConnection.BeginTransaction();

                        ProcedureTransactionExecuteEscalar(_inProcedure, out result, _transManager);
                    }
                    //Permite a execução de + de três comandos em uma transação
                    else if (_transState == DBTransactionManager.TransactionState.ContinueExecution)
                    {
                        ProcedureTransactionExecuteEscalar(_inProcedure, out result, _transManager);
                    }
                    //Finaliza a transação.DBTransactionManager(false)
                    else
                    {
                        ProcedureTransactionExecuteEscalar(_inProcedure, out result, _transManager);

                        //Comita a transação
                        _transManager.Command.Transaction.Commit();

                        //Fecha a conexão da transação
                        _transManager.EndTransaction();
                    }
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção de banco ocorreu durante a execução dos comandos. Nenhuma das operações foi executada. <BR>";
                _errorMessage = _errorMessage + exp.Message;

                _transManager.Command.Transaction.Rollback();

                //Bloqueia a execução de qualquer chamada posterior dentro da transação em execução
                _transManager.TransHold = true;

                //Fecha a conexão da transação
                _transManager.EndTransaction();
            }
            catch (Exception exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução dos comandos. Nenhuma das operações foi executada. <BR>";
                _errorMessage = _errorMessage + exp.Message;

                _transManager.Command.Transaction.Rollback();

                //Bloqueia a execução de qualquer chamada posterior dentro da transação em execução
                _transManager.TransHold = true;

                //Fecha a conexão da transação
                _transManager.EndTransaction();
            }

            return result;
        }

        /// <summary>
        /// Executa uma procedure com parâmetros, retornando um  System.Object, utilizando a conexão do banco e transações.
        /// </summary>
        /// <param name="_inProcedure">Procedure de Entrada.</param>
        /// <param name="_errorMessage">Mensagem de Erro.</param>
        /// <param name="_transManager">Objeto DBTransactionManager da transação.</param>
        /// <param name="_transState">Estado de execução do bloco da transação. (Begin, End, ContinueExecution, StandAlone)</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        /// <returns>System.Object</returns>
        public static object InstanceProcedureExecuteScalar(string _inProcedure, out string _errorMessage, ref DBTransactionManager _transManager, DBTransactionManager.TransactionState _transState, params DbParameter[] _params)
        {
            DBInstance dbInstance = new DBInstance();

            _errorMessage = String.Empty;
            object result = null;

            if (_inProcedure.Length == 0)
            {
                _errorMessage = "Nenhum comando SQL foi informado.";
                return null;
            }

            try
            {
                //Executa uma transação isolada
                if (_transState == DBTransactionManager.TransactionState.StandAlone)
                {
                    //Abre a conexão
                    if (_transManager.DbConnection.State == ConnectionState.Closed)
                        _transManager.DbConnection.Open();

                    //Inicia a transação
                    _transManager.Transaction = _transManager.DbConnection.BeginTransaction();

                    ProcedureTransactionExecuteEscalar(_inProcedure, out result, _transManager, _params);

                    //Comita a transação
                    _transManager.Command.Transaction.Commit();

                    //Fecha a conexão da transação
                    _transManager.EndTransaction();
                }
                else
                {
                    //Bloqueio de execução depois da detecção de um erro
                    if (!_transManager.TransHold)
                    {
                        //Inicia a transação.DBTransactionManager(true)
                        if (_transState == DBTransactionManager.TransactionState.Begin)
                        {
                            //Abre a conexão e inicia a transação
                            if (_transManager.DbConnection.State == ConnectionState.Closed)
                                _transManager.DbConnection.Open();

                            _transManager.Transaction = _transManager.DbConnection.BeginTransaction();

                            ProcedureTransactionExecuteEscalar(_inProcedure, out result, _transManager, _params);
                        }
                        //Permite a execução de + de três comandos em uma transação
                        else if (_transState == DBTransactionManager.TransactionState.ContinueExecution)
                        {
                            ProcedureTransactionExecuteEscalar(_inProcedure, out result, _transManager, _params);
                        }
                        //Finaliza a transação.DBTransactionManager(false)
                        else if (_transState == DBTransactionManager.TransactionState.End)
                        {
                            ProcedureTransactionExecuteEscalar(_inProcedure, out result, _transManager, _params);

                            //Comita a transação
                            _transManager.Command.Transaction.Commit();

                            //Fecha a conexão da transação
                            _transManager.EndTransaction();
                        }
                    }
                }
            }
            catch (DbException exp)
            {
                _errorMessage = "Uma exceção de banco ocorreu durante a execução dos comandos. Nenhuma das operações foi executada. <BR>";
                _errorMessage = _errorMessage + exp.Message;

                if (_transManager.Command.Transaction != null)
                    _transManager.Command.Transaction.Rollback();

                //Bloqueia a execução de qualquer chamada posterior dentro da transação em execução
                _transManager.TransHold = true;

                //Fecha a conexão da transação
                _transManager.EndTransaction();
            }
            catch (Exception exp)
            {
                _errorMessage = "Uma exceção ocorreu durante a execução dos comandos. Nenhuma das operações foi executada. <BR>";
                _errorMessage = _errorMessage + exp.Message;

                if (_transManager.Command.Transaction != null)
                    _transManager.Command.Transaction.Rollback();

                //Bloqueia a execução de qualquer chamada posterior dentro da transação em execução
                _transManager.TransHold = true;

                //Fecha a conexão da transação
                _transManager.EndTransaction();
            }

            return null;
        }

        //Métodos Privados para DBProvider.Instance

        /// <summary>
        ///  Método privado para execução de uma procedure usando uma transação.
        /// </summary>
        /// <param name="_inProcedure">Procedure de Entrada.</param>
        /// <param name="_result">System.Object com o resultado da consulta.</param>
        /// <param name="_transManager">>Objeto DBTransactionManager da transação.</param>
        private static void ProcedureTransactionExecuteEscalar(string _inProcedure, out object _result, DBTransactionManager _transManager)
        {
            //Associa a transação ao comando
            _transManager.Command.Transaction = _transManager.Transaction;

            _transManager.Command.CommandText = _inProcedure;
            _transManager.Command.CommandType = CommandType.StoredProcedure;
            _transManager.Command.Connection = _transManager.DbConnection;

            //Executa o comando
            _result = _transManager.Command.ExecuteScalar();
        }

        /// <summary>
        /// Método privado para execução de uma procedure usando uma transação.
        /// </summary>
        /// <param name="_inProcedure">Procedure de Entrada.</param>
        /// <param name="_result">System.Object com o resultado da consulta.</param>
        /// <param name="_transManager">>Objeto DBTransactionManager da transação.</param>
        /// <param name="_params">Vetor de DbParameter com os parâmetros.</param>
        private static void ProcedureTransactionExecuteEscalar(string _inProcedure, out object _result, DBTransactionManager _transManager, params DbParameter[] _params)
        {
            DbParameter parameter = null;

            //Associa a transação ao comando
            _transManager.Command.Transaction = _transManager.Transaction;

            _transManager.Command.CommandText = _inProcedure;
            _transManager.Command.CommandType = CommandType.StoredProcedure;
            _transManager.Command.Connection = _transManager.DbConnection;

            _transManager.Command.Parameters.Clear();

            //Adiciona os parâmetros de execução ao comando da transação
            parameter = _transManager.Factory.CreateParameter();
            for (int i = 0; i < _params.Length; i++)
            {
                parameter = _params[i];

                _transManager.Command.Parameters.Add(parameter);
            }

            //Executa o comando
            _result = _transManager.Command.ExecuteScalar();
        }

        #endregion
    }
}
