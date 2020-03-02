using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace luval.word.navigator.terminal
{
    /// <summary>
    /// Provides an abstraction to the operations in a relational database
    /// </summary>
    public class Database
    {

        private readonly Func<IDbConnection> _connectionFactory;

        /// <summary>
        /// Creates a new instance of the database class
        /// </summary>
        /// <param name="connectionFactory">The function that will create a new instance of the <see cref="IDbConnection"/> object</param>
        public Database(Func<IDbConnection> connectionFactory)
        {
            _connectionFactory = connectionFactory;
        }

        /// <summary>
        /// Encapsulates the use of a <see cref="IDbConnection"/> object
        /// </summary>
        /// <param name="doSomething">Action that would use the <see cref="IDbConnection"/> object</param>
        public void WithConnection(Action<IDbConnection> doSomething)
        {
            using (var conn = _connectionFactory())
            {
                if (conn == null) throw new ArgumentException("Connection is not properly provided");
                doSomething(conn);
            }
        }

        /// <summary>
        /// Encapsulates the use of a <see cref="IDbTransaction"/> object
        /// </summary>
        /// <param name="doSomething">Action that would use the <see cref="IDbTransaction"/> object</param>
        public void WithTransaction(Action<IDbTransaction> doSomething)
        {
            WithConnection((conn) =>
            {
                try
                {
                    OpenConnection(conn);
                    using (var tran = conn.BeginTransaction())
                    {
                        try
                        {
                            doSomething(tran);
                            WorkTransaction(tran, () => { tran.Commit(); });
                        }
                        catch (Exception ex)
                        {
                            WorkTransaction(tran, () => { tran.Rollback(); });
                            throw new InvalidOperationException("Failed to complete the transaction", ex);
                        }
                    }
                    CloseConnection(conn);
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException("Failed to begin or rollback the transaction", ex);
                }
            });
        }

        /// <summary>
        /// Encapsulates the use of a <see cref="IDbCommand"/> object
        /// </summary>
        /// <param name="doSomething">Action that would use the <see cref="IDbCommand"/> object</param>
        public void WithCommand(Action<IDbCommand> doSomething)
        {

            WithTransaction((tran) =>
            {
                using (var cmd = tran.Connection.CreateCommand())
                {
                    cmd.Transaction = tran;
                    try
                    {
                        doSomething(cmd);
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException(string.Format("Unable to execute command. {0}", cmd.CommandText), ex);
                    }
                }
            });
        }

        /// <summary>
        /// Encapsulates the reading of a <see cref="IDataReader"/> and its <see cref="IDataRecord"/>
        /// </summary>
        /// <param name="cmd">The <see cref="IDbCommand"/> to use for the <see cref="IDataReader"/></param>
        /// <param name="behavior">The <see cref="CommandBehavior"/> to use</param>
        /// <param name="doSomething">Action that would be used for the <see cref="IDataRecord"/></param>
        public void WhileReading(IDbCommand cmd, CommandBehavior behavior, Action<IDataRecord> doSomething)
        {
            using (var reader = cmd.ExecuteReader(behavior))
            {
                while (reader.Read())
                {
                    doSomething(reader);
                }
            }
        }

        /// <summary>
        /// Encapsulates the reading of a <see cref="IDataReader"/> and its <see cref="IDataRecord"/>
        /// </summary>
        /// <param name="query">The query to execute in the command</param>
        /// <param name="type">The <see cref="CommandType"/> to use</param>
        /// <param name="behavior">The <see cref="CommandBehavior"/> to use</param>
        /// <param name="parameters">The <see cref="IDataParameterCollection"/> to use by the command</param>
        /// <param name="doSomething">Action that would be used for the <see cref="IDataRecord"/></param>
        public void WhileReading(string query, CommandType type, CommandBehavior behavior, IDataParameterCollection parameters, Action<IDataRecord> doSomething)
        {
            WithCommand((cmd) =>
            {
                PrepareCommand(cmd, query, type, 0, parameters);
                WhileReading(cmd, behavior, doSomething);
            });
        }

        /// <summary>
        /// Encapsulates the reading of a <see cref="IDataReader"/> and its <see cref="IDataRecord"/>
        /// </summary>
        /// <param name="query">The query to execute in the command</param>
        /// <param name="type">The <see cref="CommandType"/> to use</param>
        /// <param name="behavior">The <see cref="CommandBehavior"/> to use</param>
        /// <param name="doSomething">Action that would be used for the <see cref="IDataRecord"/></param>
        public void WhileReading(string query, CommandType type, CommandBehavior behavior, Action<IDataRecord> doSomething)
        {
            WhileReading(query, type, behavior, null, doSomething);
        }

        /// <summary>
        /// Encapsulates the reading of a <see cref="IDataReader"/> and its <see cref="IDataRecord"/>
        /// </summary>
        /// <param name="query">The query to execute in the command</param>
        /// <param name="type">The <see cref="CommandType"/> to use</param>
        /// <param name="doSomething">Action that would be used for the <see cref="IDataRecord"/></param>
        public void WhileReading(string query, CommandType type, Action<IDataRecord> doSomething)
        {
            WhileReading(query, type, CommandBehavior.Default, doSomething);
        }

        /// <summary>
        /// Encapsulates the reading of a <see cref="IDataReader"/> and its <see cref="IDataRecord"/>
        /// </summary>
        /// <param name="query">The query to execute in the command</param>
        /// <param name="doSomething">Action that would be used for the <see cref="IDataRecord"/></param>
        public void WhileReading(string query, Action<IDataRecord> doSomething)
        {
            WhileReading(query, CommandType.Text, doSomething);
        }

        /// <summary>
        /// Gets the scalar value of a command
        /// </summary>
        /// <param name="query">The query to execute in the command</param>
        /// <param name="type">The <see cref="CommandType"/> to use</param>
        /// <param name="parameters">The <see cref="IDataParameterCollection"/> to use by the command</param>
        /// <returns>The resulting scalar value of the command</returns>
        public object ExecuteScalar(string query, CommandType type, IDataParameterCollection parameters)
        {
            var result = default(object);
            WithCommand((cmd) =>
            {
                PrepareCommand(cmd, query, type, 0, parameters);
                result = cmd.ExecuteScalar();
            });
            return result;
        }

        /// <summary>
        /// Gets the scalar value of a command
        /// </summary>
        /// <param name="query">The query to execute in the command</param>
        /// <param name="type">The <see cref="CommandType"/> to use</param>
        /// <returns>The resulting scalar value of the command</returns>
        public object ExecuteScalar(string query, CommandType type)
        {
            return ExecuteScalar(query, type, null);
        }

        /// <summary>
        /// Gets the scalar value of a command
        /// </summary>
        /// <param name="query">The query to execute in the command</param>
        /// <returns>The resulting scalar value of the command</returns>
        public object ExecuteScalar(string query)
        {
            return ExecuteScalar(query, CommandType.Text);
        }

        /// <summary>
        /// Gets the number of records affected by the non query command
        /// </summary>
        /// <param name="query">The query to execute in the command</param>
        /// <param name="type">The <see cref="CommandType"/> to use</param>
        /// <param name="parameters">The <see cref="IDataParameterCollection"/> to use by the command</param>
        /// <returns>number of records affected by the non query command</returns>
        public int ExecuteNonQuery(string query, CommandType type, IDataParameterCollection parameters)
        {
            var result = 0;
            WithCommand((cmd) =>
            {
                PrepareCommand(cmd, query, type, 0, parameters);
                result = cmd.ExecuteNonQuery();
            });
            return result;
        }

        /// <summary>
        /// Gets the number of records affected by the non query command
        /// </summary>
        /// <param name="query">The query to execute in the command</param>
        /// <param name="type">The <see cref="CommandType"/> to use</param>
        /// <returns>number of records affected by the non query command</returns>
        public int ExecuteNonQuery(string query, CommandType type)
        {
            return ExecuteNonQuery(query, type, null);
        }

        /// <summary>
        /// Gets the number of records affected by the non query command
        /// </summary>
        /// <param name="query">The query to execute in the command</param>
        /// <returns>number of records affected by the non query command</returns>
        public int ExecuteNonQuery(string query)
        {
            return ExecuteNonQuery(query, CommandType.Text);
        }

        /// <summary>
        /// Gets a collection of records from a command into a <see cref="Dictionary{TKey, TValue}"/> of string for the field and object for the value
        /// </summary>
        /// <param name="cmd">The <see cref="IDbCommand"/> to execute</param>
        /// <returns>A <see cref="List{T}"/> of records from a command into a <see cref="<Dictionary"/> of string for the field and object for the value</returns>
        public List<Dictionary<string, object>> ExecuteToDictionaryList(IDbCommand cmd)
        {
            var result = new List<Dictionary<string, object>>();
            WhileReading(cmd, CommandBehavior.CloseConnection, (row) =>
            {
                LoadRecordIntoDictionaryList(result, row);
            });
            return result;
        }

        /// <summary>
        /// Gets a collection of records from a command into a <see cref="Dictionary{TKey, TValue}"/> of string for the field and object for the value
        /// </summary>
        /// <param name="query">The query to execute in the command</param>
        /// <param name="type">The <see cref="CommandType"/> to use</param>
        /// <param name="parameters">The <see cref="IDataParameterCollection"/> to use by the command</param>
        /// <returns>A <see cref="List{T}"/> of records from a command into a <see cref="<Dictionary"/> of string for the field and object for the value</returns>
        public List<Dictionary<string, object>> ExecuteToDictionaryList(string query, CommandType type, IDataParameterCollection parameters)
        {
            var result = new List<Dictionary<string, object>>();
            WhileReading(query, type, CommandBehavior.Default, parameters, (row) =>
            {
                LoadRecordIntoDictionaryList(result, row);
            });
            return result;
        }

        /// <summary>
        /// Gets a collection of records from a command into a <see cref="Dictionary{TKey, TValue}"/> of string for the field and object for the value
        /// </summary>
        /// <param name="query">The query to execute in the command</param>
        /// <param name="type">The <see cref="CommandType"/> to use</param>
        /// <returns>A <see cref="List{T}"/> of records from a command into a <see cref="<Dictionary"/> of string for the field and object for the value</returns>
        public List<Dictionary<string, object>> ExecuteToDictionaryList(string query, CommandType type)
        {
            return ExecuteToDictionaryList(query, type, null);
        }

        /// <summary>
        /// Gets a collection of records from a command into a <see cref="Dictionary{TKey, TValue}"/> of string for the field and object for the value
        /// </summary>
        /// <param name="query">The query to execute in the command</param>
        /// <returns>A <see cref="List{T}"/> of records from a command into a <see cref="<Dictionary"/> of string for the field and object for the value</returns>
        public List<Dictionary<string, object>> ExecuteToDictionaryList(string query)
        {
            return ExecuteToDictionaryList(query, CommandType.Text);
        }

        /// <summary>
        /// Gets a collection of entities into a <see cref="List{T}"/>
        /// </summary>
        /// <typeparam name="T">The entity type to populate</typeparam>
        /// <param name="cmd">The <see cref="IDbCommand"/> to execute</param>
        /// <returns>A collection of entities into a <see cref="List{T}"/></returns>
        public List<T> ExecuteToEntityList<T>(IDbCommand cmd)
        {
            var result = new List<T>();
            WhileReading(cmd, CommandBehavior.CloseConnection, (r) =>
            {
                result.Add(GetEntityFromDataRecord<T>(r));
            });
            return result;
        }

        /// <summary>
        /// Gets a collection of entities into a <see cref="List{T}"/>
        /// </summary>
        /// <typeparam name="T">The entity type to populate</typeparam>
        /// <param name="query">The query to execute in the command</param>
        /// <param name="type">The <see cref="CommandType"/> to use</param>
        /// <param name="parameters">The <see cref="IDataParameterCollection"/> to use by the command</param>
        /// <returns>A collection of entities into a <see cref="List{T}"/></returns>
        public List<T> ExecuteToEntityList<T>(string query, CommandType type, IDataParameterCollection parameters)
        {
            var result = new List<T>();
            WhileReading(query, type, CommandBehavior.CloseConnection, parameters, (r) =>
            {
                result.Add(GetEntityFromDataRecord<T>(r));
            });
            return result;
        }

        /// <summary>
        /// Gets a collection of entities into a <see cref="List{T}"/>
        /// </summary>
        /// <typeparam name="T">The entity type to populate</typeparam>
        /// <param name="query">The query to execute in the command</param>
        /// <param name="type">The <see cref="CommandType"/> to use</param>
        /// <returns>A collection of entities into a <see cref="List{T}"/></returns>
        public List<T> ExecuteToEntityList<T>(string query, CommandType type)
        {
            return ExecuteToEntityList<T>(query, type, null);
        }

        /// <summary>
        /// Gets a collection of entities into a <see cref="List{T}"/>
        /// </summary>
        /// <typeparam name="T">The entity type to populate</typeparam>
        /// <param name="query">The query to execute in the command</param>
        /// <returns>A collection of entities into a <see cref="List{T}"/></returns>
        public List<T> ExecuteToEntityList<T>(string query)
        {
            return ExecuteToEntityList<T>(query, CommandType.Text);
        }

        #region Private Helper Methods

        private T GetEntityFromDataRecord<T>(IDataRecord record)
        {
            var type = typeof(T);
            var entity = Activator.CreateInstance(type);
            for (int i = 0; i < record.FieldCount; i++)
            {
                var p = type.GetProperty(record.GetName(i));
                p.SetValue(entity, Convert.ChangeType(record.GetValue(i), p.PropertyType));
            }
            return ((T)entity);
        }

        private void LoadRecordIntoDictionaryList(List<Dictionary<string, object>> recordSet, IDataRecord row)
        {
            var record = new Dictionary<string, object>();
            for (int i = 0; i < row.FieldCount; i++)
            {
                record.Add(row.GetName(i), row.GetValue(i));
            }
            recordSet.Add(record);
        }

        private void OpenConnection(IDbConnection conn)
        {
            try
            {
                if (conn.State == ConnectionState.Closed) conn.Open();
            }
            catch (Exception ex)
            {

                throw new InvalidOperationException("Unable to open the connection", ex);
            }
        }

        private void CloseConnection(IDbConnection conn)
        {
            try
            {
                if (conn != null && conn.State != ConnectionState.Closed) conn.Close();
            }
            catch (Exception ex)
            {

                throw new InvalidOperationException("Unable to close the connection", ex);
            }
        }

        private void WorkTransaction(IDbTransaction tran, Action action)
        {
            if (tran == null || tran.Connection == null || tran.Connection.State == ConnectionState.Closed) return;
            action();
        }

        private void PrepareCommand(IDbCommand cmd, string query, CommandType type, int timeout, IDataParameterCollection parameters)
        {
            cmd.CommandText = query;
            cmd.CommandType = type;
            if (parameters != null && parameters.Count > 0)
                foreach (var param in parameters)
                {
                    cmd.Parameters.Add(param);
                }
            if (timeout > 0) cmd.CommandTimeout = timeout;
            else cmd.CommandTimeout = cmd.Connection.ConnectionTimeout;
        }

        #endregion
    }
}
