using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;

namespace BHD_Framework
{
    //class SqlServer
    //{
    //}
    public static class SqlSettings
    {
        public static readonly int MinTimeoutSpace = 0; // second
        public static readonly int MaxTimeoutSpace = 2147483647; // second
        public static readonly int DefaultConnectionTimeoutSpace = 3; // second
        public static readonly int DefaultExecutionTimeoutSpace = 10; // second
        //
        public static readonly string[] ServerText = new string[] { "server", "data source" };
        public static readonly string[] DatabaseText = new string[] { "database", "initial catalog" };
        public static readonly string[] UidText = new string[] { "user id" };
        public static readonly string[] PwdText = new string[] { "pwd", "password" };
        public static readonly string[] TimeOutText = new string[] { "connection timeout" };
        //
        public static readonly string MsgInvalidConnectionString = "Connection string is not valid.";
        public static readonly string MsgInvalidTimeOut = "Time-out space value is not valid.";
        public static readonly string MsgCantOpenConnection = "Can not open connection to server.";
    }
    public class SqlObject
    {
        #region Constructor
        public SqlObject(string StringConnection) { _connectionString = standardizedConnectionString(StringConnection); }
        #endregion 

        private string _connectionString;
        private string standardizedConnectionString(string ConfigString)
        {
            string _result = "";
            //server = ofileserver.msv.dc\SQL2014; database = master ; user id = fn ; pwd =   f@1234
            //workstation id=mankichiws.mssql.somee.com; data source=mankichiws.mssql.somee.com; user id=haidang_mankichi_SQLLogin_1; pwd=bnwhbvlzp8; packet size=4096; persist security info=False;
            try
            {
                string[] arrItems = ConfigString.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                string svr = "", dtb = "", uid = "", pwd = "", timeOut = "";
                List<string> _lstAddition = new List<string>();
                foreach (string item in arrItems)
                {
                    bool assigned = false;
                    #region Server
                    if (!assigned) foreach (string s in SqlSettings.ServerText)
                        {
                            string _value = Utility.GetValueOf(s, item);
                            if (_value == "") continue;
                            svr = _value;
                            assigned = true;
                            break;
                        }
                    #endregion
                    #region Database
                    if (!assigned) foreach (string s in SqlSettings.DatabaseText)
                        {
                            string _value = Utility.GetValueOf(s, item);
                            if (_value == "") continue;
                            dtb = _value;
                            assigned = true;
                            break;
                        }
                    #endregion
                    #region User
                    if (!assigned) foreach (string s in SqlSettings.UidText)
                        {
                            string _value = Utility.GetValueOf(s, item);
                            if (_value == "") continue;
                            uid = _value;
                            assigned = true;
                            break;
                        }
                    #endregion
                    #region Password
                    if (!assigned) foreach (string s in SqlSettings.PwdText)
                        {
                            string _value = Utility.GetValueOf(s, item);
                            if (_value == "") continue;
                            pwd = _value;
                            assigned = true;
                            break;
                        }
                    #endregion
                    #region Timeout
                    if (!assigned) foreach (string s in SqlSettings.TimeOutText)
                        {
                            string _value = Utility.GetValueOf(s, item);
                            if (_value == "") continue;
                            timeOut = _value;
                            assigned = true;
                            break;
                        }
                    #endregion
                    #region Remain info
                    if (!assigned)
                    {
                        string[] strs = item.Split(new char[] { '=' }, StringSplitOptions.RemoveEmptyEntries);
                        string r = strs.Length == 2 ? string.Concat(strs[0], "=", strs[1]) : item;
                        _lstAddition.Add(r);
                    }
                    #endregion
                }
                _result = string.Concat(SqlSettings.ServerText[0], "=", svr, ";", SqlSettings.DatabaseText[0], "=", dtb, ";", SqlSettings.UidText[0], "=", uid, ";", SqlSettings.PwdText[0], "=", pwd);
                if (timeOut == "")
                {
                    timeOut = string.Concat(SqlSettings.TimeOutText[0], "=", SqlSettings.DefaultConnectionTimeoutSpace);
                    _result += string.Concat(";", timeOut);
                }
                else
                {
                    int _timeOutSpace; int.TryParse(timeOut, out _timeOutSpace);
                    _timeOutSpace = getValidTimeoutSpace(_timeOutSpace, TimeoutType.Connection);
                    timeOut = string.Concat(SqlSettings.TimeOutText[0], "=", _timeOutSpace);
                    _result += string.Concat(";", timeOut);
                    //
                    _connectionTimeout = _timeOutSpace;
                }
                //Console.WriteLine(_result);
                if (_result.Length == 0) return _result;
                foreach (string s in _lstAddition) _result += string.Concat(";", s);
            }
            catch { }
            return _result;
        }
        private void updateTimeoutConnectionString()
        {
            if (_connectionString.Length == 0)
            {
                return;
                throw new Exception(SqlSettings.MsgInvalidConnectionString);
            }
            //_connectionTimeout
            int _beginIndex1 = -1, _lengthText = 0;
            string[] arr = _connectionString.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            bool _break = false;
            foreach (string item in arr)
            {
                foreach (string s in SqlSettings.TimeOutText)
                {
                    if (!item.ToLower().StartsWith(s)) continue;
                    _beginIndex1 = _connectionString.IndexOf(s); // connection timeout=30
                    _lengthText = item.Length;
                    _break = true;
                    break;
                }
                if (_break) break;
            }
            string str = string.Concat(SqlSettings.TimeOutText[0], "=", _connectionTimeout);
            if (_beginIndex1 < 0)
            {
                _connectionString += (_connectionString.EndsWith(";") ? "" : ";") + str;
            }
            else
            {
                _connectionString = string.Concat(_connectionString.Substring(0, _beginIndex1), str, _connectionString.Substring(_beginIndex1 + _lengthText, _connectionString.Length - _beginIndex1 - _lengthText));
            }
        }
        public string ConnectionString
        {
            get { return _connectionString; }
            set
            {
                _connectionString = standardizedConnectionString(value);
            }
        }

        private SqlConnection _connection;
        private enum TimeoutType
        {
            Connection,
            Execution
        }
        private int _connectionTimeout = SqlSettings.DefaultConnectionTimeoutSpace;
        private int _executionTimeout = SqlSettings.DefaultExecutionTimeoutSpace;

        private int getValidTimeoutSpace(int? TimeoutSpace, TimeoutType TimeoutOption)
        {
            int _defaultTimeout = (TimeoutOption == TimeoutType.Connection ? SqlSettings.DefaultConnectionTimeoutSpace : SqlSettings.DefaultExecutionTimeoutSpace);
            if (TimeoutSpace != null)
            {
                if (TimeoutSpace < SqlSettings.MinTimeoutSpace || SqlSettings.MaxTimeoutSpace < TimeoutSpace)
                {
                    throw new Exception(SqlSettings.MsgInvalidTimeOut);
                }
            }
            else
            {
                TimeoutSpace = _defaultTimeout;
            }
            //
            int _timeOutSpace = (TimeoutSpace.Value == _defaultTimeout ? _defaultTimeout : TimeoutSpace.Value);
            return _timeOutSpace;
        }
        //
        public int ConnectionTimeout
        {
            get { return _connectionTimeout; }
            set
            {
                if (SqlSettings.MinTimeoutSpace <= value && value <= SqlSettings.MaxTimeoutSpace)
                {
                    _connectionTimeout = value;
                    updateTimeoutConnectionString();
                }
                else
                    throw new Exception(SqlSettings.MsgInvalidTimeOut);
            }
        }
        public int ExecutionTimeout
        {
            get { return _executionTimeout; }
            set
            {
                if (SqlSettings.MinTimeoutSpace <= value && value <= SqlSettings.MaxTimeoutSpace)
                    _executionTimeout = value;
                else
                    throw new Exception(SqlSettings.MsgInvalidTimeOut);
            }
        }

        private void OpenConnection()
        {
            try
            {
                _connection = new SqlConnection(_connectionString);
                _connection.Open();
                //Console.WriteLine("Connection open.");
            }
            catch
            {
                //Console.WriteLine("Can not open connection!");
                throw new Exception(SqlSettings.MsgCantOpenConnection);
            }
        }
        private void CloseConnection()
        {
            try
            {
                if (_connection != null)
                {
                    _connection.Close();
                    _connection.Dispose();
                }
            }
            catch { }
        }

        public bool GetConnectionState()
        {
            bool _result = false;
            try
            {
                OpenConnection();
                _result = true;
            }
            catch { }
            finally { CloseConnection(); }
            return _result;
        }

        private SqlCommand getSqlCommand(string SqlString, int? TimeoutSpace)
        {
            int _timeout = getValidTimeoutSpace(TimeoutSpace, TimeoutType.Execution);
            SqlCommand SqlCmd = new SqlCommand(SqlString, _connection);
            SqlCmd.CommandType = CommandType.Text;
            SqlCmd.CommandTimeout = _timeout;
            return SqlCmd;
        }
        public void ExecSql(string SqlString) { ExecSql(SqlString, TimeoutSpace: null); }
        public void ExecSql(string SqlString, int? TimeoutSpace)
        {
            OpenConnection();
            SqlCommand sqlCmd = getSqlCommand(SqlString, TimeoutSpace);
            sqlCmd.ExecuteNonQuery();
            sqlCmd.Dispose();
            CloseConnection();
        }
        public void ExecSql(string SqlString, SqlParameter[] SqlParams) { ExecSql(SqlString, SqlParams, TimeoutSpace: null); }
        public void ExecSql(string SqlString, SqlParameter[] SqlParams, int? TimeoutSpace)
        {
            OpenConnection();
            SqlCommand sqlCmd = getSqlCommand(SqlString, TimeoutSpace);
            if (SqlParams != null)
            {
                foreach (SqlParameter parameter in SqlParams)
                    sqlCmd.Parameters.Add(parameter);
            }
            sqlCmd.ExecuteNonQuery();
            sqlCmd.Parameters.Clear();
            sqlCmd.Dispose();
            CloseConnection();
        }

        public object ExecScalar(string SqlString) { return ExecScalar(SqlString, TimeoutSpace: null); }
        public object ExecScalar(string SqlString, int? TimeoutSpace)
        {
            OpenConnection();
            SqlCommand sqlCmd = getSqlCommand(SqlString, TimeoutSpace);
            object obj = sqlCmd.ExecuteScalar();
            sqlCmd.Dispose();
            CloseConnection();
            return obj;
        }
        public object ExecScalar(string SqlString, int? TimeoutSpace, object ExceptionResult)
        {
            try { return ExecScalar(SqlString, TimeoutSpace); }
            catch { return ExceptionResult; }
        }
        public object ExecScalar(string SqlString, SqlParameter[] SqlParams) { return ExecScalar(SqlString, SqlParams, TimeoutSpace: null); }
        public object ExecScalar(string SqlString, SqlParameter[] SqlParams, int? TimeoutSpace)
        {
            OpenConnection();
            SqlCommand sqlCmd = getSqlCommand(SqlString, TimeoutSpace);
            if (SqlParams != null)
            {
                foreach (SqlParameter parameter in SqlParams)
                    sqlCmd.Parameters.Add(parameter);
            }
            object obj = sqlCmd.ExecuteScalar();
            sqlCmd.Parameters.Clear();
            sqlCmd.Dispose();
            CloseConnection();
            return obj;
        }
        public object ExecScalar(string SqlString, SqlParameter[] SqlParams, object ExceptionResult) { return ExecScalar(SqlString, SqlParams, ExceptionResult, TimeoutSpace: null); }
        public object ExecScalar(string SqlString, SqlParameter[] SqlParams, object ExceptionResult, int? TimeoutSpace)
        {
            try { return ExecScalar(SqlString, SqlParams, TimeoutSpace); }
            catch { return ExceptionResult; }
        }

        private SqlDataAdapter getSqlAdapter(string SqlString, int? TimeoutSpace)
        {
            SqlCommand sqlCmd = getSqlCommand(SqlString, TimeoutSpace);
            SqlDataAdapter sqlAdapter = new SqlDataAdapter(sqlCmd);
            sqlCmd.Dispose();
            return sqlAdapter;
        }
        public DataTable ExecDataTable(string SqlString) { return ExecDataTable(SqlString, TimeoutSpace: null); }
        public DataTable ExecDataTable(string SqlString, int? TimeoutSpace)
        {
            int _timeout = getValidTimeoutSpace(TimeoutSpace, TimeoutType.Execution);
            DataTable dt = new DataTable();
            OpenConnection();
            SqlDataAdapter sqlAdapter = getSqlAdapter(SqlString, TimeoutSpace);
            sqlAdapter.Fill(dt);
            sqlAdapter.Dispose();
            CloseConnection();
            return dt;
        }

        public DataSet ExecDataSet(string SqlString) { return ExecDataSet(SqlString, TimeoutSpace: null); }
        public DataSet ExecDataSet(string SqlString, int? TimeoutSpace)
        {
            DataSet ds = new DataSet();
            OpenConnection();
            SqlDataAdapter sqlAdapter = getSqlAdapter(SqlString, TimeoutSpace);
            sqlAdapter.Fill(ds);
            sqlAdapter.Dispose();
            CloseConnection();
            return ds;
        }

        private SqlDataAdapter getSqlAdapter(string SqlString, SqlParameter[] SqlParams, int? TimeoutSpace)
        {
            SqlCommand sqlCmd = getSqlCommand(SqlString, TimeoutSpace);
            if (SqlParams != null)
            {
                foreach (SqlParameter parameter in SqlParams)
                    sqlCmd.Parameters.Add(parameter);
            }
            SqlDataAdapter sqlAdapter = new SqlDataAdapter(sqlCmd);
            sqlCmd.Dispose();
            return sqlAdapter;
        }
        public DataTable ExecDataTable(string SqlString, SqlParameter[] SqlParams) { return ExecDataTable(SqlString, SqlParams, TimeoutSpace: null); }
        public DataTable ExecDataTable(string SqlString, SqlParameter[] SqlParams, int? TimeoutSpace)
        {
            DataTable dt = new DataTable();
            OpenConnection();
            SqlDataAdapter sqlAdapter = getSqlAdapter(SqlString, SqlParams, TimeoutSpace);
            sqlAdapter.Fill(dt);
            sqlAdapter.Dispose();
            CloseConnection();
            return dt;
        }
        public DataSet ExecDataSet(string SqlString, SqlParameter[] SqlParams) { return ExecDataSet(SqlString, SqlParams, TimeoutSpace: null); }
        public DataSet ExecDataSet(string SqlString, SqlParameter[] SqlParams, int? TimeoutSpace)
        {
            DataSet ds = new DataSet();
            OpenConnection();
            SqlDataAdapter sqlAdapter = getSqlAdapter(SqlString, SqlParams, TimeoutSpace);
            sqlAdapter.Fill(ds);
            sqlAdapter.Dispose();
            CloseConnection();
            return ds;
        }

    }


}
