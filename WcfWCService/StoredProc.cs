using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using System.Runtime.InteropServices;
using System.Collections;
using System.Diagnostics;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.Configuration;

namespace WcfWCService
{
    public class Environment
    {
        public string Get_Environment_String_Value(string sName)
        {
            string sLine = "";
            string sVariable = "";
            string sValue = "";
            string sFilePath = "";

            sFilePath = "C:\\WebRoot\\WCIIS\\Environment\\Environment.txt";

            using (StreamReader sr = new StreamReader(sFilePath))
            {
                while (sr.Peek() > 0)
                {
                    sLine = sr.ReadLine();
                    //Get bit before the =
                    sVariable = sLine.Substring(0, sLine.IndexOf("="));
                    sVariable = sVariable.Trim();
                    sValue = sLine.Substring(sLine.IndexOf("=") + 1, sLine.Length - sLine.IndexOf("=") - 1);
                    sValue = sValue.Trim();
                    if (sVariable == sName)
                    {
                        sr.Close();
                        return sValue;
                    }

                }
                sr.Close();
                return "";
            }
        }

    }
    public class RecordSet
    {
        public int m_RecordCount = 0;
        public int m_ColumnCount = 0;
        public string m_ErrMessage = "";
        public string m_sCatalog = "";
        public int m_iWebAppId = 2;

        public void SetWebApp(int iWebAppId)
        {
            m_iWebAppId = iWebAppId;
        }
        public string SqlConnectionStr()
        {
            Environment env = new Environment();
            string sEnvironment = env.Get_Environment_String_Value("Environment");

            switch(m_iWebAppId)
            {
                case 1:
                    m_sCatalog = "pims";
                    break;
                default:
                    m_sCatalog = "Regain";
                    break;

            }
            if (sEnvironment == "DEVWCREGAIN")
            {
                return @"Data Source=VSRS12;Initial Catalog=" + m_sCatalog + ";User ID=pimsadmin;Password=N33dt0kn0w";
            }
            else if (sEnvironment == "PRODUCTIONREGAIN")
            {
                return @"Data Source=VSRS06;Initial Catalog=" + m_sCatalog + ";User ID=pimsadmin;Password=N33dt0kn0w";
            }
            else
            {
                return @"Data Source=VSRS12;Initial Catalog=" + m_sCatalog + ";User ID=pimsadmin;Password=N33dt0kn0w";
            }

        }

        public bool ExecuteSQL(string m_sSQL)
        {
            SqlConnection myConnection = new SqlConnection(SqlConnectionStr());

            try
            {

                myConnection.Open();

                SqlCommand SqlCmd = new SqlCommand(m_sSQL, myConnection);
                m_RecordCount = SqlCmd.ExecuteNonQuery();
                myConnection.Close();

                m_ErrMessage = "";
                return true;
            }
            catch
            {
                m_ErrMessage = "Class GenericRstTable method 'rstInsert' failed Argumments sSQL: " + m_sSQL;
                if (myConnection.State == ConnectionState.Open)
                {
                    myConnection.Close();
                }
                m_RecordCount = -1;
                return false;
            }
        }

        public string Get_NVarchar(DataSet ds, string sColumnName, int iRow)
        {
            int iColumnNo = 0;

            if (ds.Tables.Count > 0)
            {
                iColumnNo = ds.Tables[0].Columns.IndexOf(sColumnName);

                if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                {
                    if (Convert.IsDBNull(ds.Tables[0].Rows[iRow].ItemArray.GetValue(iColumnNo)))
                    {
                        return "";
                    }
                    else
                    {
                        return ds.Tables[0].Rows[iRow].ItemArray.GetValue(iColumnNo).ToString();
                    }
                }
                else
                {
                    return "";
                }
            }
            else
            {
                return "";
            }
        }

        public DataSet OpenRecordset(string sSQL, string cnn)
        {
            try
            {

                SqlConnection Conn = new SqlConnection(cnn);
                Conn.Open();
                SqlCommand SqlCmd = new SqlCommand(sSQL, Conn);
                SqlCmd.CommandTimeout = 0;
                SqlDataReader OpenRecordset = SqlCmd.ExecuteReader(CommandBehavior.CloseConnection);
                DataSet ds = new DataSet();
                ds = convertDataReaderToDataSet(OpenRecordset);
                OpenRecordset.Close();
                Conn.Close();
                m_RecordCount = ds.Tables[0].Rows.Count;
                m_ColumnCount = ds.Tables[0].Columns.Count;
                return ds;
                //                return SqlCmd.ExecuteReader();
            }
            catch
            {
                return (null);
            }

        }

        public static DataSet convertDataReaderToDataSet(SqlDataReader reader)
        {
            DataSet dataSet = new DataSet();
            do
            {
                // Create new data table

                DataTable schemaTable = reader.GetSchemaTable();
                DataTable dataTable = new DataTable();

                if (schemaTable != null)
                {
                    // A query returning records was executed

                    for (int i = 0; i < schemaTable.Rows.Count; i++)
                    {
                        DataRow dataRow = schemaTable.Rows[i];
                        // Create a column name that is unique in the data table
                        string columnName = (string)dataRow["ColumnName"]; //+ "<C" + i + "/>";
                        // Add the column definition to the data table
                        DataColumn column = new DataColumn(columnName, (Type)dataRow["DataType"]);
                        dataTable.Columns.Add(column);
                    }

                    dataSet.Tables.Add(dataTable);

                    // Fill the data table we just created

                    while (reader.Read())
                    {
                        DataRow dataRow = dataTable.NewRow();

                        for (int i = 0; i < reader.FieldCount; i++)
                            dataRow[i] = reader.GetValue(i);

                        dataTable.Rows.Add(dataRow);
                    }
                }
                else
                {
                    // No records were returned

                    DataColumn column = new DataColumn("RowsAffected");
                    dataTable.Columns.Add(column);
                    dataSet.Tables.Add(dataTable);
                    DataRow dataRow = dataTable.NewRow();
                    dataRow[0] = reader.RecordsAffected;
                    dataTable.Rows.Add(dataRow);
                }
            }
            while (reader.NextResult());
            return dataSet;
        }

    }

    public class StoredProc
    {
            string m_StoredProcName = "";
            ArrayList m_ParamArray = new ArrayList();
            ArrayList m_ParamName = new ArrayList();
            int m_ParamCount = 0;
            DataSet m_DataSet = new DataSet();
            DataTable m_DataTable = new DataTable();
            RecordSet GenRstTab = new RecordSet();
            string m_ErrorMsg = "";
            SqlTransaction m_trans = null;
            SqlConnection m_Connection = new SqlConnection();


            public string GetErrorMsg()
            {
                return m_ErrorMsg;
            }

            public string SetProcName(string sProcName)
            {
                m_StoredProcName = sProcName;
                return m_StoredProcName;
            }

            public string GetProcName()
            {
                return m_StoredProcName;
            }

            //This is a zero based return set so call the row number from zero NOT one
            public string Get_NVarchar(string sColumnName, int iRow)
            {
                int iColumnNo = 0;

                if (m_DataSet.Tables.Count > 0)
                {
                    iColumnNo = m_DataSet.Tables[0].Columns.IndexOf(sColumnName);

                    if (m_DataSet.Tables.Count > 0 && m_DataSet.Tables[0].Rows.Count > 0)
                    {
                        if (Convert.IsDBNull(m_DataSet.Tables[0].Rows[iRow].ItemArray.GetValue(iColumnNo)))
                        {
                            return "";
                        }
                        else
                        {
                            return m_DataSet.Tables[0].Rows[iRow].ItemArray.GetValue(iColumnNo).ToString();
                        }
                    }
                    else
                    {
                        return "";
                    }
                }
                else
                {
                    return "";
                }
            }

            public int Get_Int(string sColumnName, int iRow)
            {
                int iColumnNo = 0;

                iColumnNo = m_DataSet.Tables[0].Columns.IndexOf(sColumnName);

                if (m_DataSet.Tables.Count > 0 && m_DataSet.Tables[0].Rows.Count > 0)
                {
                    if (Convert.IsDBNull(m_DataSet.Tables[0].Rows[iRow].ItemArray.GetValue(iColumnNo)))
                    {
                        return 0;
                    }
                    else
                    {
                        return Convert.ToInt32(m_DataSet.Tables[0].Rows[iRow].ItemArray.GetValue(iColumnNo));
                    }
                }
                else
                {
                    return 0;
                }
            }

            public bool Get_Bit(string sColumnName, int iRow)
            {
                int iColumnNo = 0;

                if (m_DataSet.Tables.Count > 0)
                {
                    iColumnNo = m_DataSet.Tables[0].Columns.IndexOf(sColumnName);

                    if (m_DataSet.Tables.Count > 0 && m_DataSet.Tables[0].Rows.Count > 0)
                    {
                        if (Convert.IsDBNull(m_DataSet.Tables[0].Rows[iRow].ItemArray.GetValue(iColumnNo)))
                        {
                            return false;
                        }
                        else
                        {
                            return Convert.ToBoolean(m_DataSet.Tables[0].Rows[iRow].ItemArray.GetValue(iColumnNo));
                        }
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }


            public DateTime Get_DateTime(string sColumnName, int iRow)
            {
                int iColumnNo = 0;

                if (m_DataSet.Tables.Count > 0)
                {
                    iColumnNo = m_DataSet.Tables[0].Columns.IndexOf(sColumnName);

                    if (m_DataSet.Tables.Count > 0 && m_DataSet.Tables[0].Rows.Count > 0)
                    {
                        if (Convert.IsDBNull(m_DataSet.Tables[0].Rows[iRow].ItemArray.GetValue(iColumnNo)))
                        {
                            return Convert.ToDateTime("1/1/1900");
                        }
                        else
                        {
                            return Convert.ToDateTime(m_DataSet.Tables[0].Rows[iRow].ItemArray.GetValue(iColumnNo));
                        }
                    }
                    else
                    {
                        return Convert.ToDateTime("1/1/1900");
                    }
                }
                else
                {
                    return Convert.ToDateTime("1/1/1900");
                }
            }


            public double Get_Float(string sColumnName, int iRow)
            {
                int iColumnNo = 0;

                if (m_DataSet.Tables.Count > 0)
                {
                    iColumnNo = m_DataSet.Tables[0].Columns.IndexOf(sColumnName);

                    if (m_DataSet.Tables.Count > 0 && m_DataSet.Tables[0].Rows.Count > 0)
                    {
                        if (Convert.IsDBNull(m_DataSet.Tables[0].Rows[iRow].ItemArray.GetValue(iColumnNo)))
                        {
                            return 0.0;
                        }
                        else
                        {
                            return Convert.ToDouble(m_DataSet.Tables[0].Rows[iRow].ItemArray.GetValue(iColumnNo));
                        }
                    }
                    else
                    {
                        return 0.0;
                    }
                }
                else
                {
                    return 0.0;
                }
            }


            public object GetParamArray()
            {
                return m_ParamArray;
            }

            public string[] GetParamNameArray()
            {
                string[] ParamStringArray = new string[m_ParamCount];
                int i = 0;

                for (i = 0; i < m_ParamCount; i++)
                {
                    ParamStringArray[i] = m_ParamName[i].ToString();
                }
                return ParamStringArray;
            }
            public int GetParamCount()
            {
                return m_ParamCount;
            }
            public DataSet GetDataSet()
            {
                return m_DataSet;
            }

            public DataTable GetDataTable()
            {
                return m_DataTable;
            }

            public bool SetParam(string sParamName, object oValue)
            {
                m_ParamArray.Add(oValue);
                m_ParamName.Add(sParamName);
                m_ParamCount += 1;
                return true;
            }

            public bool BeginTransaction()
            {
                try
                {
                    m_Connection.ConnectionString = GenRstTab.SqlConnectionStr();
                    m_Connection.Open();
                    m_trans = m_Connection.BeginTransaction();
                    return true;
                }
                catch (Exception ex)
                {
                    m_ErrorMsg = ex.Message.ToString();
                    return false;
                }
            }

            public bool CommitTransaction()
            {
                try
                {
                    m_trans.Commit();
                    m_Connection.Close();
                    return true;
                }
                catch (Exception ex)
                {
                    m_ErrorMsg = ex.Message.ToString();
                    return false;
                }
            }

            public bool RollbackTransaction()
            {
                try
                {
                    m_trans.Rollback();
                    m_Connection.Close();
                    return true;
                }
                catch (Exception ex)
                {
                    m_ErrorMsg = ex.Message.ToString();
                    return false;
                }
            }

            public int RunStoredProcDataSet()
            {

                System.Type sysType = null;
                SqlConnection cnnConnection = new SqlConnection();
                SqlCommand cmd = null;
                SqlParameter param = null;
                string sName = null;
                int i = 0;
                int iRowCount = 0;

                cnnConnection.ConnectionString = GenRstTab.SqlConnectionStr();
                cnnConnection.Open();
                cmd = cnnConnection.CreateCommand();
                cmd.CommandTimeout = 0;
                try
                {
                    for (i = 0; i < m_ParamCount; i++)
                    {
                        sysType = m_ParamArray[i].GetType();
                        sName = m_ParamName[i].ToString();
                        param = cmd.Parameters.AddWithValue(sName, m_ParamArray[i]);
                        param.Value = m_ParamArray[i];
                    }

                    cmd.CommandText = SetProcName(m_StoredProcName);
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;
                    cmd.Prepare();

                    SqlDataAdapter dAdapter = new SqlDataAdapter(cmd);
                    DataSet ds = new DataSet("returnedDS");

                    dAdapter.Fill(ds); //This actually runs the stored proc and fills the dataset all at once.

                    cnnConnection.Close();
                    m_DataSet = ds;

                    if (m_DataSet.Tables.Count > 0)
                    {
                        iRowCount = m_DataSet.Tables[0].Rows.Count;
                    }
                    else
                    {
                        iRowCount = -1;
                    }

                }
                catch (Exception ex)
                {
                    m_ParamCount = 0;
                    iRowCount = -1;
                    m_ErrorMsg = ex.Message.ToString();
                }
                finally
                {
                    m_ParamCount = 0;
                }
                //Added in stuff to clear the proc name & parameters
                //as when running more then 1 proc etc the array values were been retained from previous procs run
                //at this stage done below
                m_ParamName.Clear();
                m_ParamArray.Clear();

                return iRowCount;

            }

            public int RunStoredProcDataSet(string sConnectionString)
            {

                System.Type sysType = null;
                SqlConnection cnnConnection = new SqlConnection();
                SqlCommand cmd = null;
                SqlParameter param = null;
                string sName = null;
                int i = 0;
                int iRowCount = 0;

                cnnConnection.ConnectionString = sConnectionString;
                cnnConnection.Open();
                cmd = cnnConnection.CreateCommand();
                cmd.CommandTimeout = 0;
                try
                {
                    for (i = 0; i < m_ParamCount; i++)
                    {
                        sysType = m_ParamArray[i].GetType();
                        sName = m_ParamName[i].ToString();
                        param = cmd.Parameters.AddWithValue(sName, m_ParamArray[i]);
                        param.Value = m_ParamArray[i];
                    }

                    cmd.CommandText = SetProcName(m_StoredProcName);
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;
                    cmd.Prepare();
                    cmd.ExecuteNonQuery();

                    SqlDataAdapter dAdapter = new SqlDataAdapter(cmd);
                    DataSet ds = new DataSet(cmd.CommandText);
                    DataTable dt = new DataTable(cmd.CommandText);

                    dAdapter.Fill(ds);
                    dAdapter.Fill(dt);

                    cnnConnection.Close();
                    m_DataSet = ds;
                    m_DataTable = dt;

                    if (m_DataSet.Tables.Count > 0)
                    {
                        iRowCount = m_DataSet.Tables[0].Rows.Count;
                    }
                    else
                    {
                        iRowCount = -1;
                    }

                }
                catch (Exception ex)
                {
                    m_ParamCount = 0;
                    iRowCount = -1;
                    m_ErrorMsg = ex.Message.ToString();
                }
                finally
                {
                    m_ParamCount = 0;
                }
                //Added in stuff to clear the proc name & parameters
                //as when running more then 1 proc etc the array values were been retained from previous procs run
                //at this stage done below
                m_ParamName.Clear();
                m_ParamArray.Clear();

                return iRowCount;

            }

            public int RunStoredProc()
            {

                System.Type sysType = null;
                SqlConnection cnnConnection = new SqlConnection();
                SqlCommand cmd = null;
                SqlParameter param = null;
                string sName = null;
                int i = 0;
                int iRowcount = 0;
                cnnConnection.ConnectionString = GenRstTab.SqlConnectionStr();
                cnnConnection.Open();
                cmd = cnnConnection.CreateCommand();
                cmd.CommandTimeout = 0;
                try
                {
                    for (i = 0; i < m_ParamCount; i++)
                    {
                        sysType = m_ParamArray[i].GetType();
                        sName = m_ParamName[i].ToString();
                        param = cmd.Parameters.AddWithValue(sName, m_ParamArray[i]);
                        param.Value = m_ParamArray[i];
                    }

                    cmd.CommandText = SetProcName(m_StoredProcName);
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;
                    cmd.Prepare();
                    iRowcount = cmd.ExecuteNonQuery();
                    cnnConnection.Close();
                }
                catch (Exception ex)
                {
                    iRowcount = -1;
                    m_ParamCount = 0;
                    m_ErrorMsg = ex.Message.ToString();
                }
                finally
                {
                    m_ParamCount = 0;
                }

                m_ParamName.Clear();
                m_ParamArray.Clear();
                return iRowcount;

            }

            public int RunStoredProc_OpenConnection()
            {

                System.Type sysType = null;
                SqlConnection cnnConnection = new SqlConnection();
                SqlCommand cmd = null;
                SqlParameter param = null;
                string sName = null;
                int i = 0;
                int iRowcount = 0;
                cnnConnection = m_Connection;
                cmd = cnnConnection.CreateCommand();
                cmd.CommandTimeout = 0;
                cmd.Transaction = m_trans;
                try
                {
                    for (i = 0; i < m_ParamCount; i++)
                    {
                        sysType = m_ParamArray[i].GetType();
                        sName = m_ParamName[i].ToString();
                        param = cmd.Parameters.AddWithValue(sName, m_ParamArray[i]);
                        param.Value = m_ParamArray[i];
                    }

                    cmd.CommandText = SetProcName(m_StoredProcName);
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;
                    cmd.Prepare();
                    iRowcount = cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    iRowcount = -1;
                    m_ParamCount = 0;
                    m_ErrorMsg = ex.Message.ToString();
                }
                finally
                {
                    m_ParamCount = 0;
                }

                m_ParamName.Clear();
                m_ParamArray.Clear();
                return iRowcount;

            }

            public int RunStoredProc(string sConnectionString)
            {

                System.Type sysType = null;
                SqlConnection cnnConnection = new SqlConnection();
                SqlCommand cmd = null;
                SqlParameter param = null;
                string sName = null;
                int i = 0;
                int iRowcount = 0;
                cnnConnection.ConnectionString = sConnectionString;
                cnnConnection.Open();
                cmd = cnnConnection.CreateCommand();
                cmd.CommandTimeout = 0;
                try
                {
                    for (i = 0; i < m_ParamCount; i++)
                    {
                        sysType = m_ParamArray[i].GetType();
                        sName = m_ParamName[i].ToString();
                        param = cmd.Parameters.AddWithValue(sName, m_ParamArray[i]);
                        param.Value = m_ParamArray[i];
                    }

                    cmd.CommandText = SetProcName(m_StoredProcName);
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;
                    cmd.Prepare();
                    iRowcount = cmd.ExecuteNonQuery();
                    cnnConnection.Close();
                }
                catch (Exception ex)
                {
                    iRowcount = -1;
                    m_ParamCount = 0;
                    m_ErrorMsg = ex.Message.ToString();
                }
                finally
                {
                    m_ParamCount = 0;
                }

                m_ParamName.Clear();
                m_ParamArray.Clear();
                return iRowcount;

            }

            public long RunStoredProcReturnValue()
            {

                System.Type sysType = null;
                SqlConnection cnnConnection = new SqlConnection();
                SqlCommand cmd = null;
                SqlParameter param = null;
                string sName = null;
                int i = 0;
                long iReturnValue = -1;
                cnnConnection.ConnectionString = GenRstTab.SqlConnectionStr();
                cnnConnection.Open();
                cmd = cnnConnection.CreateCommand();
                cmd.CommandTimeout = 0;
                try
                {
                    cmd.Parameters.Add("@RETVAL", SqlDbType.Int).Direction = ParameterDirection.ReturnValue;
                    for (i = 0; i < m_ParamCount; i++)
                    {
                        sysType = m_ParamArray[i].GetType();
                        sName = m_ParamName[i].ToString();
                        param = cmd.Parameters.AddWithValue(sName, m_ParamArray[i]);
                        param.Value = m_ParamArray[i];
                    }

                    cmd.CommandText = SetProcName(m_StoredProcName);
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;
                    cmd.Prepare();
                    cmd.ExecuteNonQuery();
                    iReturnValue = (int)cmd.Parameters["@RETVAL"].Value;

                }
                catch (Exception ex)
                {
                    iReturnValue = -1;
                    m_ParamCount = 0;
                    m_ErrorMsg = ex.Message.ToString();
                }
                finally
                {
                    m_ParamCount = 0;
                }

                m_ParamName.Clear();
                m_ParamArray.Clear();
                return iReturnValue;

            }

            public string RunStoredProcReturnValueVarChar()
            {

                System.Type sysType = null;
                SqlConnection cnnConnection = new SqlConnection();
                SqlCommand cmd = null;
                SqlParameter param = null;
                string sName = null;
                int i = 0;
                string sReturnValue = "";
                cnnConnection.ConnectionString = GenRstTab.SqlConnectionStr();
                cnnConnection.Open();
                cmd = cnnConnection.CreateCommand();
                cmd.CommandTimeout = 0;
                try
                {
                    cmd.Parameters.Add("@RETVAL", SqlDbType.Int).Direction = ParameterDirection.ReturnValue;
                    for (i = 0; i < m_ParamCount; i++)
                    {
                        sysType = m_ParamArray[i].GetType();
                        sName = m_ParamName[i].ToString();
                        param = cmd.Parameters.AddWithValue(sName, m_ParamArray[i]);
                        param.Value = m_ParamArray[i];
                    }

                    cmd.CommandText = SetProcName(m_StoredProcName);
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;
                    cmd.Prepare();
                    cmd.ExecuteNonQuery();
                    sReturnValue = (string)cmd.Parameters["@RETVAL"].Value;

                }
                catch (Exception ex)
                {
                    sReturnValue = "";
                    m_ParamCount = 0;
                    m_ErrorMsg = ex.Message.ToString();
                }
                finally
                {
                    m_ParamCount = 0;
                }

                m_ParamName.Clear();
                m_ParamArray.Clear();
                return sReturnValue;

            }
        }
}