using System;
using System.Configuration;
using System.Text;
using ClosedXML.Excel;
using System.Data;
using SAPbobsCOM;
using System.IO;
using Newtonsoft.Json;
using System.Data.Odbc;



namespace ConsoleApp2
{
    class Program
    {
        public static SAPbobsCOM.Company oCompany;
        static OdbcConnection CnnHANA;
        public static String DataBaseName = ConfigurationManager.AppSettings["CompanyDB"];



        static void Main(string[] args)


        {
            Conexion();
            //ConecctOdbc();
            String consulta;
            var respuesta = false;
            Recordset oRecord = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);



            #region Recorrer Parametros SAP
            try
            {
                consulta = "SELECT T0.\"Code\", T0.\"U_SOL_VISTA\", T0.\"U_SOL_ARCHIVO\", T0.\"U_SOL_RUTA\", T0.\"U_SOL_FORMATO\", T0.\"U_SOL_ACTIVO\", T0.\"U_SOL_AHORA\" FROM \"@SOL_EXPORT_VIEW\"  T0";

                DataTable Parametros = ObtenerParametos(consulta);
                try
                {
                    if (Parametros != null)
                    {
                        for(int i = 0; i < Parametros.Rows.Count; i++)
                        {
                            String QueryView = Parametros.Rows[i]["U_SOL_VISTA"].ToString();
                            String Formato = Parametros.Rows[i]["U_SOL_FORMATO"].ToString();
                            String RutaExport = Parametros.Rows[i]["U_SOL_RUTA"].ToString();
                            String NombreArchivo = Parametros.Rows[i]["U_SOL_ARCHIVO"].ToString();

                            if (!string.IsNullOrEmpty(QueryView) )
                            {

                                respuesta = ExportView($"Select * from " + QueryView, Formato, RutaExport, NombreArchivo);
                            }
                            
                        }
                    }
                }
                catch (Exception ex)
                {
                    EscribeLog("Error en conexion SAP" + ex.StackTrace.ToString());
                }

            }
            catch (Exception) 
            {
            }
            #endregion
        }

        static void EscribeLog(String Message)
        {
            var LogFilePath = ConfigurationManager.AppSettings["LogFilePath"];
            using (StreamWriter SW = new StreamWriter(LogFilePath, true))
            {
                SW.WriteLine(DateTime.Now + "|" + Message);
            }
        }


        #region conexionSAP
        static void Conexion()
        {
            int ErrCode = 0;
            var ErrMsg = "";
            try
            {
                oCompany = new SAPbobsCOM.Company();
                oCompany.Server = ConfigurationManager.AppSettings["Server"];
                oCompany.DbServerType = (BoDataServerTypes)Enum.Parse(typeof(BoDataServerTypes), ConfigurationManager.AppSettings["ServerType"]);
                oCompany.CompanyDB = ConfigurationManager.AppSettings["CompanyDB"];
                oCompany.UserName = ConfigurationManager.AppSettings["UserName"];
                oCompany.Password = ConfigurationManager.AppSettings["Password"];
                oCompany.language = BoSuppLangs.ln_Spanish_La;

                ErrCode = oCompany.Connect();
                if (ErrCode != 0)
                {
                    oCompany.GetLastError(out ErrCode, out ErrMsg);
                    EscribeLog("Error de conexion con SAP | " + ErrCode +" "+ ErrMsg);
                    RegistroLogsSAP("Conexion SAP", ErrCode + " " + ErrMsg);
                    
                }
                else
                {

                }
                   
            }
            catch (Exception ex)
            {
                RegistroLogsSAP("Conexion SAP", ex.StackTrace.ToString()) ;
                
            }

        }
        #endregion
        #region ConexionODBC
        static void ConecctOdbc()
        {

            string CadOdbc = ConfigurationManager.AppSettings["CadenaODBC"];
            
            try

            {

                if (IntPtr.Size != 8) 
                {
 
                    CadOdbc = CadOdbc.Replace("HDBODBC", "HDBODBC32");

                }
                else
                {
                    
                }

                CnnHANA = new OdbcConnection(CadOdbc); 
                if (CnnHANA.State == System.Data.ConnectionState.Closed) 
                {
                    CnnHANA.Open();
                    
                }
                else 
                {

                }

            }
            catch (Exception ex) 
            {
                EscribeLog("Error de conexion ODBC |" + ex.StackTrace.ToString());
            }

        }
        #endregion


        #region FUNobtenerParametros
        static DataTable ObtenerParametos(String query)
        {
            try
            {
                int ErrCode = 0;
                var ErrMsg = "";

                DataTable Parametros = new DataTable();
                Recordset oRecord = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                oRecord.DoQuery(query);
                if (query != null)
                {


                    for (int i = 0; i < oRecord.Fields.Count; i++)
                    {
                        Parametros.Columns.Add(oRecord.Fields.Item(i).Name);

                    }


                    {
                        while (!oRecord.EoF)
                        {
                            Parametros.Rows.Add(
                            oRecord.Fields.Item(0).Value.ToString(),
                            oRecord.Fields.Item(1).Value.ToString(),
                            oRecord.Fields.Item(2).Value.ToString(),
                            oRecord.Fields.Item(3).Value.ToString(),
                            oRecord.Fields.Item(4).Value.ToString(),
                            oRecord.Fields.Item(5).Value.ToString(),
                            oRecord.Fields.Item(6).Value.ToString()
                            );

                            oRecord.MoveNext();
                            //return "Ok";
                        }
                        return Parametros;
                    }
                }
                else
                {
                    oCompany.GetLastError(out ErrCode, out ErrMsg);
                    return null;
                }
            } catch (Exception e)
            {
                RegistroLogsSAP("Lectura de UDO", e.StackTrace.ToString());
                return null;
            }



        }
        #endregion

        static bool ExportarNExcel(DataTable dataTable, String ruta)
        {
            try
            {

                string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), ruta);

                using (XLWorkbook wb = new XLWorkbook())
                {
                    var ws = wb.Worksheets.Add(dataTable, "VistaExportada");

                    ws.Columns().AdjustToContents(); // Ajustar ancho de columnas
                    wb.SaveAs(filePath);
                }

                Console.WriteLine($"Archivo Excel creado exitosamente: {filePath}");
            }
            catch (Exception ex)
            {
                RegistroLogsSAP("Exportar a excel", ex.StackTrace.ToString());
            }
            return true;
        }
        #region FUNexportarArchivo
        static bool ExportView(String queryExport, String Format,  String FilePath, String FileName) 
        {
            var separadorCsv = ConfigurationManager.AppSettings["SeparadorCSV"];
            String FileExport = FilePath+"\\"+FileName+"."+Format;
            ConecctOdbc();
            DataTable dataTable = new DataTable();
            using (OdbcDataAdapter adapter = new OdbcDataAdapter(queryExport, CnnHANA))
            {
                adapter.Fill(dataTable);
            }
            
            try
            {
                if(dataTable != null)
                {
                    switch (Format)
                    {

                        case "JSON":
                            
                            string jsonString = JsonConvert.SerializeObject(dataTable, Formatting.Indented);

                            File.WriteAllText(FileExport, jsonString);
                            break;

                        case "TXT":

                            ExportarDataTable(dataTable, FileExport, "\t");



                            break;

                        case "CSV":

                            ExportarDataTable(dataTable, FileExport, separadorCsv);

                            break;

                        case "XLSX":

                            ExportarNExcel(dataTable,FileExport);
                            break;
                        
                        default: Console.WriteLine(Format);
                            break;


                    }


                }

                CnnHANA.Close();
            }
            catch (Exception ex)
            {
                RegistroLogsSAP("Exportar a excel", ex.StackTrace.ToString());
                CnnHANA.Close();
            }
            
            return false;
        }
        #endregion

        static void ExportarDataTable(DataTable dataTable, string rutaArchivo, string separador)
        {
            StringBuilder sb = new StringBuilder();

            // Escribir encabezados
            string[] columnNames = new string[dataTable.Columns.Count];
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                columnNames[i] = dataTable.Columns[i].ColumnName;
            }
            sb.AppendLine(string.Join(separador, columnNames));

            // Escribir filas
            foreach (DataRow row in dataTable.Rows)
            {
                string[] fields = new string[dataTable.Columns.Count];
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    fields[i] = row[i].ToString();
                }
                sb.AppendLine(string.Join(separador, fields));
            }

            // Guardar el archivo CSV
            File.WriteAllText(rutaArchivo, sb.ToString(), Encoding.UTF8);
        }

        static void RegistroLogsSAP(string Etapa, string Mensaje)
        {
            //var CadOdbc = ConfigurationManager.AppSettings["CadenaODBC"];
            var StoredProcedure = $"CALL \"{DataBaseName}\".\"SOL_SP_EXPORTVIEW\" ('{Etapa}', '{Mensaje}')";
            try
            {
                ConecctOdbc();
                if (CnnHANA.State != ConnectionState.Open)
                {
                    CnnHANA.Open();
                }
                OdbcCommand cmd = new OdbcCommand(StoredProcedure, CnnHANA);
                cmd.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                EscribeLog("Escribir log en tabla de SAP" + ex.StackTrace.ToString());
            }
            
        }

    }

}

