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
        public static OdbcConnection CnnHANA;

        static void Main(string[] args)


        {
            Conexion();
            ConecctOdbc();
            String consulta;
            var respuesta = false;
            Recordset oRecord = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);



            #region Recorrer Parametros SAP
            try
            {
                consulta = "SELECT T0.\"Code\", T0.\"U_SOL_VISTA\", T0.\"U_SOL_ARCHIVO\", T0.\"U_SOL_RUTA\", T0.\"U_SOL_FORMATO\", T0.\"U_SOL_ACTIVO\", T0.\"U_SOL_AHORA\" FROM \"@SOL_EXPORT_VIEW\"  T0";
                //Result = 
                //ObtenerParametos(consulta);
                //Console.WriteLine(Result);
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
                                //respuesta = ExportView($"SELECT * FROM \"OCTG\"" , Formato, RutaExport, NombreArchivo);
                                //respuesta = ExportView($"SELECT * FROM \"TEST_SURCOMPANY_270125\".OCTG" , Formato, RutaExport, NombreArchivo);
                                respuesta = ExportView($"Select * from " + QueryView, Formato, RutaExport, NombreArchivo);
                            }
                            
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.Write(ex.ToString());
                }

            }
            catch (Exception) 
            {
            }
            #endregion
        }

       

        #region conexionSAP
        static bool Conexion()
        {
            try
            {
                oCompany = new SAPbobsCOM.Company();
                oCompany.Server = ConfigurationManager.AppSettings["Server"];
                oCompany.DbServerType = (BoDataServerTypes)Enum.Parse(typeof(BoDataServerTypes), ConfigurationManager.AppSettings["ServerType"]);
                oCompany.CompanyDB = ConfigurationManager.AppSettings["CompanyDB"];
                oCompany.UserName = ConfigurationManager.AppSettings["UserName"];
                oCompany.Password = ConfigurationManager.AppSettings["Password"];
                oCompany.language = BoSuppLangs.ln_Spanish_La;

                if (oCompany.Connect() != 0)
                {
                    Console.WriteLine($"Error al conectar: {oCompany.GetLastErrorDescription()}");
                    return false;
                }
                else
                    return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }

        }
        #endregion
        #region ConexionODBC
        static string ConecctOdbc()
        {

            string CadOdbc = ConfigurationManager.AppSettings["CadenaODBC"];
            String R = "";
            
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
                if (CnnHANA.State == System.Data.ConnectionState.Open) 
                {
                    R = "OK";
                    
                }
                else 
                { 
                    R = "Error"; 
                }

            }
            catch (Exception) 
            {
                R = "Error";         
            }

            return R.ToString(); 
        }
        #endregion


        #region FUNobtenerParametros
        static DataTable ObtenerParametos(String query)
        {
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
                return null;
          

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
            catch
            {
                Exception ex;
            }
            return true;
        }
        #region FUNexportarArchivo
        static bool ExportView(String queryExport, String Format,  String FilePath, String FileName) 
        {
            var separadorCsv = ConfigurationManager.AppSettings["SeparadorCSV"];
            String FileExport = FilePath+"\\"+FileName+"."+Format;
            CnnHANA.Open();
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
            catch
            { 
            Exception exception;
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



    }

}

