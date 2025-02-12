using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using System.Threading.Tasks;
using System.Data;
using SAPbobsCOM; // SAP
using System.IO; // Exportar plano
using Newtonsoft.Json;
using Spire.Xls; // Exportar a Excel
using System.Data.Odbc;
using System.Security.Cryptography;
//using System.Data.OleDb; // Exportar csv


namespace ConsoleApp2
{
    class Program
    {
        public static SAPbobsCOM.Company oCompany;
        public static String cadena_txt;
        public static String cadena_csv; 
        public static OdbcConnection CnnHANA;

        static void Main(string[] args)


        {
            Conexion();
            ConecctOdbc();
            String consulta;
            String Result;
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
            catch (Exception ex) 
            {
            }
            #endregion

            /*Recordset oRecordset;
            if (Conexion())
            {
                Console.WriteLine("Exitoso");

                oRecordset = RecordSet();
                DataTable dataTable = ConvertRecordsetToDataTable(oRecordset);
                ExportJson(oRecordset, dataTable);
                ExportPlano();
                ExportarExcel(dataTable);

                Console.WriteLine("Proceso completado.");
               //oCompany.Disconnect();
            }*/
        }

        /*static Recordset RecordSet()
        {
            String query;
            Recordset oRecordset = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
               
                    query = "SELECT T0.\"ItemCode\", T0.\"ItemName\" FROM \"OITM\" T0 WHERE T0.\"QryGroup5\" ='Y'";
                    oRecordset.DoQuery(query);
                    */


        /* oRecordset.MoveFirst();
         while (!oRecordset.EoF)
         {
            // int docEntry = (int)oRecordset.Fields.Item("DocEntry").Value;
             ItemCode = oRecordset.Fields.Item(0).Value.ToString();
             ItemName = oRecordset.Fields.Item(1).Value.ToString();


         //Console.WriteLine(ItemCode+" "+ItemName);
             cadena = cadena+linea+" "+ItemCode + "\t" + ItemName + "\r\n";
             oRecordset.MoveNext();
             linea++;
         }*/

        /*
                return oRecordset;
            }
            catch (Exception err)
            {
                Console.WriteLine(err);
                return oRecordset;
            }
            return oRecordset;
        }*/

        #region conexionSAP
        static bool Conexion()
        {
            try
            {
                oCompany = new Company();
                oCompany.Server = "NDB@192.168.160.12:30013";
                oCompany.DbServerType = BoDataServerTypes.dst_HANADB;
                oCompany.CompanyDB = "TEST_SURCOMPANY_270125";
                oCompany.UserName = "manager";
                oCompany.Password = "HYC909";
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

            string CadOdbc = "";
            String R = "";
            
            try

            {

                if (IntPtr.Size == 8) //64 bits - diferente de 8 32 bits
                {
                    //CadOdbc = "DRIVER={HDBODBC};SERVERNODE=@solsetec;CS=SURCOMPANY;"; 
                    //CadOdbc = "Driver={HDBODBC};SERVERNODE=@solsetec;CS=SURCOMPANY;"; 
                     CadOdbc = "Driver={HDBODBC}; ServerNode = 192.168.160.12:30013; Uid=system; Pwd=Asdf1234$; databaseName=NDB"; 

                }
                else
                {
                    //CadOdbc = "Driver={HDBODBC32};SERVERNODE=@solsetec;CS=SURCOMPANY;";
                    CadOdbc = "Driver={HDBODBC32}; ServerNode = 192.168.160.12:30013; Uid=system; Pwd=Asdf1234$; databaseName=NDB";
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
            catch (Exception ex) 
            {
                R = "Error";         
            }

            return R.ToString(); 
        }
        #endregion



        /**
        * Crear Archivo Plano en TXT
        */
        /*static bool ExportPlano()
        {
            string path = @"c:\Temp\archivo_plano.txt";
            if (!File.Exists(path))
            {
                // Create a file to write to.
                using (StreamWriter sw = File.CreateText(path))
                {
                    sw.Write(cadena_txt);
                }
            }

            // Open the file to read from.
            using (StreamReader sr = File.OpenText(path))
            {
                string s = "";
                while ((s = sr.ReadLine()) != null)
                {
                    Console.WriteLine(s);
                }
                return true;
            }

            return true;
        }*/

        /*static bool ExportJson(Recordset oRecordset, DataTable dataTable)
        {
            string outputFilePath = @"c:\\temp\\archivo_json.json";
            try
            {
                // 2. Convertir el Recordset a DataTable
                //DataTable dataTable = ConvertRecordsetToDataTable(oRecordset);

                // 3. Serializar el DataTable a JSON
                string jsonString = JsonConvert.SerializeObject(dataTable, Formatting.Indented);

                // 4. Guardar la cadena JSON en un archivo
                File.WriteAllText(outputFilePath, jsonString);

                Console.WriteLine("Archivo JSON creado con éxito en: " + outputFilePath);

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }*/

        //static DataTable ConvertRecordsetToDataTable(Recordset recordset)
        //{
        //    DataTable dataTable = new DataTable();
            
        //    // Crear las columnas del DataTable
        //    for (int i = 0; i < recordset.Fields.Count; i++)
        //    {
        //        dataTable.Columns.Add(recordset.Fields.Item(i).Name);

        //    }

        //    // Rellenar el DataTable con los datos del Recordset
        //    while (!recordset.EoF)
        //    {
        //        DataRow row = dataTable.NewRow();
        //        for (int i = 0; i < recordset.Fields.Count; i++)
        //        {
        //            row[i] = recordset.Fields.Item(i).Value;
        //            if (i < recordset.Fields.Count+1)
        //                cadena_txt = cadena_txt + row[i] + "\t";
        //            else
        //                cadena_txt = cadena_txt + row[i];
        //        }
        //        dataTable.Rows.Add(row);
                
        //        cadena_txt = cadena_txt + "\r\n";
        //        recordset.MoveNext();
        //    }
        //    return dataTable;
        //}

        
        //static bool ExportarExcel(DataTable dataTable)
        //{
        //    string ruta = @"c:\\temp\\archivo_excel.xls";
        //    // Crear un nuevo libro de trabajo de Excel
        //    Workbook workbook = new Workbook();
        //    Worksheet sheet = workbook.Worksheets[0];

        //    // Insertar el DataTable en la hoja de cálculo
        //    sheet.InsertDataTable(dataTable, true, 1, 1);

        //    // Guardar el archivo de Excel
        //    workbook.SaveToFile(ruta, ExcelVersion.Version2013);
        //    return true;
        //}


        #region FUNobtenerParametros
        static DataTable ObtenerParametos(String query)
        {
            DataTable Parametros = new DataTable();
            String Query = "";
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

            
                // Crear un nuevo libro de trabajo de Excel
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
            bool bandera = true;
            String json = "";
            String FileExport = FilePath+"\\"+FileName+"."+Format;
            //Recordset ors = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            //ors.DoQuery(queryExport);
            CnnHANA.Open();
            CnnHANA.ChangeDatabase("TEST_SURCOMPANY_270125");
            DataTable dataTable = new DataTable();
            using (OdbcDataAdapter adapter = new OdbcDataAdapter(queryExport, CnnHANA))
            {
                 //reader2 = cmd.ExecuteReader();
                //reader2.Close();
                adapter.Fill(dataTable);
            }
                //DataTable dataTable = ConvertRecordsetToDataTable(ors);
            
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

                            ExportarDataTableACSV(dataTable, FileExport, "\t");



                            break;

                        case "CSV":

                            ExportarDataTableACSV(dataTable, FileExport, ";");

                            break;

                        case "XLSX":

                            //string ruta = @"c:\\temp\\ArchivoExcel.xlsx";
                            // Crear un nuevo libro de trabajo de Excel
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
            Exception ex;
                CnnHANA.Close();
            }
            
            return false;
        }
        #endregion

        static void ExportarDataTableACSV(DataTable dataTable, string rutaArchivo, string separador)
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

