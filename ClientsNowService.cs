using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Net;
using System.Data.SqlClient;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ClientsNowService
{
    public partial class ClientsNowService : ServiceBase
    {
        System.Timers.Timer timer;
        bool logEnabled = true;

        //Excel.Application oXL;
        //Excel.Workbooks oWBs;
        //Excel._Workbook oWB;
        //Excel._Worksheet oSheet;

        public const string connectionString = @"Server=SERVIDOR;Database=colibri8;User Id=miguelqm;Password=migSql@#";
        string[] strFeriados = {"01-12-17", 
                                "27-02-17", 
                                "28-02-17", 
                                "14-04-17", 
                                "21-04-17", 
                                "23-04-17", 
                                "01-05-17", 
                                "15-06-17", 
                                "07-09-17", 
                                "08-09-17", 
                                "12-10-17", 
                                "13-10-17", 
                                "02-11-17", 
                                "03-11-17", 
                                "15-11-17", 
                                "25-12-17" };

        private double lastProj = 0, lastProjFinal = 0, lastTicket = 0, lastDifMed = 0, lastDifMax = 0;
        private DateTime[] arrayFeriados;

        string tempHtmlFile = Path.Combine(Path.GetTempPath(), @"D:\Avenca\ClientsNowService\clients_now_tmp.html");

        public ClientsNowService()
        {
            InitializeComponent();

            eventLog1 = new System.Diagnostics.EventLog();
            if (!System.Diagnostics.EventLog.SourceExists("ClientesNowService"))
            {
                System.Diagnostics.EventLog.CreateEventSource(
                    "ClientesNowService", "Avenca");
            }
            eventLog1.Source = "ClientesNowService";
            eventLog1.Log = "Avenca";
        }

        protected override void OnStart(string[] args)
        {
            try
            {
                eventLog1.WriteEntry("Start", EventLogEntryType.Information, 0);

                //initExcel();
                logFichasAbertas(DateTime.Now.Millisecond);
                //cleanupExcel();
                timer = new System.Timers.Timer();
                timer.Interval = 800;
                timer.Elapsed += new System.Timers.ElapsedEventHandler(this.OnTimer);
                timer.Start();
                getData();
            }
            catch(Exception ex)
            {
                
                eventLog1.WriteEntry("ERROR: OnStart - " + ex.Message, EventLogEntryType.Error, 10);
                //cleanupExcel();
            }
        }

        protected override void OnStop()
        {
            //cleanupExcel();

            eventLog1.WriteEntry("Stop.", EventLogEntryType.Information, 100);
        }

        private void OnTimer(object sender, EventArgs e)
        {
            //eventLog1.WriteEntry("Timer click", EventLogEntryType.Information, eventId++);

            TimeSpan t = DateTime.Now.Subtract(DateTime.Parse("18:59"));
            if (t.TotalMinutes > 0)
            {
                if (timer.Interval == 60000)
                    timer.Interval = 800;
                //eventLog1.WriteEntry(t.TotalMinutes.ToString() + " minutos - NOT RUN", EventLogEntryType.Information, 1);
                return;
            }                
            t = DateTime.Now.Subtract(DateTime.Parse("11:30"));
            if (t.TotalMinutes > 0)
            {
                if (DateTime.Now.Second == 59)
                {
                    timer.Stop();
                    timer.Interval = 60000;
                    timer.Start();
                }
                //eventLog1.WriteEntry(t.TotalMinutes.ToString() + " minutos - RUNNING", EventLogEntryType.Information, 2);
                if (timer.Interval == 60000) 
                    getData();
            }
        }

        private void initializeListFeriados()
        {
            List<DateTime> termsList = new List<DateTime>();
            for (int i = 0; i < strFeriados.Length; i++)
            {
                termsList.Add(DateTime.ParseExact(strFeriados[i], "dd-MM-yy", CultureInfo.InvariantCulture));
            }
            arrayFeriados = termsList.ToArray();
        }

        private bool isWeekDay()
        {
            if (arrayFeriados == null)
                initializeListFeriados();

            if (arrayFeriados.Contains(DateTime.Today) ||
                DateTime.Today.DayOfWeek == DayOfWeek.Saturday ||
                DateTime.Today.DayOfWeek == DayOfWeek.Sunday ||
                (DateTime.Today.AddDays(1).DayOfWeek == DayOfWeek.Saturday && arrayFeriados.Contains(DateTime.Today.AddDays(-1))) ||
                (DateTime.Today.AddDays(-1).DayOfWeek == DayOfWeek.Sunday && arrayFeriados.Contains(DateTime.Today.AddDays(1)))
                )
                return false;
            return true;
        }

        private void getData()
        {
            //eventLog1.WriteEntry("GetData", EventLogEntryType.Information, 3);

            try
            {
                string queryString = File.ReadAllText(@"D:\Avenca\ClientsNowService\Previsao.sql", Encoding.UTF8);

                TimeSpan t = DateTime.Now.Subtract(DateTime.Parse("18:00"));
                if(t.TotalMinutes > 0)
                {
                    queryString = queryString.Replace("#STARTTIME#", "17:30");
                    queryString = queryString.Replace("#ENDTIME#", "23:59");
                }
                else
                {
                    queryString = queryString.Replace("#STARTTIME#", "11:00");
                    queryString = queryString.Replace("#ENDTIME#", "17:00");
                }
                queryString = (isWeekDay() ? queryString.Replace('#', '<') : queryString.Replace('#', '='));
                //eventLog1.WriteEntry("SQL: " + queryString, EventLogEntryType.Error, 5);

                using (SqlConnection connection =
                    new SqlConnection(connectionString))
                {

                SqlCommand command = new SqlCommand(queryString, connection);
                    connection.Open();
                    SqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        logFichasAbertas((int)reader["fichas_abertas"]);

                        replace(reader["tickets"].ToString(),
                                reader["proj"].ToString(),
                                reader["dif_media"].ToString(),
                                reader["dif_max"].ToString(),
                                reader["media_antes"].ToString(),
                                reader["min_antes"].ToString(),
                                reader["max_antes"].ToString(),
                                reader["media_depois"].ToString(),
                                reader["min_depois"].ToString(),
                                reader["max_depois"].ToString(),
                                reader["valor"].ToString(),
                                reader["media_valor"].ToString(),
                                reader["min_valor"].ToString(),
                                reader["max_valor"].ToString(),
                                reader["desconto"].ToString(),
                                reader["desc_globo"].ToString(),
                                reader["n_globo"].ToString(),
                                reader["n_desc_globo"].ToString(),
                                reader["n_desc_outros"].ToString(),
                                reader["desc_outros"].ToString(),
                                reader["comida"].ToString(),
                                reader["bebida"].ToString(),
                                reader["sobremesa"].ToString(),
                                reader["bomboniere"].ToString(),
                                reader["ticket_medio"].ToString(),
                                reader["media_globo"].ToString(),
                                reader["total_mes"].ToString(),
                                reader["media_comida"].ToString(),
                                reader["comida_por_pessoa"].ToString(),
                                reader["fichas_abertas"].ToString(),
                                reader["media_quilo"].ToString());                        ;

                        Object[] values = new Object[reader.FieldCount];
                        reader.GetValues(values);
                    }
                    reader.Close();

                    ftpUpload();

                    //eventLog1.WriteEntry(DateTime.Now.ToShortTimeString() + " - OK!", EventLogEntryType.Information, 3);
                }
            }
            catch (Exception ex)
            {
                eventLog1.WriteEntry("ERROR: getData - " + ex.Message, EventLogEntryType.Error, 10);
            }
        }

        private string compareToLast(double value, ref double lastValue)
        {
            string result = value.ToString();

            try
            {
                //result = result.IndexOf(',') < 0 ? result : result.Substring(0, result.IndexOf(',') + 3);

                string lastResult = lastValue.ToString();
                //lastResult = lastResult.IndexOf(',') < 0 ? lastResult : lastResult.Substring(0, lastResult.IndexOf(',') + 3);

                if (value > lastValue)
                    result = result + "<b><font color=\"green\">&#x25B2;</font></b><font size=\"6\">(" + lastResult + ")</font>";
                else if (value == lastValue)
                    result = result + " <font size=\"6\">(" + lastResult + ")</font>";
                else result = result + "<b><font color=\"red\">&#x25BC;</font></b><font size=\"6\">(" + lastResult + ")</font>";

                lastValue = value;
            }
            catch (Exception ex)
            {
                eventLog1.WriteEntry("ERROR: compareToLast - " + ex.Message, EventLogEntryType.Error, 10);
            }
            return result;
        }

        private void replace(string tickets, string proj, string dif_media, string dif_max, string media_antes, string min_antes, string max_antes, string media_depois, string min_depois, string max_depois,
                                string valor, string media_valor, string min_valor, string max_valor, string desconto, string desc_globo, string n_globo, string n_desc_globo, string n_desc_outros, string desc_outros,
                                string comida, string bebida, string sobremesa, string bomboniere, string ticket_medio, string media_globo, string total_mes, string media_comida, string comida_por_pessoa, string fichas_abertas, string media_quilo)
        {
            //eventLog1.WriteEntry("Replace", EventLogEntryType.Information, 4);
            string line = "0";

            try
            {
                string text = File.ReadAllText(@"D:\Avenca\ClientsNowService\avenca_clients.html");

                string data = DateTime.Now.ToShortDateString();
                string hora = DateTime.Now.ToString("HH:mm:ss");// ToShortTimeString();

                string projecao;

                if (ticket_medio != "" && proj != "")
                {
                    projecao = compareToLast(double.Parse(ticket_medio) * Int32.Parse(proj), ref lastProjFinal);
                    ticket_medio = compareToLast(double.Parse(ticket_medio), ref lastTicket);
                }
                else projecao = "";

                proj = compareToLast(double.Parse(proj), ref lastProj);

                //if (double.Parse(dif_media)>= 0)
                //    dif_media = "<font color=\"green\">" + dif_media + "</font>";
                //else dif_media = "<font color=\"red\">" + dif_media + "</font>";
                //if (double.Parse(dif_max) >= 0)
                //    dif_max = "<font color=\"green\">" + dif_max + "</font>";
                //else dif_max = "<font color=\"red\">" + dif_max + "</font>";
                
                double medComida = double.Parse(media_comida);
                double medQuilo = double.Parse(media_quilo);

                dif_media = compareToLast(double.Parse(dif_media), ref lastDifMed);
                dif_max = compareToLast(double.Parse(dif_max), ref lastDifMax);
                comida = compareToLast(double.Parse(comida), ref medComida);
                comida_por_pessoa = compareToLast(double.Parse(comida_por_pessoa), ref medQuilo);

                text = text.Replace("@data", data);
                text = text.Replace("@time", hora);
                text = text.Replace("@tickets", tickets);
                text = text.Replace("@proj", proj.Replace('.', ','));
                text = text.Replace("@dif_media", dif_media);
                text = text.Replace("@dif_max", dif_max);
                text = text.Replace("@media_antes", media_antes);
                text = text.Replace("@min_antes", min_antes);
                text = text.Replace("@max_antes", max_antes);
                text = text.Replace("@media_depois", media_depois);
                text = text.Replace("@min_depois", min_depois);
                text = text.Replace("@max_depois", max_depois);
                text = text.Replace("@n_globo", n_globo.Replace('.', ','));
                text = text.Replace("@n_desc_globo", n_desc_globo.Replace('.', ','));
                text = text.Replace("@n_desc_outros", n_desc_outros.Replace('.', ','));
                text = text.Replace("@media_globo", media_globo.Replace('.', ','));
                text = text.Replace("@ticket_medio", ticket_medio.Replace('.', ','));
                text = text.Replace("@bebida", bebida);
                text = text.Replace("@sobremesa", sobremesa);
                text = text.Replace("@bomboniere", bomboniere);
                text = text.Replace("@comida_por_pessoa", comida_por_pessoa.Replace('.', ','));
                text = text.Replace("@comida", comida.Replace('.', ','));
                text = text.Replace("@pr_valor", projecao.Replace('.', ','));
                //text = text.Replace("@media_comida", media_comida);

                text = text.Replace("@valor", valor.Replace('.', ','));
                text = text.Replace("@media_valor", media_valor.Replace('.', ','));
                text = text.Replace("@min_valor", min_valor.Replace('.', ','));
                text = text.Replace("@max_valor", max_valor.Replace('.', ','));
                text = text.Replace("@desconto", desconto.Replace('.', ','));
                text = text.Replace("@desc_globo", desc_globo.Replace('.', ','));
                text = text.Replace("@desc_outros", desc_outros.Replace('.', ','));
                text = text.Replace("@total_mes", total_mes.Replace('.', ','));
                text = text.Replace("@fichas_abertas", fichas_abertas);

                //try
                //{
                //    line = "1";
                //    text = text.Replace("@valor", valor.Substring(0, valor.IndexOf(',') + 3));
                //    line = "2";
                //    text = text.Replace("@media_valor", media_valor.Substring(0, media_valor.IndexOf(',') + 3));
                //    line = "3";
                //    text = text.Replace("@min_valor", min_valor.Substring(0, min_valor.IndexOf(',') + 3));
                //    line = "4";
                //    text = text.Replace("@max_valor", max_valor.Substring(0, max_valor.IndexOf(',') + 3));
                //    line = "5";
                //    text = text.Replace("@desconto", desconto.Substring(0, desconto.IndexOf(',') + 3));
                //    line = "6";
                //    text = text.Replace("@desc_globo", desc_globo.Substring(0, desc_globo.IndexOf(',') + 3));
                //    line = "7";
                //    text = text.Replace("@desc_outros", desc_outros.Substring(0, desc_outros.IndexOf(',') + 3));
                //    line = "8";
                //    text = text.Replace("@total_mes", total_mes.Substring(0, desc_outros.IndexOf(',') + 3));
                //}
                //catch (Exception e)
                //{
                //    eventLog1.WriteEntry("ERROR: replace(sub) - line: " + line + " - " + e.Message, EventLogEntryType.Error, 10);
                //}

                File.WriteAllText(tempHtmlFile, text);
                //File.Delete(tempHtmlFile);
            }
            catch (Exception ex)
            {
                eventLog1.WriteEntry("ERROR: replace - line: " + line + " - " + ex.Message, EventLogEntryType.Error, 10);
            }
        }

        void ftpUpload()
        {
            //eventLog1.WriteEntry("ftpUpload", EventLogEntryType.Information, 5);
            try
            {
                FtpWebRequest request = (FtpWebRequest)WebRequest.Create("ftp://avencarestaurante@ftp.avencarestaurante.hospedagemdesites.ws/Web/avenca_clients_now.html");
                request.Method = WebRequestMethods.Ftp.UploadFile;

                request.Credentials = new NetworkCredential("avencarestaurante", "ftpAvenca15");

                StreamReader sourceStream = new StreamReader(tempHtmlFile);
                byte[] fileContents = Encoding.UTF8.GetBytes(sourceStream.ReadToEnd());
                sourceStream.Close();
                request.ContentLength = fileContents.Length;

                Stream requestStream = request.GetRequestStream();
                requestStream.Write(fileContents, 0, fileContents.Length);
                requestStream.Close();

                FtpWebResponse response = (FtpWebResponse)request.GetResponse();

                response.Close();

                File.Delete(tempHtmlFile);
            }
            catch (Exception ex)
            {
                eventLog1.WriteEntry("ERROR: ftpUpload - " + ex.Message, EventLogEntryType.Error, 10);
            }
        }

        //private void initExcel()
        //{
        //    oXL = new Microsoft.Office.Interop.Excel.Application();
        //    oXL.DisplayAlerts = true;
                              
        //    //if (File.Exists(@"D:\Avenca\ClientsNowService\LogFichasAbertas.xlsx"))
        //    //    oWB = (Excel._Workbook)(oXL.Workbooks.Open(@"D:\Avenca\ClientsNowService\LogFichasAbertas.xlsx"));
        //    //else
        //        oWB = (Excel._Workbook)(oXL.Workbooks.Add(Type.Missing));

        //    oSheet = (Excel._Worksheet)oWB.ActiveSheet;
        //}

        //private void cleanupExcel()
        //{
        //    oWB.Close();
        //    oXL.Quit();
        //    Marshal.ReleaseComObject(oWB);
        //    Marshal.ReleaseComObject(oXL);
        //}
        private void logFichasAbertas(int nFichas)
        {
            try
            {
                if (!logEnabled)
                    return;

                if ((DateTime.Now.Hour > 16) && (nFichas == 0))
                    logEnabled = false;

                //string path = @"D:\Avenca\ClientsNowService\logFichasAbertas" + DateTime.Now.ToString("yyMMdd") + ".txt";
                //// This text is added only once to the file.
                //if (!File.Exists(path))
                //{
                //    // Create a file to write to.
                //    using (StreamWriter sw = File.CreateText(path))
                //    {
                //    }
                //}
                //using (StreamWriter sw = File.AppendText(path))
                //{
                //    sw.WriteLine(DateTime.Now.ToString().Replace(' ', '\t') + "\t" + nFichas.ToString());
                //}

                //logFichasAbertasExcel(DateTime.Now.Date, DateTime.Now.ToString("HH:mm:00"), nFichas);
                //logFichasAbertasExcel(dt, "12:44", 56);
                logFichasAbertasExcel(DateTime.Now.Date, DateTime.Now.ToString("HH:mm"), nFichas);
            }
            catch(Exception ex)
            {
                eventLog1.WriteEntry("ERROR: logFichasAbertas - " + ex.Message, EventLogEntryType.Error, 10);
            }
        }

        //private void logFichasAbertasExcel(DateTime pDate, string pTime, int pFichas)
        //{
        //    Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
        //    Excel.Workbooks oWBs = oXL.Workbooks;
        //    Excel._Workbook oWB;
        //    Excel._Worksheet oSheet;

        //    string a = "1";
        //    if (File.Exists(@"LogFichasAbertas.xlsx"))
        //        oWB = oWBs.Open(@"LogFichasAbertas.xlsx");
        //    else
        //        oWB = oWBs.Add(Type.Missing);
        //    a = "2";
        //    oSheet = oWB.ActiveSheet;
        //    a = "3";
        //    Excel.Range r = oSheet.Range["A:A"];
        //    Excel.Range found = null;
        //    Excel.Range endRange = null;
        //    Excel.Range rowRange = null;
        //    Excel.Range cells = oSheet.Cells;
        //    Excel.Range cellDate = null;
        //    Excel.Range cellFicha = null;
        //    Excel.Font font = null;
        //    a = "4";
        //    try
        //    {
        //        int col = 2;
        //        a = "5";
        //        found = r.Find(pDate.ToString("yyMMdd"), SearchOrder: Excel.XlSearchOrder.xlByRows, LookIn: Excel.XlFindLookIn.xlValues);
        //        if (found == null)
        //        {
        //            a = "6";
        //            endRange = r.End[Excel.XlDirection.xlToRight];
        //            col = endRange.Column > 16000 ? 2 : endRange.Column + 1;
        //            a = "7";
        //            cellDate = cells[1, col];
        //            cellDate.Value = pDate.ToString("yyMMdd");
        //            font = cellDate.Font;
        //            font.Bold = true;
        //            cellDate.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
        //            a = "8";
        //        }
        //        else
        //            col = found.Column;
        //        a = "9";
        //        rowRange = r.Find(pTime, SearchOrder: Excel.XlSearchOrder.xlByColumns);
        //        int row = 14;// rowRange.Row;
        //        a = "10";
        //        cellFicha = cells[row, col];
        //        cellFicha.Value = pFichas;
        //        a = "11";
        //        oXL.DisplayAlerts = false;
        //        oWB.SaveAs(@"LogFichasAbertas.xlsx", ConflictResolution: Excel.XlSaveConflictResolution.xlLocalSessionChanges);
        //        eventLog1.WriteEntry("SUCCESS! logFichasAbertasExcel - " + a, EventLogEntryType.Error, 10);
        //    }
        //    catch (Exception ex)
        //    {
        //        eventLog1.WriteEntry("ERROR: logFichasAbertasExcel - " + a + " -- " +ex.Message, EventLogEntryType.Error, 10);
        //    }
        //    finally
        //    {
        //        oWB.Close();
        //        oWBs.Close();
        //        oXL.Quit();

        //        releaseObject(font);
        //        releaseObject(found);
        //        releaseObject(r);
        //        releaseObject(cells);
        //        releaseObject(cellFicha);
        //        releaseObject(cellDate);
        //        releaseObject(endRange);
        //        releaseObject(rowRange);

        //        releaseObject(oSheet);
        //        releaseObject(oWB);
        //        releaseObject(oWBs);
        //        releaseObject(oXL);
        //    }
        //}

        //private void releaseObject(object obj)
        //{
        //    if (obj == null) return;

        //    try
        //    {
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
        //        obj = null;
        //    }
        //    catch (Exception ex)
        //    {
        //        obj = null;
        //        eventLog1.WriteEntry("ERROR: releaseObject - " + ex.Message, EventLogEntryType.Error, 10);
        //    }
        //    finally
        //    {
        //        GC.Collect();
        //    }
        //}

        private void logFichasAbertasExcel(DateTime pDate, string pTime, int pFichas)
        {
            Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbooks oWBs = oXL.Workbooks;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;

            if (File.Exists(@"D:\Avenca\LogFichasAbertas.xlsx"))
                oWB = oWBs.Open(@"D:\Avenca\LogFichasAbertas.xlsx");
            else
                oWB = oWBs.Add(Type.Missing);

            oSheet = oWB.ActiveSheet;

            Excel.Range r = oSheet.Range["A1"];
            Excel.Range found = null;
            Excel.Range endRange = null;
            Excel.Range rowRange = null;
            Excel.Range cells = oSheet.Cells;
            Excel.Range cellDate = null;
            Excel.Range cellFicha = null;
            Excel.Font font = null;

            try
            {
                int col =2;

                found = r.Find(pDate.ToString("yyMMdd"), SearchOrder: Excel.XlSearchOrder.xlByRows, LookIn: Excel.XlFindLookIn.xlValues);
                if (found == null)
                {
                    endRange = r.End[Excel.XlDirection.xlToRight];
                    col = endRange.Column > 16000 ? 2 : endRange.Column + 1;

                    cellDate = cells[1, col];
                    cellDate.Value = pDate.ToString("yyMMdd");
                    font = cellDate.Font;
                    font.Bold = true;
                    cellDate.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                }
                else
                    col = found.Column;

                rowRange = r.Find(pTime, SearchOrder: Excel.XlSearchOrder.xlByColumns);
                int row = rowRange.Row;

                cellFicha = cells[row, col];
                cellFicha.Value = pFichas;

                oXL.DisplayAlerts = false;
                oWB.SaveAs(@"D:\Avenca\LogFichasAbertas.xlsx", ConflictResolution: Excel.XlSaveConflictResolution.xlLocalSessionChanges);
            }
            catch (Exception ex)
            {
                eventLog1.WriteEntry("ERROR: logFichasAbertasExcel - " + ex.Message, EventLogEntryType.Error, 10);
            }
            finally
            {
                oWB.Close();
                oWBs.Close();
                oXL.Quit();

                releaseObject(font);
                releaseObject(found);
                releaseObject(r);
                releaseObject(cells);
                releaseObject(cellFicha);
                releaseObject(cellDate);
                releaseObject(endRange);
                releaseObject(rowRange);
                
                releaseObject(oSheet);
                releaseObject(oWB);
                releaseObject(oWBs);
                releaseObject(oXL);
            }
        }

        private void releaseObject(object obj)
        {
            if (obj == null) return;

            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                eventLog1.WriteEntry("ERROR: releaseObject - " + ex.Message, EventLogEntryType.Error, 10);
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
