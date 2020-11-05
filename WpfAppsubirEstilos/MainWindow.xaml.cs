using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
//using System.Windows.Shapes;

//--< using >-- 
using Microsoft.Win32; //FileDialog 
using WinForms = System.Windows.Forms; //FolderDialog 
using System.IO; //Folder, Directory 
using System.Diagnostics; //Debug.WriteLine 
using System.Data.OleDb;
using System.Data;
using System.ComponentModel;
using System.Threading;
//using System.Windows.Forms;
using System.Text.RegularExpressions;
using SistemaAuditores.DataAccess;

namespace WpfAppsubirEstilos
{
    /// <summary>
    /// Lógica de interacción para MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        class objMedidas
        {

            public int idMedida { get; set; }
            public string medida { get; set; }
            public decimal valor { get; set; }
        }

        class objtallas
        {

            public int idtalla { get; set; }
            public string talla { get; set; }
            public int valor { get; set; }


            public int idPOM { get; set; }
            public string puntomedida { get; set; }

            public int idmedidas { get; set; }
            public string medidas { get; set; }
            public decimal medidasdec { get; set; }

        }

        class objespecificaciones
        {
            public int idPOM { get; set; }
            public int idtalla { get; set; }
            public int idestilo { get; set; }
            public int rango { get; set; }
            public string Estilo { get; set; }
            public double valorspec { get; set; }
            public double valorMax { get; set; }
            public double valorMin { get; set; }
        }

        private BackgroundWorker _worker;
      //  private BackgroundWorker _worker1;

        private void CancelWorker(object sender, RoutedEventArgs e)
        {
            _worker.CancelAsync();
        }

        public MainWindow()
        {
            InitializeComponent();

            _worker = new BackgroundWorker();

            _worker.WorkerReportsProgress = true;

            _worker.DoWork += new System.ComponentModel.DoWorkEventHandler(_worker_DoWork);

            _worker.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(_worker_ProgressChanged);

            _worker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(_worker_RunWorkerCompleted);



           // _worker1.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(_worker_RunWorkerCompleted);
        }


        async Task<List<FileInfo>> GetFile(DirectoryInfo path)
        {
            return await Task.Run(() =>
            {
                var listFile = path.GetFiles("*.xlsx").ToList();
                return listFile;
            });
            //otra forma ****
            //System.IO.Directory.GetFiles(currentDirName, "*.txt");
        }

        private async void Button_Click(object sender, RoutedEventArgs e)
        {
            lblprogress.Content = "Cargando";
            WinForms.FolderBrowserDialog folderDialog = new WinForms.FolderBrowserDialog();
            folderDialog.ShowNewFolderButton = false;
            folderDialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            WinForms.DialogResult result = folderDialog.ShowDialog();

            if (result == WinForms.DialogResult.OK)
            {
                String sPath = folderDialog.SelectedPath;

                DirectoryInfo folder = new DirectoryInfo(sPath);

                var list = GetFile(folder);

                await Task.WhenAll(list);

                List<object> arguments = new List<object>();
                arguments.Add(list.Result);
                arguments.Add(sPath);

                _worker.RunWorkerAsync(arguments);

            }
        }

        void _worker_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            List<object> genericlist = e.Argument as List<object>;

            List<FileInfo> list = (List<FileInfo>)genericlist[0];
            string sPath = (string)genericlist[1];

            BackgroundWorker worker = sender as BackgroundWorker;

            StringBuilder TEXTO = new StringBuilder();
            TEXTO.AppendLine("*** XLS con problemas de lectura *** \n");
            TEXTO.AppendLine("--------------------------------------------\n");
            var errores = false;

            List<objMedidas> listaMedidas = new List<objMedidas>();

            var dtm = DataAccess.Get_DataTable("select * from tblMedidas order by valor desc");

            foreach (DataRow item in dtm.Rows)
            {
                var obj = new objMedidas
                {
                    idMedida = Convert.ToInt16(item[0].ToString()),
                    medida = item[1].ToString().Trim(),
                    valor = Convert.ToDecimal(item[2].ToString())
                };

                listaMedidas.Add(obj);

            }

            var val = 0; //list.Count;

            // var progress = (val - list.Count) * 100;

            foreach (var item in list)
            {
                val++;
                var progress = Convert.ToDouble(val) / list.Count * 100;

                List<objtallas> lista = new List<objtallas>();
                List<objespecificaciones> listaFinal = new List<objespecificaciones>();

                var est = item.Name.Split('.');

                var dt1 = DataAccess.Get_DataTable("select * from style where Style='" + est[0] + "'");

                if (dt1 == null || dt1.Rows.Count < 1)
                {
                    TEXTO.AppendLine("el estilo " + item + " no existe \n");
                    errores = true;
                }
                else
                {
                    try
                    {
                        #region    
                        DataTable dt = new DataTable();
                        DataTable dt2 = new DataTable();
                        DataTable dt3 = new DataTable();

                        string cadena = "estilo";

                        int idestilo = 0;

                        using (OleDbConnection conn = new OleDbConnection())
                        {

                            string Import_FileName = Path.Combine(sPath.ToString(), item.Name);
                            string fileExtension = Path.GetExtension(Import_FileName);
                            if (fileExtension == ".xls")
                                conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Import_FileName + ";" + "Extended Properties='Excel 8.0;HDR=NO;IMEX=1;'";
                            if (fileExtension == ".xlsx")
                                conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Import_FileName + ";" + "Extended Properties='Excel 12.0 Xml;HDR=NO;IMEX=1;'";
                            else
                            {
                                return;
                            }

                            idestilo = Convert.ToInt32(dt1.Rows[0][0].ToString());

                            using (OleDbCommand comm = new OleDbCommand())
                            {
                                comm.CommandText = "Select * from [" + cadena + "$A1:GU25]";
                                comm.Connection = conn;
                                dt = new DataTable();
                             
                                using (OleDbDataAdapter da = new OleDbDataAdapter())
                                {
                                    da.SelectCommand = comm;
                                    da.Fill(dt);

                                }
                                
                            }//fin using OleDbDataAdapter

                            if (dt.Columns.Count>=199)
                            {
                                using (OleDbCommand comm = new OleDbCommand())
                                {
                                    comm.CommandText = "Select * from [" + cadena + "$GV1:OM25]";
                                    comm.Connection = conn;
                                    dt2 = new DataTable();

                                    using (OleDbDataAdapter da = new OleDbDataAdapter())
                                    {
                                        da.SelectCommand = comm;
                                        da.Fill(dt2);
                                    }

                                }//fin using OleDbDataAdapter


                                if (dt2.Columns.Count>=199)
                                {
                                    using (OleDbCommand comm = new OleDbCommand())
                                    {
                                        comm.CommandText = "Select * from [" + cadena + "$ON1:WE25]";
                                        comm.Connection = conn;
                                        dt3 = new DataTable();

                                        using (OleDbDataAdapter da = new OleDbDataAdapter())
                                        {
                                            da.SelectCommand = comm;
                                            da.Fill(dt3);
                                        }

                                    }//fin using OleDbDataAdapter
                                }
                                
                            }
                          

                        }//fin de using OleDbCommand


                        //extraer tallas de formato
                        int columnas = 0;
                        for (int i = 3; i < dt.Columns.Count; i++)
                        {
                            if (!dt.Rows[3][i].ToString().Equals(""))
                            {
                                columnas = i;
                            }
                            else
                                break;
                        }

                        int columnas1 = -1;
                        for (int i = 0; i < dt2.Columns.Count; i++)
                        {
                            if (!dt2.Rows[3][i].ToString().Equals(""))
                            {
                                columnas1=i;
                            }
                            else
                                break;
                        }

                        int columnas2 = -1;
                        for (int i = 0; i < dt3.Columns.Count; i++)
                        {
                            if (!dt3.Rows[3][i].ToString().Equals(""))
                            {
                                columnas2=i;
                            }
                            else
                                break;
                        }

                        //extraer tallas de formato
                        for (int i = 3; i <= columnas; i++)
                        {
                            var size = dt.Rows[1][i].ToString().Replace(" ", "").Trim().ToUpper();

                            var idtalla = OrderDetailDA.SaveTallaParametroSalida(size.Replace("*", "X"));

                            var obj = new objtallas();
                            obj.idtalla = idtalla;
                            obj.talla = size;
                            obj.valor = 1;

                            lista.Add(obj);

                        }

                        for (int i = 0; i <= columnas1; i++)
                        {
                            var size = dt2.Rows[1][i].ToString().Replace(" ", "").Trim().ToUpper();

                            var idtalla = OrderDetailDA.SaveTallaParametroSalida(size.Replace("*", "X"));

                            var obj = new objtallas();
                            obj.idtalla = idtalla;
                            obj.talla = size;
                            obj.valor = 1;

                            lista.Add(obj);

                        }

                        for (int i = 0; i <= columnas2; i++)
                        {
                            var size = dt3.Rows[1][i].ToString().Replace(" ", "").Trim().ToUpper();

                            var idtalla = OrderDetailDA.SaveTallaParametroSalida(size.Replace("*", "X"));

                            var obj = new objtallas();
                            obj.idtalla = idtalla;
                            obj.talla = size;
                            obj.valor = 1;

                            lista.Add(obj);

                        }

                        string[] separators = { "," };


                        for (int i = 3; i < dt.Rows.Count; i++)
                        {

                            var puntomedida = dt.Rows[i][1].ToString().ToUpper().TrimEnd().TrimStart();

                            if (string.IsNullOrEmpty(puntomedida))
                            {
                                break;
                            }

                            var idpuntoM = OrderDetailDA.SavePuntoMedidaParametroSalida(puntomedida);
                            var tol = dt.Rows[i][2].ToString().Trim().Replace(" ", "");

                            var band = tol.IndexOf(',');

                            var tolMX = 0.0;
                            var tolMN = 0.0;

                            if (band > 1)
                            {
                                var arr = tol.Split(separators, StringSplitOptions.RemoveEmptyEntries);

                                tolMX = Convert.ToDouble(listaMedidas.FirstOrDefault(x => x.medida.Equals(arr[0].Replace(" ", ""))).valor); //dt1.Rows[0][2].ToString()
                                tolMN = Convert.ToDouble(listaMedidas.FirstOrDefault(x => x.medida.Equals(arr[1].Replace(" ", ""))).valor);

                            }
                            else
                            {
                                var valor1 = tol.Replace("+/-", "+");
                                var valor2 = tol.Replace("+/-", "-");

                                tolMX = Convert.ToDouble(listaMedidas.FirstOrDefault(x => x.medida.Equals(valor1.Replace(" ", ""))).valor); //dt1.Rows[0][2].ToString()
                                tolMN = Convert.ToDouble(listaMedidas.FirstOrDefault(x => x.medida.Equals(valor2.Replace(" ", ""))).valor);

                            }
                          

                            int contList = 0;// lista.Count();
                            for (int j = 3; j <= columnas; j++)
                            {
                                try
                                {
                                    var valor = dt.Rows[i][j].ToString().Trim() == "" ? "0" : dt.Rows[i][j].ToString().Trim();

                                    if (!valor.Equals("0"))
                                    {
                                        var valorReal = double.Parse(dt.Rows[i][j].ToString());

                                        var valorAddTolmax = valorReal + tolMX;
                                        var valorAddTolmn = valorReal - Math.Abs(tolMN);

                                        var obj3 = lista[contList];

                                        var objespc = new objespecificaciones
                                        {
                                            idPOM = idpuntoM,
                                            idtalla = obj3.idtalla,
                                            idestilo = idestilo,
                                            valorspec = valorReal,
                                            valorMax = valorAddTolmax,
                                            valorMin = valorAddTolmn,
                                            rango = obj3.valor,
                                            Estilo = cadena

                                        };

                                        listaFinal.Add(objespc);
                                    }

                                    contList++;
                                }
                                catch (Exception ex)
                                {
                                    try
                                    {
                                        var desarr = dt.Rows[i][j].ToString().Trim().Split(' ');

                                        var entero = double.Parse(desarr[0].ToString());

                                        var arrA = desarr[1].Trim().Split('/');
                                        var decim = double.Parse(arrA[0].ToString()) / double.Parse(arrA[1].ToString());

                                        var valorReal = Math.Round(entero + decim, 3);

                                        var valorAddTolmax = valorReal + tolMX;
                                        var valorAddTolmn = valorReal - Math.Abs(tolMN);

                                        var obj3 = lista[contList];

                                        var objespc = new objespecificaciones
                                        {
                                            idPOM = idpuntoM,
                                            idtalla = obj3.idtalla,
                                            idestilo = idestilo,
                                            valorspec = valorReal,
                                            valorMax = valorAddTolmax,
                                            valorMin = valorAddTolmn,
                                            rango = obj3.valor,
                                            Estilo = cadena

                                        };

                                        listaFinal.Add(objespc);

                                        contList++;

                                    }
                                    catch (Exception)
                                    {
                                        contList++;
                                        // throw;
                                    }
                                    
                                    var r = contList;
                                    var resppp = ex.Message;
                                    // throw;
                                }
                              

                             
                            }

                            for (int j = 0; j <= columnas1; j++)
                            {
                                try
                                {
                                    var valor = dt2.Rows[i][j].ToString().Trim() == "" ? "0" : dt2.Rows[i][j].ToString().Trim();

                                    if (!valor.Equals("0"))
                                    {
                                        var valorReal = double.Parse(dt2.Rows[i][j].ToString());

                                        var valorAddTolmax = valorReal + tolMX;
                                        var valorAddTolmn = valorReal - Math.Abs(tolMN);

                                        var obj3 = lista[contList];

                                        var objespc = new objespecificaciones
                                        {
                                            idPOM = idpuntoM,
                                            idtalla = obj3.idtalla,
                                            idestilo = idestilo,
                                            valorspec = valorReal,
                                            valorMax = valorAddTolmax,
                                            valorMin = valorAddTolmn,
                                            rango = obj3.valor,
                                            Estilo = cadena

                                        };

                                        listaFinal.Add(objespc);

                                        contList++;
                                    }
                                    else
                                        contList++;
                                }
                                catch (Exception ex)
                                {
                                    try
                                    {
                                        var desarr = dt2.Rows[i][j].ToString().Trim().Split(' ');

                                        var entero = double.Parse(desarr[0].ToString());

                                        var arrA = desarr[1].Trim().Split('/');
                                        var decim = double.Parse(arrA[0].ToString()) / double.Parse(arrA[1].ToString());

                                        var valorReal = Math.Round(entero + decim, 3);

                                        var valorAddTolmax = valorReal + tolMX;
                                        var valorAddTolmn = valorReal - Math.Abs(tolMN);

                                        var obj3 = lista[contList];

                                        var objespc = new objespecificaciones
                                        {
                                            idPOM = idpuntoM,
                                            idtalla = obj3.idtalla,
                                            idestilo = idestilo,
                                            valorspec = valorReal,
                                            valorMax = valorAddTolmax,
                                            valorMin = valorAddTolmn,
                                            rango = obj3.valor,
                                            Estilo = cadena

                                        };

                                        listaFinal.Add(objespc);

                                        contList++;

                                    }
                                    catch (Exception)
                                    {
                                        contList++;
                                        // throw;
                                    }

                                    var r = contList;
                                    var resppp = ex.Message;
                                }
                            }

                            for (int j = 0; j <= columnas2; j++)
                            {
                                try
                                {
                                    var valor = dt3.Rows[i][j].ToString().Trim() == "" ? "0" : dt3.Rows[i][j].ToString().Trim();

                                    if (!valor.Equals("0"))
                                    {
                                        var valorReal = double.Parse(dt3.Rows[i][j].ToString());

                                        var valorAddTolmax = valorReal + tolMX;
                                        var valorAddTolmn = valorReal - Math.Abs(tolMN);

                                        var obj3 = lista[contList];

                                        var objespc = new objespecificaciones
                                        {
                                            idPOM = idpuntoM,
                                            idtalla = obj3.idtalla,
                                            idestilo = idestilo,
                                            valorspec = valorReal,
                                            valorMax = valorAddTolmax,
                                            valorMin = valorAddTolmn,
                                            rango = obj3.valor,
                                            Estilo = cadena

                                        };

                                        listaFinal.Add(objespc);

                                        contList++;
                                    }
                                    else
                                        contList++;
                                }
                                catch (Exception ex)
                                {
                                    try
                                    {
                                        var desarr = dt3.Rows[i][j].ToString().Trim().Split(' ');

                                        var entero = double.Parse(desarr[0].ToString());

                                        var arrA = desarr[1].Trim().Split('/');
                                        var decim = double.Parse(arrA[0].ToString()) / double.Parse(arrA[1].ToString());

                                        var valorReal = Math.Round(entero + decim, 3);

                                        var valorAddTolmax = valorReal + tolMX;
                                        var valorAddTolmn = valorReal - Math.Abs(tolMN);

                                        var obj3 = lista[contList];

                                        var objespc = new objespecificaciones
                                        {
                                            idPOM = idpuntoM,
                                            idtalla = obj3.idtalla,
                                            idestilo = idestilo,
                                            valorspec = valorReal,
                                            valorMax = valorAddTolmax,
                                            valorMin = valorAddTolmn,
                                            rango = obj3.valor,
                                            Estilo = cadena

                                        };

                                        listaFinal.Add(objespc);

                                        contList++;

                                    }
                                    catch (Exception)
                                    {
                                        contList++;
                                        // throw;
                                    }

                                    var r = contList;
                                    var resppp = ex.Message;
                                    // throw;
                                }
                            }

                        }//fin ciclo for

                        worker.ReportProgress(Convert.ToInt16(progress) + 1000);

                        int contador = 0;
                        int contprogress = 0;

                        //  Thread.Sleep(2000);

                        foreach (var item2 in listaFinal)
                        {

                            contprogress++;
                            try
                            {
                                //  Thread.Sleep(5);
                                var resp = OrderDetailDA.saveEspecificacionNew(item2.idestilo, item2.idtalla, item2.idPOM, item2.valorspec, item2.valorMax, item2.valorMin, "Desktop", item2.rango);

                                if (resp != "OK")
                                {

                                    contador++;
                                }
                            }
                            catch (Exception)
                            {

                                contador++;

                            }

                            var progress2 = Convert.ToDouble(contprogress) / listaFinal.Count * 100;
                            worker.ReportProgress(Convert.ToInt16(progress2));

                        }

                        if (contador > 0)
                        {
                            errores = true;
                            TEXTO.AppendLine("el estilo " + item + " Numero de errores " + contador + " no existe \n");
                            TEXTO.AppendLine("--------------------------------------------");
                        }

                        #endregion
                    }
                    catch (Exception ex)
                    {

                        errores = true;
                        TEXTO.AppendLine("Error de lectura estilo " + item + " Verifique la informacion del archivo formato, tallas, tolerencia, punto de medida  \n");
                        TEXTO.AppendLine("--------------------DETALLE DL ERROR------------------------");
                        TEXTO.AppendLine(ex.Message);
                        TEXTO.AppendLine("--------------------------------------------");
                    }
                   
                }//fin de using OleDbConnection 


            }


            if (errores)
            {
                TEXTO.AppendLine("--------------------------------------------");
                var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                var PATH = System.IO.Path.Combine(desktopPath, "XLSX_CON_PROBLEMAS_DE_LECTURA_" + DateTime.Now.ToShortDateString().Replace("/", "-") + "_" + DateTime.Now.ToShortTimeString().Replace(":", "-") + ".TXT");
                System.IO.File.WriteAllText(PATH, TEXTO.ToString());

            }

            var listas = new List<object>();

        
            listas.Add(errores);
            

            e.Result = listas;
        }


        void _worker_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {

            if (e.ProgressPercentage>1000)
            {
                var percentaje = e.ProgressPercentage - 1000;
                lblprogress.Content = "Lectura de Archivo!!!";
                progressBar.Value = percentaje;
            }
            else
            { 
              lblprogress1.Content = "Ingresando Registros!!!";
              progressBar1.Value = e.ProgressPercentage;
            }
        }


        void _worker_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            List<object> result;
            result = e.Result as List<object>;

            var resp = (bool)result[0];


            progressBar.Value = 0;
            progressBar1.Value = 0;

            lblprogress.Content = "Cargado completamente!!!";
            lblprogress1.Content = "Cargado completamente!!!";

            if (resp)
            {
                System.Windows.MessageBox.Show("Se ha creado un archivo de texto en su escritorio que muestra los XLXS con problemas de lectura");
            }
        }

    }
}
