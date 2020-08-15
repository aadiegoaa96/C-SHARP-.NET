using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
// Communicate via Serial
using System.IO;
//using System.IO.StreamReader;
using System.IO.Ports;
using System.Xml;
// Add ZedGraph
using ZedGraph;
using System.Timers;
using Label = System.Windows.Forms.Label;
using Microsoft.Office.Interop.Excel;
using Spire.Xls;
using System.Runtime.InteropServices;


namespace ExcelPrueba1_11_10_2019
{
    public partial class Valvula_de_Venteo_de_Baja : Form
    {
        System.Timers.Timer timer;
        System.Timers.Timer timer2;
        DateTime iDate2;
        DateTime iDate3;
        string Estado_de_la_valvula = String.Empty;    // Esta variable se utiliza para saber el estado de la válvula
        string SDatas2 = String.Empty;                 // Aquí se guardarán los valores de presión
        string SDatas = String.Empty;                  // Declare string to save sensor data sent via Serial   Aquí se guardarán los valores de temperatura 
        string SRealTime = String.Empty;               // Declare string to save time sent via Serial    
        string SRealTime2 = String.Empty;              // Declare string to save time sent via Serial    
        int status = 4;                                // Declare variables to handle graphing events  
        double realtime = 0;                           // Declare time variables to graph  
        double realtime2 = 0;                          // Declare time variables to graph  
        double datas = 0;                              // Declare sensor data to draw graphs
        double datas2 = 0;                             // Declare sensor data to draw graphs
        double Edo_Valvula = 0;                        // Aquí se guaradará lo que contenga el primer byte para saber el estado de la válvula
        int valv;                                      // Aquí se guardará el tipo de válvula para realizar su reporte correspondiente.
        String rutaArchivo = string.Empty;
        String rutaArchivo2 = string.Empty;
        //String rutaArchivo3 = string.Empty; // hoja membretada
        GraphPane Grafica1 = new GraphPane();
        GraphPane Grafica2 = new GraphPane();
        readonly PointPairList Punto1 = new PointPairList();
        readonly PointPairList Punto2 = new PointPairList();
        //private readonly int currentTime;
        LineItem Curva1;
        LineItem Curva2;
        double Valor1 = 0;
        double maxi1 = 0;    // °C
        double mini1 = 250;
        double maxi2 = 0;    //psi
        double mini2 = 5000;
        String ruta = string.Empty; //hoja membretada
        string path = String.Empty;

        public System.Timers.Timer Timer { get => timer; set => timer = value; }
        public System.Timers.Timer Timer2 { get => timer2; set => timer2 = value; }
        public DateTime IDate2 { get => iDate2; set => iDate2 = value; }
        public DateTime IDate3 { get => iDate3; set => iDate3 = value; }
        public Valvula_de_Venteo_de_Baja()
        {
            InitializeComponent();
            IDate3 = dateTimePicker1.Value; //Se guarda en la variable IDate3 la fecha de inicio de la prueba
            bt_Connect.Enabled = false;
            bt_Save.Enabled = false;
            btSalvarFoto.Enabled = false;
            groupBox4.Visible = true;
            bt_Cargar.Enabled = false;
        }

        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void RealeaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hwnd, int wmsg, int wparam, int lparam);


        private void Valvula_de_Venteo_de_Baja_Load(object sender, EventArgs e)
        {
            Timer = new System.Timers.Timer();
            //Timer.Interval = 2400000; //para cada 40 minutos
            Timer.Interval = 1000;
            Timer.Elapsed += Timer_Elapsed;
            Timer2 = new System.Timers.Timer();
            //Timer.Interval = 2400000; //para cada 40 minutos
            Timer2.Interval = 1000;
            Timer2.Elapsed += Timer_Elapsed2;
            comboBox1.DataSource = SerialPort.GetPortNames(); // Get the source for comboBox is the name of the COM port   
            comboBox1.Text = Properties.Settings.Default.ComName; // Get ComName did in step 5 for comboBox                                                      
            Control.CheckForIllegalCrossThreadCalls = false;
            status = 4;
            status = 0;   // status = 0;  Si no se pone el status la prueba nunca se para
        }


        private void Timer_Elapsed2(object sender, System.Timers.ElapsedEventArgs e)
        {
            status = 0;
            Data_Listview();
            DateTime currentTime = DateTime.Now;
            DateTime userTime = dateTimePicker1.Value;
            //      if (currentTime.Hour == userTime.Hour && currentTime.Minute == userTime.Minute && currentTime.Second == userTime.Second  && IDate2 == null)
            if (currentTime.Hour == userTime.Hour && currentTime.Minute == userTime.Minute && currentTime.Second == userTime.Second)
            {
                IDate2 = dateTimePicker1.Value; //Se guarda en la variable IDate2 la fecha de fin de la prueba
                                                //  serialPort1.Close();        // CERRAMOS EL PUERTO SERIAL
                Timer.Stop();
                Timer2.Stop();
                bt_Connect.Enabled = false;
                bt_Save.Enabled = false;
                btSalvarFoto.Enabled = true;
                groupBox4.Visible = true;
                bt_Cargar.Enabled = false;
                MessageBox.Show("PRUEBA FINALIZADA 2, VÁLVULA EN ÓPTIMAS CONDICIONES. ", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                serialPort1.Close();        // CERRAMOS EL PUERTO SERIAL
            }
        }
        private void Timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            DateTime currentTime2 = DateTime.Now;
            DateTime userTime2 = dateTimePicker1.Value;
        }


        private void Timer1_Tick(object sender, EventArgs e)
        {
            if (!serialPort1.IsOpen)
            {
                progressBar1.Value = 0;
            }

            else if (serialPort1.IsOpen)
            {
                progressBar1.Value = 100;
                status = 0;
            }
        }


        // This function stores the selected COM port for the connection
        private void SaveSetting()
        {
            Properties.Settings.Default.ComName = comboBox1.Text;
            Properties.Settings.Default.Save();
        }


        private void Graficar()
        {
            Grafica1 = zedGraphControl1.GraphPane;
            Grafica2 = zedGraphControl2.GraphPane;
            Grafica1.Title.Text = "GRÁFICA EN TIEMPO REAL";
            Grafica2.Title.Text = "GRÁFICA EN TIEMPO REAL";
            Grafica1.XAxis.Title.Text = " t(s) ";
            Grafica2.XAxis.Title.Text = " t(s) ";
            Grafica1.YAxis.Title.Text = "PRESIÓN (PSI)";
            Grafica2.YAxis.Title.Text = "TEMPERTATURA (°C)";
            Grafica1.YAxis.Scale.Min = 0;
            Grafica2.YAxis.Scale.Min = 0;
            Grafica1.YAxis.Scale.Max = 1024;
            Grafica2.YAxis.Scale.Max = 1024;
            Curva1 = Grafica1.AddCurve(null, Punto1, Color.Red, SymbolType.None);
            Curva2 = Grafica2.AddCurve(null, Punto2, Color.Blue, SymbolType.None);
            Curva1.Line.Width = 1;
            Curva2.Line.Width = 1;
        }


        private void SerialPort1_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            status = 1; // Get the string processed event, change starus to 1 to display data in ListView and graph 

            try
            {
                string[] arrList = serialPort1.ReadLine().Split('|'); // Read a line of Serial, cut the string when encountered hyphen characters 
                //  string[] arrList = serialPort1.ReadLine().Split('*'); // Read a line of Serial, cut the string when encountered hyphen characters 
                Estado_de_la_valvula = arrList[0]; // The first string saved in SRealTime  
                double.TryParse(Estado_de_la_valvula, out Edo_Valvula); // Convert to double  
                if (Edo_Valvula == 666)  // Parar la prueba cuando la válvula esté dañada recibiendo el caracter 666 y detectandolo aqui
                {
                    serialPort1.Write("0"); // Send character "0" via Serial, Stop Arduino 
                    bt_Connect.Enabled = false;
                    bt_Save.Enabled = false;
                    btSalvarFoto.Enabled = true;
                    groupBox4.Visible = true;
                    status = 4;
                    status = 0;

                    if (serialPort1.IsOpen)
                    {
                        Timer.Stop();
                        Timer2.Stop(); IDate2 = dateTimePicker1.Value; //Se guarda en la variable IDate2 la fecha de fin de la prueba
                        MessageBox.Show("PRUEBA FINALIZADA, VÁLVULA DAÑADA. ", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        bt_Connect.Enabled = false;
                        bt_Save.Enabled = false;
                        btSalvarFoto.Enabled = true;
                        serialPort1.Close();
                    }
                }



                SRealTime = arrList[0]; // The first string saved in SRealTime  
                SDatas = arrList[1]; // The second string saved to SDatas  

                double.TryParse(SDatas, out datas); // Convert to double  
                double.TryParse(SRealTime, out realtime);
                if (datas > maxi1)                                   // °C
                {
                    maxi1 = datas;
                }
                if (datas < mini1)
                {
                    mini1 = datas;
                }
                SRealTime2 = arrList[2]; // The first string saved in SRealTime  
                SDatas2 = arrList[3]; // The second string saved to SDatas  
                double.TryParse(SDatas2, out datas2); // Convert to double  
                double.TryParse(SRealTime2, out realtime2);
                //realtime2 = realtime2 / 1000.0; // For ms to s  
                if (datas2 > maxi2)                                  //PSI
                {
                    maxi2 = datas2;
                }
                if (datas2 < mini2)
                {
                    mini2 = datas2;
                }
                Valor1 += 0.05;
                Punto1.Add(new PointPair(Valor1, Convert.ToDouble(arrList[3].ToString())));
                Punto2.Add(new PointPair(Valor1, Convert.ToDouble(arrList[1].ToString())));
                Grafica2.XAxis.Scale.Max = Valor1;
                Grafica1.XAxis.Scale.Max = Valor1;
                Grafica1.AxisChange();
                Grafica2.AxisChange();
                zedGraphControl1.Refresh();
                zedGraphControl2.Refresh();
            }
            catch
            {
                return;
            }
        }

        // Display data in ListView
        private void Data_Listview()
        {
            if (status == 4)
                return;
            else
            {
                ListViewItem item = new ListViewItem(realtime.ToString()); // Assign the realtime variable to the first column of ListView   
                item.SubItems.Add(datas.ToString());
                listView3.Items.Add(item); // Assign datas variable to the next column of ListView 
                                           // Do not assign SDatas string because when exporting data to Excel as a string, it cannot perform calculations
                listView3.Items[listView3.Items.Count - 1].EnsureVisible(); // Display the most recently assigned line in ListView, that is, I scroll ListView according to the latest data   
                ListViewItem item2 = new ListViewItem(realtime2.ToString()); // Assign the realtime variable to the first column of ListView   
                item2.SubItems.Add(datas2.ToString());
                listView4.Items.Add(item2); // Assign datas variable to the next column of ListView 
                                            // Do not assign SDatas string because when exporting data to Excel as a string, it cannot perform calculations
                listView4.Items[listView4.Items.Count - 1].EnsureVisible(); // Display the most recently assigned line in ListView, that is, I scroll ListView according to the latest data   
            }
            if (status == 0)
                return;
            else
            {
                ListViewItem item = new ListViewItem(realtime.ToString()); // Assign the realtime variable to the first column of ListView   
                item.SubItems.Add(datas.ToString());
                listView1.Items.Add(item); // Assign datas variable to the next column of ListView 
                                           // Do not assign SDatas string because when exporting data to Excel as a string, it cannot perform calculations
                listView1.Items[listView1.Items.Count - 1].EnsureVisible(); // Display the most recently assigned line in ListView, that is, I scroll ListView according to the latest data   
                ListViewItem item2 = new ListViewItem(realtime2.ToString()); // Assign the realtime variable to the first column of ListView   
                item2.SubItems.Add(datas2.ToString());
                listView2.Items.Add(item2); // Assign datas variable to the next column of ListView 
                                            // Do not assign SDatas string because when exporting data to Excel as a string, it cannot perform calculations
                listView2.Items[listView2.Items.Count - 1].EnsureVisible(); // Display the most recently assigned line in ListView, that is, I scroll ListView according to the latest data   
            }
        }


        // Delete the graph, with ZedGraph must be declared again as in Form1_Load, otherwise will not display
        private void ClearZedGraph()
        {
            zedGraphControl1.GraphPane.CurveList.Clear(); // Delete a line 
            zedGraphControl1.GraphPane.GraphObjList.Clear(); // Delete object 
            zedGraphControl1.AxisChange();
            zedGraphControl1.Invalidate();
            GraphPane myPane = zedGraphControl1.GraphPane;
            myPane.Title.Text = "Graph data over time";
            myPane.XAxis.Title.Text = "Time (s)";
            myPane.YAxis.Title.Text = "Data";
            RollingPointPairList list = new RollingPointPairList(60000);
            LineItem curve = myPane.AddCurve("Data", list, Color.Red, SymbolType.None);
            myPane.XAxis.Scale.Min = 0;
            myPane.XAxis.Scale.Max = 30;
            myPane.XAxis.Scale.MinorStep = 1;
            myPane.XAxis.Scale.MajorStep = 5;
            myPane.YAxis.Scale.Min = -100;
            myPane.YAxis.Scale.Max = 100;
            zedGraphControl1.AxisChange();
        }




        // Delete the graph, with ZedGraph must be declared again as in Form1_Load, otherwise will not display
        private void ClearZedGraph2()
        {
            zedGraphControl2.GraphPane.CurveList.Clear(); // Delete a line 
            zedGraphControl2.GraphPane.GraphObjList.Clear(); // Delete object 
            zedGraphControl2.AxisChange();
            zedGraphControl2.Invalidate();
            GraphPane myPane = zedGraphControl2.GraphPane;
            myPane.Title.Text = "Graph data over time";
            myPane.XAxis.Title.Text = "Time (s)";
            myPane.YAxis.Title.Text = "Data";
            RollingPointPairList list = new RollingPointPairList(60000);
            LineItem curve = myPane.AddCurve("Data", list, Color.Red, SymbolType.None);
            myPane.XAxis.Scale.Min = 0;
            myPane.XAxis.Scale.Max = 30;
            myPane.XAxis.Scale.MinorStep = 1;
            myPane.XAxis.Scale.MajorStep = 5;
            myPane.YAxis.Scale.Min = -100;
            myPane.YAxis.Scale.Max = 100;
            zedGraphControl2.AxisChange();
        }

        // Function to delete data
        private void ResetValue()
        {
            valv = 0;
            realtime2 = 0;
            realtime = 0;
            datas = 0;
            datas2 = 0;
            SDatas = String.Empty;
            SDatas2 = String.Empty;
            SRealTime = String.Empty;
            SRealTime2 = String.Empty;
            status = 4; // Change status to 0  
            maxi1 = 0;
            maxi2 = 0;
            mini1 = 0;
            mini2 = 0;
            timer2 = null;
            dateTimePicker1 = null;
            Edo_Valvula = 0;
        }



        // Function to save ListView to Excel
        private void Atemperadora_Sobrecalentado()
        {
     

            Microsoft.Office.Interop.Excel.Application xla = new Microsoft.Office.Interop.Excel.Application();
            xla.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook wb = xla.Workbooks.Add(Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);
            Spire.Xls.Workbook book = new Spire.Xls.Workbook();
            Spire.Xls.Worksheet sheet = book.Worksheets[0];
            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws2 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws3 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws4 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws5 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws6 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws7 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws8 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws9 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws10 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws11 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws12 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws13 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws14 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws15 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws16 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws17 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws18 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws19 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws20 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws21 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws22 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            // Name the two cells A1. B1 are "Time (s)" and "Data" respectively, then automatically expand
            Microsoft.Office.Interop.Excel.Range rg18 = (Microsoft.Office.Interop.Excel.Range)ws18.get_Range("A1", "B1");
            Microsoft.Office.Interop.Excel.Range rg = (Microsoft.Office.Interop.Excel.Range)ws.get_Range("A10", "B10"); // EN LA COORDENADA A10,B10 ESTARÁN LOS VALORES DE TEMPERATUR
            Microsoft.Office.Interop.Excel.Range rg2 = (Microsoft.Office.Interop.Excel.Range)ws2.get_Range("C10", "D10");
            Microsoft.Office.Interop.Excel.Range rg3 = (Microsoft.Office.Interop.Excel.Range)ws3.get_Range("A38", "B38");
            Microsoft.Office.Interop.Excel.Range rg4 = (Microsoft.Office.Interop.Excel.Range)ws4.get_Range("A40", "B40");
            Microsoft.Office.Interop.Excel.Range rg19 = (Microsoft.Office.Interop.Excel.Range)ws19.get_Range("A53", "B53");
            Microsoft.Office.Interop.Excel.Range rg5 = (Microsoft.Office.Interop.Excel.Range)ws5.get_Range("A105", "B105");
            Microsoft.Office.Interop.Excel.Range rg6 = (Microsoft.Office.Interop.Excel.Range)ws6.get_Range("A106", "B106");
            Microsoft.Office.Interop.Excel.Range rg7 = (Microsoft.Office.Interop.Excel.Range)ws7.get_Range("A107", "B107");
            Microsoft.Office.Interop.Excel.Range rg8 = (Microsoft.Office.Interop.Excel.Range)ws8.get_Range("A108", "B108");
            Microsoft.Office.Interop.Excel.Range rg9 = (Microsoft.Office.Interop.Excel.Range)ws9.get_Range("A109", "B109");
            Microsoft.Office.Interop.Excel.Range rg10 = (Microsoft.Office.Interop.Excel.Range)ws10.get_Range("A110", "B110");
            Microsoft.Office.Interop.Excel.Range rg11 = (Microsoft.Office.Interop.Excel.Range)ws11.get_Range("A111", "B111");
            Microsoft.Office.Interop.Excel.Range rg12 = (Microsoft.Office.Interop.Excel.Range)ws12.get_Range("A112", "B112");
            Microsoft.Office.Interop.Excel.Range rg13 = (Microsoft.Office.Interop.Excel.Range)ws13.get_Range("A113", "B113");
            Microsoft.Office.Interop.Excel.Range rg14 = (Microsoft.Office.Interop.Excel.Range)ws14.get_Range("A114", "B114");
            Microsoft.Office.Interop.Excel.Range rg15 = (Microsoft.Office.Interop.Excel.Range)ws15.get_Range("A118", "B118");
            Microsoft.Office.Interop.Excel.Range rg16 = (Microsoft.Office.Interop.Excel.Range)ws16.get_Range("A122", "B122");
            Microsoft.Office.Interop.Excel.Range rg17 = (Microsoft.Office.Interop.Excel.Range)ws17.get_Range("A117", "B117");
            Microsoft.Office.Interop.Excel.Range rg20 = (Microsoft.Office.Interop.Excel.Range)ws20.get_Range("A118", "B118");
            Microsoft.Office.Interop.Excel.Range rg21 = (Microsoft.Office.Interop.Excel.Range)ws21.get_Range("A78", "B78");
            Microsoft.Office.Interop.Excel.Range rg22 = (Microsoft.Office.Interop.Excel.Range)ws22.get_Range("A78", "B78");
            ws3 = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.get_Item(1);
            ws4 = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.get_Item(1);
            ws5 = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.get_Item(1);
            ws6 = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.get_Item(1);
            //imagen reporte
            ws3.Shapes.AddPicture(rutaArchivo, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 20, 870, 400, 150);   // 710 mide los pixeles de una hoja. El primer parámetro posición X, segundo posicón Y, tercero tamaño en Y, cuarto tamaño en X.
            ws4.Shapes.AddPicture(rutaArchivo2, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 20, 1050, 400, 150); // 710 mide los pixeles de una hoja. El primer parámetro posición X, segundo posicón Y, tercero tamaño en Y, cuarto tamaño en X.
            ws5.Cells[105, 1] = "Fecha de inicio: " + IDate3;
            ws6.Cells[106, 1] = "Fecha límite elegida: " + IDate2;
            //imagen hoja membretada
            ws7.Shapes.AddPicture(path + "//Sin título.Jpeg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, 0, 400, 120);    // 710 mide los pixeles de una hoja. El primer parámetro posición X, segundo posicón Y, tercero tamaño en Y, cuarto tamaño en X.
            ws8.Shapes.AddPicture(path + "//Sin título2.Jpeg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, 585, 400, 120); // 710 mide los pixeles de una hoja. El primer parámetro posición X, segundo posicón Y, tercero tamaño en Y, cuarto tamaño en X.
            ws19.Shapes.AddPicture(path + "//Sin título.Jpeg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, 710, 400, 120);   // 710 mide los pixeles de una hoja. El primer parámetro posición X, segundo posicón Y, tercero tamaño en Y, cuarto tamaño en X.
            ws20.Shapes.AddPicture(path + "//Sin título2.Jpeg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, 1285, 400, 120); // 710 mide los pixeles de una hoja. El primer parámetro posición X, segundo posicón Y, tercero tamaño en Y, cuarto tamaño en X.
            ws21.Shapes.AddPicture(path + "//Sin título.Jpeg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, 1420, 400, 120);   // 710 mide los pixeles de una hoja. El primer parámetro posición X, segundo posicón Y, tercero tamaño en Y, cuarto tamaño en X.
            ws22.Shapes.AddPicture(path + "//Sin título2.Jpeg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, 1995, 400, 120); // 710 mide los pixeles de una hoja. El primer parámetro posición X, segundo posicón Y, tercero tamaño en Y, cuarto tamaño en X.
            ws.Cells[10, 1] = "Tiempo (segundos): ";
            ws.Cells[10, 2] = "Temperatura (°C): ";
            int i = 11; //tiempo1
            int j = 12; //°C
            foreach (ListViewItem comp in listView1.Items)
            {
                ws.Cells[i, j] = comp.Text.ToString();
                foreach (ListViewItem.ListViewSubItem drv in comp.SubItems)
                {
                    ws.Cells[i, j] = drv.Text.ToString();
                    j++;
                }
                j = 1;
                i++;
            }

            rg.Columns.AutoFit();
            ws2.Cells[10, 3] = "Tiempo (segundos): ";
            ws2.Cells[10, 4] = "Presión (PSI): ";
            int w = 11; //tiempo2
            int q = 12; //PSI
            foreach (ListViewItem comp in listView2.Items) //de esta manera se sale solo una casilla, es la de presión, y solo no concuerda esa de presión que se sale.
            {
                ws.Cells[w, q] = comp.Text.ToString();
                // ws.Protect("def-345", SheetProtectionType.All);
                foreach (ListViewItem.ListViewSubItem drv in comp.SubItems)
                {
                    ws.Cells[w, q] = drv.Text.ToString();
                    //numero = ws.Cells[w, q];
                    q++;
                }
                q = 3;
                w++;
            }
            rg2.Columns.AutoFit();


            //mayor se debe envia como resultado
            ws7.Cells[107, 1] = "Temperatura max. : " + maxi1;
            ws8.Cells[108, 1] = "Temperatura min. : " + mini1;
            ws9.Cells[109, 1] = "Presión max. : " + maxi2;
            ws10.Cells[110, 1] = "Presión min. : " + mini2;
            ws11.Cells[111, 1] = "Tipo de válvula: Válvula Atemperadora de Vapor de Sobrecalentado";   //realtime , realtime2
            ws12.Cells[112, 1] = "Observaciones: _____________________________________________________________________";   //realtime , realtime2
            ws13.Cells[113, 1] = "____________________________________________________________________________________";   //realtime , realtime2
            ws14.Cells[114, 1] = "____________________________________________________________________________________";   //realtime , realtime2
            ws15.Cells[118, 1] = "Elaboró: ______________________________________________________";   //realtime , realtime2
            ws16.Cells[122, 1] = "Revisó:  ______________________________________________________";   //realtime , realtime2
            ws.Protect("def-345", SheetProtectionType.All);


        }









        private void Purga_Continua()
        {

            Microsoft.Office.Interop.Excel.Application xla = new Microsoft.Office.Interop.Excel.Application();
            xla.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook wb = xla.Workbooks.Add(Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);
            Spire.Xls.Workbook book = new Spire.Xls.Workbook();
            Spire.Xls.Worksheet sheet = book.Worksheets[0];
            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws2 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws3 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws4 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws5 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws6 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws7 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws8 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws9 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws10 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws11 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws12 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws13 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws14 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws15 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws16 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws17 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws18 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws19 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws20 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws21 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws22 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            // Name the two cells A1. B1 are "Time (s)" and "Data" respectively, then automatically expand
            Microsoft.Office.Interop.Excel.Range rg18 = (Microsoft.Office.Interop.Excel.Range)ws18.get_Range("A1", "B1");
            Microsoft.Office.Interop.Excel.Range rg = (Microsoft.Office.Interop.Excel.Range)ws.get_Range("A10", "B10"); // EN LA COORDENADA A10,B10 ESTARÁN LOS VALORES DE TEMPERATUR
            Microsoft.Office.Interop.Excel.Range rg2 = (Microsoft.Office.Interop.Excel.Range)ws2.get_Range("C10", "D10");
            Microsoft.Office.Interop.Excel.Range rg3 = (Microsoft.Office.Interop.Excel.Range)ws3.get_Range("A38", "B38");
            Microsoft.Office.Interop.Excel.Range rg4 = (Microsoft.Office.Interop.Excel.Range)ws4.get_Range("A40", "B40");
            Microsoft.Office.Interop.Excel.Range rg19 = (Microsoft.Office.Interop.Excel.Range)ws19.get_Range("A53", "B53");
            Microsoft.Office.Interop.Excel.Range rg5 = (Microsoft.Office.Interop.Excel.Range)ws5.get_Range("A105", "B105");
            Microsoft.Office.Interop.Excel.Range rg6 = (Microsoft.Office.Interop.Excel.Range)ws6.get_Range("A106", "B106");
            Microsoft.Office.Interop.Excel.Range rg7 = (Microsoft.Office.Interop.Excel.Range)ws7.get_Range("A107", "B107");
            Microsoft.Office.Interop.Excel.Range rg8 = (Microsoft.Office.Interop.Excel.Range)ws8.get_Range("A108", "B108");
            Microsoft.Office.Interop.Excel.Range rg9 = (Microsoft.Office.Interop.Excel.Range)ws9.get_Range("A109", "B109");
            Microsoft.Office.Interop.Excel.Range rg10 = (Microsoft.Office.Interop.Excel.Range)ws10.get_Range("A110", "B110");
            Microsoft.Office.Interop.Excel.Range rg11 = (Microsoft.Office.Interop.Excel.Range)ws11.get_Range("A111", "B111");
            Microsoft.Office.Interop.Excel.Range rg12 = (Microsoft.Office.Interop.Excel.Range)ws12.get_Range("A112", "B112");
            Microsoft.Office.Interop.Excel.Range rg13 = (Microsoft.Office.Interop.Excel.Range)ws13.get_Range("A113", "B113");
            Microsoft.Office.Interop.Excel.Range rg14 = (Microsoft.Office.Interop.Excel.Range)ws14.get_Range("A114", "B114");
            Microsoft.Office.Interop.Excel.Range rg15 = (Microsoft.Office.Interop.Excel.Range)ws15.get_Range("A118", "B118");
            Microsoft.Office.Interop.Excel.Range rg16 = (Microsoft.Office.Interop.Excel.Range)ws16.get_Range("A122", "B122");
            Microsoft.Office.Interop.Excel.Range rg17 = (Microsoft.Office.Interop.Excel.Range)ws17.get_Range("A117", "B117");
            Microsoft.Office.Interop.Excel.Range rg20 = (Microsoft.Office.Interop.Excel.Range)ws20.get_Range("A118", "B118");
            Microsoft.Office.Interop.Excel.Range rg21 = (Microsoft.Office.Interop.Excel.Range)ws21.get_Range("A78", "B78");
            Microsoft.Office.Interop.Excel.Range rg22 = (Microsoft.Office.Interop.Excel.Range)ws22.get_Range("A78", "B78");
            ws3 = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.get_Item(1);
            ws4 = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.get_Item(1);
            ws5 = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.get_Item(1);
            ws6 = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.get_Item(1);
            //imagen reporte
            ws3.Shapes.AddPicture(rutaArchivo, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 20, 870, 400, 150);   // 710 mide los pixeles de una hoja. El primer parámetro posición X, segundo posicón Y, tercero tamaño en Y, cuarto tamaño en X.
            ws4.Shapes.AddPicture(rutaArchivo2, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 20, 1050, 400, 150); // 710 mide los pixeles de una hoja. El primer parámetro posición X, segundo posicón Y, tercero tamaño en Y, cuarto tamaño en X.
            ws5.Cells[105, 1] = "Fecha de inicio: " + IDate3;
            ws6.Cells[106, 1] = "Fecha límite elegida: " + IDate2;
            //imagen hoja membretada
            ws7.Shapes.AddPicture(path + "//Sin título.Jpeg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, 0, 400, 120);    // 710 mide los pixeles de una hoja. El primer parámetro posición X, segundo posicón Y, tercero tamaño en Y, cuarto tamaño en X.
            ws8.Shapes.AddPicture(path + "//Sin título2.Jpeg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, 585, 400, 120); // 710 mide los pixeles de una hoja. El primer parámetro posición X, segundo posicón Y, tercero tamaño en Y, cuarto tamaño en X.
            ws19.Shapes.AddPicture(path + "//Sin título.Jpeg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, 710, 400, 120);   // 710 mide los pixeles de una hoja. El primer parámetro posición X, segundo posicón Y, tercero tamaño en Y, cuarto tamaño en X.
            ws20.Shapes.AddPicture(path + "//Sin título2.Jpeg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, 1285, 400, 120); // 710 mide los pixeles de una hoja. El primer parámetro posición X, segundo posicón Y, tercero tamaño en Y, cuarto tamaño en X.
            ws21.Shapes.AddPicture(path + "//Sin título.Jpeg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, 1420, 400, 120);   // 710 mide los pixeles de una hoja. El primer parámetro posición X, segundo posicón Y, tercero tamaño en Y, cuarto tamaño en X.
            ws22.Shapes.AddPicture(path + "//Sin título2.Jpeg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, 1995, 400, 120); // 710 mide los pixeles de una hoja. El primer parámetro posición X, segundo posicón Y, tercero tamaño en Y, cuarto tamaño en X.
            ws.Cells[10, 1] = "Tiempo (segundos): ";
            ws.Cells[10, 2] = "Temperatura (°C): ";
            int i = 11; //tiempo1
            int j = 12; //°C
            foreach (ListViewItem comp in listView1.Items)
            {
                ws.Cells[i, j] = comp.Text.ToString();
                foreach (ListViewItem.ListViewSubItem drv in comp.SubItems)
                {
                    ws.Cells[i, j] = drv.Text.ToString();
                    j++;
                }
                j = 1;
                i++;
            }

            rg.Columns.AutoFit();
            ws2.Cells[10, 3] = "Tiempo (segundos): ";
            ws2.Cells[10, 4] = "Presión (PSI): ";
            int w = 11; //tiempo2
            int q = 12; //PSI
            foreach (ListViewItem comp in listView2.Items) //de esta manera se sale solo una casilla, es la de presión, y solo no concuerda esa de presión que se sale.
            {
                ws.Cells[w, q] = comp.Text.ToString();
                // ws.Protect("def-345", SheetProtectionType.All);
                foreach (ListViewItem.ListViewSubItem drv in comp.SubItems)
                {
                    ws.Cells[w, q] = drv.Text.ToString();
                    //numero = ws.Cells[w, q];
                    q++;
                }
                q = 3;
                w++;
            }
            rg2.Columns.AutoFit();


            //mayor se debe envia como resultado
            ws7.Cells[107, 1] = "Temperatura max. : " + maxi1;
            ws8.Cells[108, 1] = "Temperatura min. : " + mini1;
            ws9.Cells[109, 1] = "Presión max. : " + maxi2;
            ws10.Cells[110, 1] = "Presión min. : " + mini2;
            ws11.Cells[111, 1] = "Tipo de válvula: Válvula de Purga Continua";   //realtime , realtime2
            ws12.Cells[112, 1] = "Observaciones: _____________________________________________________________________";   //realtime , realtime2
            ws13.Cells[113, 1] = "____________________________________________________________________________________";   //realtime , realtime2
            ws14.Cells[114, 1] = "____________________________________________________________________________________";   //realtime , realtime2
            ws15.Cells[118, 1] = "Elaboró: ______________________________________________________";   //realtime , realtime2
            ws16.Cells[122, 1] = "Revisó:  ______________________________________________________";   //realtime , realtime2
            ws.Protect("def-345", SheetProtectionType.All);
        }



        private void Purga_Intermitente()
        {


            Microsoft.Office.Interop.Excel.Application xla = new Microsoft.Office.Interop.Excel.Application();
            xla.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook wb = xla.Workbooks.Add(Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);
            Spire.Xls.Workbook book = new Spire.Xls.Workbook();
            Spire.Xls.Worksheet sheet = book.Worksheets[0];
            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws2 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws3 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws4 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws5 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws6 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws7 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws8 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws9 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws10 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws11 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws12 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws13 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws14 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws15 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws16 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws17 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws18 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws19 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws20 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws21 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws22 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            // Name the two cells A1. B1 are "Time (s)" and "Data" respectively, then automatically expand
            Microsoft.Office.Interop.Excel.Range rg18 = (Microsoft.Office.Interop.Excel.Range)ws18.get_Range("A1", "B1");
            Microsoft.Office.Interop.Excel.Range rg = (Microsoft.Office.Interop.Excel.Range)ws.get_Range("A10", "B10"); // EN LA COORDENADA A10,B10 ESTARÁN LOS VALORES DE TEMPERATUR
            Microsoft.Office.Interop.Excel.Range rg2 = (Microsoft.Office.Interop.Excel.Range)ws2.get_Range("C10", "D10");
            Microsoft.Office.Interop.Excel.Range rg3 = (Microsoft.Office.Interop.Excel.Range)ws3.get_Range("A38", "B38");
            Microsoft.Office.Interop.Excel.Range rg4 = (Microsoft.Office.Interop.Excel.Range)ws4.get_Range("A40", "B40");
            Microsoft.Office.Interop.Excel.Range rg19 = (Microsoft.Office.Interop.Excel.Range)ws19.get_Range("A53", "B53");
            Microsoft.Office.Interop.Excel.Range rg5 = (Microsoft.Office.Interop.Excel.Range)ws5.get_Range("A105", "B105");
            Microsoft.Office.Interop.Excel.Range rg6 = (Microsoft.Office.Interop.Excel.Range)ws6.get_Range("A106", "B106");
            Microsoft.Office.Interop.Excel.Range rg7 = (Microsoft.Office.Interop.Excel.Range)ws7.get_Range("A107", "B107");
            Microsoft.Office.Interop.Excel.Range rg8 = (Microsoft.Office.Interop.Excel.Range)ws8.get_Range("A108", "B108");
            Microsoft.Office.Interop.Excel.Range rg9 = (Microsoft.Office.Interop.Excel.Range)ws9.get_Range("A109", "B109");
            Microsoft.Office.Interop.Excel.Range rg10 = (Microsoft.Office.Interop.Excel.Range)ws10.get_Range("A110", "B110");
            Microsoft.Office.Interop.Excel.Range rg11 = (Microsoft.Office.Interop.Excel.Range)ws11.get_Range("A111", "B111");
            Microsoft.Office.Interop.Excel.Range rg12 = (Microsoft.Office.Interop.Excel.Range)ws12.get_Range("A112", "B112");
            Microsoft.Office.Interop.Excel.Range rg13 = (Microsoft.Office.Interop.Excel.Range)ws13.get_Range("A113", "B113");
            Microsoft.Office.Interop.Excel.Range rg14 = (Microsoft.Office.Interop.Excel.Range)ws14.get_Range("A114", "B114");
            Microsoft.Office.Interop.Excel.Range rg15 = (Microsoft.Office.Interop.Excel.Range)ws15.get_Range("A118", "B118");
            Microsoft.Office.Interop.Excel.Range rg16 = (Microsoft.Office.Interop.Excel.Range)ws16.get_Range("A122", "B122");
            Microsoft.Office.Interop.Excel.Range rg17 = (Microsoft.Office.Interop.Excel.Range)ws17.get_Range("A117", "B117");
            Microsoft.Office.Interop.Excel.Range rg20 = (Microsoft.Office.Interop.Excel.Range)ws20.get_Range("A118", "B118");
            Microsoft.Office.Interop.Excel.Range rg21 = (Microsoft.Office.Interop.Excel.Range)ws21.get_Range("A78", "B78");
            Microsoft.Office.Interop.Excel.Range rg22 = (Microsoft.Office.Interop.Excel.Range)ws22.get_Range("A78", "B78");
            ws3 = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.get_Item(1);
            ws4 = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.get_Item(1);
            ws5 = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.get_Item(1);
            ws6 = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.get_Item(1);
            //imagen reporte
            ws3.Shapes.AddPicture(rutaArchivo, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 20, 870, 400, 150);   // 710 mide los pixeles de una hoja. El primer parámetro posición X, segundo posicón Y, tercero tamaño en Y, cuarto tamaño en X.
            ws4.Shapes.AddPicture(rutaArchivo2, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 20, 1050, 400, 150); // 710 mide los pixeles de una hoja. El primer parámetro posición X, segundo posicón Y, tercero tamaño en Y, cuarto tamaño en X.
            ws5.Cells[105, 1] = "Fecha de inicio: " + IDate3;
            ws6.Cells[106, 1] = "Fecha límite elegida: " + IDate2;
            //imagen hoja membretada
            ws7.Shapes.AddPicture(path + "//Sin título.Jpeg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, 0, 400, 120);    // 710 mide los pixeles de una hoja. El primer parámetro posición X, segundo posicón Y, tercero tamaño en Y, cuarto tamaño en X.
            ws8.Shapes.AddPicture(path + "//Sin título2.Jpeg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, 585, 400, 120); // 710 mide los pixeles de una hoja. El primer parámetro posición X, segundo posicón Y, tercero tamaño en Y, cuarto tamaño en X.
            ws19.Shapes.AddPicture(path + "//Sin título.Jpeg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, 710, 400, 120);   // 710 mide los pixeles de una hoja. El primer parámetro posición X, segundo posicón Y, tercero tamaño en Y, cuarto tamaño en X.
            ws20.Shapes.AddPicture(path + "//Sin título2.Jpeg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, 1285, 400, 120); // 710 mide los pixeles de una hoja. El primer parámetro posición X, segundo posicón Y, tercero tamaño en Y, cuarto tamaño en X.
            ws21.Shapes.AddPicture(path + "//Sin título.Jpeg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, 1420, 400, 120);   // 710 mide los pixeles de una hoja. El primer parámetro posición X, segundo posicón Y, tercero tamaño en Y, cuarto tamaño en X.
            ws22.Shapes.AddPicture(path + "//Sin título2.Jpeg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, 1995, 400, 120); // 710 mide los pixeles de una hoja. El primer parámetro posición X, segundo posicón Y, tercero tamaño en Y, cuarto tamaño en X.
            ws.Cells[10, 1] = "Tiempo (segundos): ";
            ws.Cells[10, 2] = "Temperatura (°C): ";
            int i = 11; //tiempo1
            int j = 12; //°C
            foreach (ListViewItem comp in listView1.Items)
            {
                ws.Cells[i, j] = comp.Text.ToString();
                foreach (ListViewItem.ListViewSubItem drv in comp.SubItems)
                {
                    ws.Cells[i, j] = drv.Text.ToString();
                    j++;
                }
                j = 1;
                i++;
            }

            rg.Columns.AutoFit();
            ws2.Cells[10, 3] = "Tiempo (segundos): ";
            ws2.Cells[10, 4] = "Presión (PSI): ";
            int w = 11; //tiempo2
            int q = 12; //PSI
            foreach (ListViewItem comp in listView2.Items) //de esta manera se sale solo una casilla, es la de presión, y solo no concuerda esa de presión que se sale.
            {
                ws.Cells[w, q] = comp.Text.ToString();
                // ws.Protect("def-345", SheetProtectionType.All);
                foreach (ListViewItem.ListViewSubItem drv in comp.SubItems)
                {
                    ws.Cells[w, q] = drv.Text.ToString();
                    //numero = ws.Cells[w, q];
                    q++;
                }
                q = 3;
                w++;
            }
            rg2.Columns.AutoFit();


            //mayor se debe envia como resultado
            ws7.Cells[107, 1] = "Temperatura max. : " + maxi1;
            ws8.Cells[108, 1] = "Temperatura min. : " + mini1;
            ws9.Cells[109, 1] = "Presión max. : " + maxi2;
            ws10.Cells[110, 1] = "Presión min. : " + mini2;
            ws11.Cells[111, 1] = "Tipo de válvula: Válvula de Purga Intermitente";   //realtime , realtime2
            ws12.Cells[112, 1] = "Observaciones: _____________________________________________________________________";   //realtime , realtime2
            ws13.Cells[113, 1] = "____________________________________________________________________________________";   //realtime , realtime2
            ws14.Cells[114, 1] = "____________________________________________________________________________________";   //realtime , realtime2
            ws15.Cells[118, 1] = "Elaboró: ______________________________________________________";   //realtime , realtime2
            ws16.Cells[122, 1] = "Revisó:  ______________________________________________________";   //realtime , realtime2
            ws.Protect("def-345", SheetProtectionType.All);
        }



        // Function to save ListView to Excel
        private void SaveToExcel2()
        {
            Microsoft.Office.Interop.Excel.Application xla = new Microsoft.Office.Interop.Excel.Application();
            xla.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook wb = xla.Workbooks.Add(Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);

            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws2 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws7 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws8 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws19 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws20 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws21 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;
            Microsoft.Office.Interop.Excel.Worksheet ws22 = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;

            // Name the two cells A1. B1 are "Time (s)" and "Data" respectively, then automatically expand
            Microsoft.Office.Interop.Excel.Range rg = (Microsoft.Office.Interop.Excel.Range)ws.get_Range("A2", "B2");
            Microsoft.Office.Interop.Excel.Range rg2 = (Microsoft.Office.Interop.Excel.Range)ws2.get_Range("C2", "D2");
            Microsoft.Office.Interop.Excel.Range rg7 = (Microsoft.Office.Interop.Excel.Range)ws7.get_Range("A107", "B107");
            Microsoft.Office.Interop.Excel.Range rg8 = (Microsoft.Office.Interop.Excel.Range)ws8.get_Range("A108", "B108");
            Microsoft.Office.Interop.Excel.Range rg19 = (Microsoft.Office.Interop.Excel.Range)ws19.get_Range("A53", "B53");
            Microsoft.Office.Interop.Excel.Range rg20 = (Microsoft.Office.Interop.Excel.Range)ws20.get_Range("A118", "B118");
            Microsoft.Office.Interop.Excel.Range rg21 = (Microsoft.Office.Interop.Excel.Range)ws21.get_Range("A78", "B78");
            Microsoft.Office.Interop.Excel.Range rg22 = (Microsoft.Office.Interop.Excel.Range)ws22.get_Range("A78", "B78");
            //imagen hoja membretada
            ws7.Shapes.AddPicture(path + "//Sin título.Jpeg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, 0, 400, 120);    // 710 mide los pixeles de una hoja. El primer parámetro posición X, segundo posicón Y, tercero tamaño en Y, cuarto tamaño en X.
            ws8.Shapes.AddPicture(path + "//Sin título2.Jpeg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, 585, 400, 120); // 710 mide los pixeles de una hoja. El primer parámetro posición X, segundo posicón Y, tercero tamaño en Y, cuarto tamaño en X.
            ws19.Shapes.AddPicture(path + "//Sin título.Jpeg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, 710, 400, 120);   // 710 mide los pixeles de una hoja. El primer parámetro posición X, segundo posicón Y, tercero tamaño en Y, cuarto tamaño en X.
            ws20.Shapes.AddPicture(path + "//Sin título2.Jpeg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, 1285, 400, 120); // 710 mide los pixeles de una hoja. El primer parámetro posición X, segundo posicón Y, tercero tamaño en Y, cuarto tamaño en X.
            ws21.Shapes.AddPicture(path + "//Sin título.Jpeg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, 1420, 400, 120);   // 710 mide los pixeles de una hoja. El primer parámetro posición X, segundo posicón Y, tercero tamaño en Y, cuarto tamaño en X.
            ws22.Shapes.AddPicture(path + "//Sin título2.Jpeg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, 1995, 400, 120); // 710 mide los pixeles de una hoja. El primer parámetro posición X, segundo posicón Y, tercero tamaño en Y, cuarto tamaño en X.

            ws.Cells[57, 1] = "Tiempo (segundos): ";
            ws.Cells[57, 2] = "Temperatura (°C): ";
            ws2.Cells[57, 3] = "Tiempo (segundos): ";
            ws2.Cells[57, 4] = "Presión (PSI): ";
            ws.Cells[104, 1] = "Tiempo (segundos): ";
            ws.Cells[104, 2] = "Temperatura (°C): ";
            ws2.Cells[104, 3] = "Tiempo (segundos): ";
            ws2.Cells[104, 4] = "Presión (PSI): ";

            ws.Cells[10, 1] = "Tiempo (segundos): ";
            ws.Cells[10, 2] = "Temperatura (°C): ";
            int i = 11; //tiempo1
            int j = 12; //°C
            foreach (ListViewItem comp in listView1.Items)
            {
                ws.Cells[i, j] = comp.Text.ToString();
                foreach (ListViewItem.ListViewSubItem drv in comp.SubItems)
                {
                    ws.Cells[i, j] = drv.Text.ToString();
                    j++;
                }
                j = 1;
                i++;
            }

            rg.Columns.AutoFit();
            ws2.Cells[10, 3] = "Tiempo (segundos): ";
            ws2.Cells[10, 4] = "Presión (PSI): ";
            int w = 11; //tiempo2
            int q = 12; //PSI
            foreach (ListViewItem comp in listView2.Items) //de esta manera se sale solo una casilla, es la de presión, y solo no concuerda esa de presión que se sale.
            {
                ws.Cells[w, q] = comp.Text.ToString();
                // ws.Protect("def-345", SheetProtectionType.All);
                foreach (ListViewItem.ListViewSubItem drv in comp.SubItems)
                {
                    ws.Cells[w, q] = drv.Text.ToString();
                    //numero = ws.Cells[w, q];
                    q++;
                }
                q = 3;
                w++;
            }
            rg2.Columns.AutoFit();
            ws.Protect("def-345", SheetProtectionType.All);
        }


        private void BtConnect_Click_1(object sender, EventArgs e)
        {
            serialPort1.Close();

            if (!serialPort1.IsOpen)
            {
                serialPort1.PortName = comboBox1.Text; // Get COM port  
                serialPort1.BaudRate = 9600; // Baudrate is 9600, same with Arduino baudrate   
                try
                {
                    serialPort1.Open();
                    serialPort1.Write("1"); // Send character "2" via Serial 
                    Timer.Start();
                    Timer2.Start();
                    Graficar();
                    serialPort1.Write("3"); // Send "1" para realizar mediciones de temperatura y presión
                    bt_Save.Enabled = false;
                    btSalvarFoto.Enabled = false;
                    groupBox4.Visible = true;
                    bt_Cargar.Enabled = false;
                    bt_Connect.Enabled = false;
                    serialPort1.DataBits = 8;
                    serialPort1.StopBits = System.IO.Ports.StopBits.One;
                    serialPort1.Parity = Parity.None;
                    serialPort1.Handshake = Handshake.None;
                }
                catch
                {
                    MessageBox.Show("No se pudo abrir el puerto serial" + serialPort1.PortName, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void Bt_Cargar_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                rutaArchivo = openFileDialog.FileName;
            }
            OpenFileDialog openFileDialog2 = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                rutaArchivo2 = openFileDialog.FileName;
            }
            textBox1.Text = rutaArchivo + "  " + rutaArchivo2 + "  " + path;
            bt_Save.Enabled = true;
            btSalvarFoto.Enabled = false;
            groupBox4.Visible = true;
            bt_Cargar.Enabled = false;
            bt_Connect.Enabled = false;
        }

        private void BtSave_Click_1(object sender, EventArgs e)
        {

            DialogResult traloi;
            traloi = MessageBox.Show("Do you want to save data?", "Save", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
            if (traloi == DialogResult.OK && valv == 1)
            {
                Atemperadora_Sobrecalentado(); // Execute the function to save ListView to Excel 
                SaveToExcel2(); // Execute the function to save ListView to Excel 

                bt_Connect.Enabled = false;
                bt_Save.Enabled = false;
                btSalvarFoto.Enabled = false;
                groupBox4.Visible = true;
            }
            else if (traloi == DialogResult.OK && valv == 2)
            {
                Purga_Continua(); // Execute the function to save ListView to Excel 
                SaveToExcel2(); // Execute the function to save ListView to Excel 

                bt_Connect.Enabled = false;
                bt_Save.Enabled = false;
                btSalvarFoto.Enabled = false;
                groupBox4.Visible = true;
            }
            else if (traloi == DialogResult.OK && valv == 3)
            {
                Purga_Intermitente(); // Execute the function to save ListView to Excel 
                SaveToExcel2(); // Execute the function to save ListView to Excel 

                bt_Connect.Enabled = false;
                bt_Save.Enabled = false;
                btSalvarFoto.Enabled = false;
                groupBox4.Visible = true;
            }

        }


        private void Valvula_Dump_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult dialogo = MessageBox.Show("¿Desea cerrar el programa?",
                  "Cerrar el programa", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialogo == DialogResult.No)
            {
                e.Cancel = true;
                if (serialPort1.IsOpen)
                {
                    Timer.Stop();
                    Timer2.Stop();
                    ResetValue();
                    serialPort1.Write("1"); // Send character "2" via Serial para reiniciar los valores de los sensores y temporizadores
                    serialPort1.Close();
                }
            }
            else
            {
                e.Cancel = false;
            }
        }


        private void DateTimePicker1_ValueChanged_1(object sender, EventArgs e)
        {
            bt_Connect.Enabled = true;
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "yyyy/MM/dd hh:mm:ss";

            bt_Cargar.Enabled = false;
            bt_Connect.Enabled = false;
            bt_Save.Enabled = false;
            btSalvarFoto.Enabled = false;
            groupBox4.Visible = true;
            bt_Cargar.Enabled = false;
        }
        private void BtSalvarFoto_Click_1(object sender, EventArgs e)
        {
            path = Directory.GetCurrentDirectory();
            this.zedGraphControl1.SaveAsBitmap(); //los guarda bien, pero ¿cómo saber la ruta si el usuario ya no dejó capturarla... 
            this.zedGraphControl2.SaveAsBitmap();

            bt_Cargar.Enabled = true;
            bt_Save.Enabled = false;
            bt_Connect.Enabled = false;
            btSalvarFoto.Enabled = false;
            bt_Cargar.Enabled = true;
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            groupBox4.Visible = true;
            radioButton12.Enabled = false;
            radioButton13.Enabled = false;
            radioButton14.Enabled = false;
            this.Close();
        }

        private void RadioButton12_CheckedChanged(object sender, EventArgs e)
        {

            bt_Connect.Enabled = true;
            bt_Save.Enabled = false;
            btSalvarFoto.Enabled = false;
            groupBox4.Visible = true;


            radioButton12.Enabled = true;
            radioButton13.Enabled = false;
            radioButton14.Enabled = false;
            valv = 1;
        }

        private void RadioButton13_CheckedChanged(object sender, EventArgs e)
        {
            bt_Connect.Enabled = true;
            bt_Save.Enabled = false;
            btSalvarFoto.Enabled = false;
            groupBox4.Visible = true;
            radioButton12.Enabled = false;
            radioButton13.Enabled = true;
            radioButton14.Enabled = false;

            valv = 2;
        }

        private void RadioButton14_CheckedChanged(object sender, EventArgs e)
        {
            bt_Connect.Enabled = true;
            bt_Save.Enabled = false;
            btSalvarFoto.Enabled = false;
            groupBox4.Visible = true;
            radioButton12.Enabled = false;
            radioButton13.Enabled = false;
            radioButton14.Enabled = true;
            valv = 3;
        }

        private void Panel1_MouseDown(object sender, MouseEventArgs e)
        {

        }


        private void Panel2_MouseDown(object sender, MouseEventArgs e)
        {
            RealeaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private void Panel2_Paint(object sender, PaintEventArgs e)
        {

        }

    }
  }

