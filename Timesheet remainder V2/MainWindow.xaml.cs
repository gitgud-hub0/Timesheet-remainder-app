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
using System.Windows.Shapes;


using System.IO;
using OfficeOpenXml;
using System.Drawing;
using System.Timers;
using System.Data;


namespace Timesheet_remainder
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string loadPath = string.Empty; //new file file path
        string sheetDay = DateTime.Now.ToString("dd/MM/yy");
        string sheetTime = DateTime.Now.ToString("HH:mm:ss");

        int popupCounter = 0;

        private static Timer aTimer;

        public MainWindow()
        {
            InitializeComponent();


            //Sets sheetDate text block to the current time
            sheetDate.Text = DateTime.Now.ToString("dd/MM/yy   HH:mm");


            //Sets status to no file load

            statusMsg.Foreground = System.Windows.Media.Brushes.Red;
            statusMsg.Text = "Timesheet not loaded";


            // Create a timer and set a 20 s interval.
            aTimer = new System.Timers.Timer();
            aTimer.Interval = 20 * 1000;

            // Hook up the Elapsed event for the timer. 
            aTimer.Elapsed += OnTimedEvent;

            // Have the timer fire repeated events (true is the default)
            aTimer.AutoReset = true;

            // Start the timer
            aTimer.Enabled = true;



            //Minimise to tray
            System.Windows.Forms.NotifyIcon ni = new System.Windows.Forms.NotifyIcon();

            //set the tray icon
            //ni.Icon = new System.Drawing.Icon(@filePath); //tray icon is in the path
            ni.Icon = System.Drawing.Icon.ExtractAssociatedIcon(System.Windows.Forms.Application.ExecutablePath);
            //ni.Icon = new Icon(this.Icon, 40, 40); //trayIcon is your NotifyIcon

            //System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName;
            ni.Visible = true;

            //Maximise when double click            
            ni.Click +=
                delegate (object sender, EventArgs args)
                {
                    this.Show();
                    this.WindowState = WindowState.Normal;
                };
        }

        //State change events minimise to hide window
        protected override void OnStateChanged(EventArgs e)
        {
            if (WindowState == System.Windows.WindowState.Minimized)
                this.Hide();

            base.OnStateChanged(e);
        }

        //Timer event
        private void OnTimedEvent(Object source, System.Timers.ElapsedEventArgs e)
        {
            //Console.WriteLine("The Elapsed event was raised at {0}", e.SignalTime);  //use for debug
            DateTime currentTime = DateTime.Now;
            

            this.Dispatcher.Invoke(() =>
            {
                sheetDate.Text = DateTime.Now.ToString("dd/MM/yy   HH:mm");
                sheetDay = DateTime.Now.ToString("dd/MM/yyyy");
                sheetTime = DateTime.Now.ToString("HH:mm:ss");
                //Autopopup at set minutes
                if (currentTime.Minute == 30 | currentTime.Minute == 00)
                {                    
                    if (popupCounter == 0)
                    {                        
                        this.Show();
                        this.WindowState = WindowState.Normal;                       
                        popupCounter += 1;
                        Console.WriteLine("autopopup triggered, counter= " + popupCounter.ToString() + ", Time = " + currentTime.Minute.ToString()); //debug
                    }                                                           
                    //Console.WriteLine(counter.ToString()); //debug output for 
                } 
                else
                {
                    popupCounter = 0;
                }             
            });
        }

        //buttons
        private void btnNewSheet_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();

            dlg.DefaultExt = ".xlsx";

            Nullable<bool> result = dlg.ShowDialog();

            if (result == true)
            {
                // Open document 
                loadPath = dlg.FileName;
            }

            Console.WriteLine(loadPath); //use for debug

            if (dlg.FileName != "")
            {
                var fi = new FileInfo(@loadPath);
                using (var p = new ExcelPackage(fi))
                {
                    //Get the Worksheet created in the previous codesample. 
                    try
                    {
                        var ws = p.Workbook.Worksheets.Add("MySheet");


                        //Set headers
                        ws.Cells["A1"].Value = "Date";
                        ws.Cells["B1"].Value = "Time";
                        ws.Cells["C1"].Value = "Task Description";

                        p.Save();

                        statusMsg.Text = string.Empty; //removes timesheet error msg
                    }
                    catch
                    {
                        statusMsg.Foreground = System.Windows.Media.Brushes.Red;
                        statusMsg.Text = "Invalid operation";
                    }

                }

                /*               // Saves the Image via a FileStream created by the OpenFile method.
                               System.IO.FileStream fs =
                                   (System.IO.FileStream)dlg.OpenFile();
                               fs.Close();*/
            }
        }

        private void btnLoadSheet_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".xlsx";

            // Display OpenFileDialog by calling ShowDialog method 
            Nullable<bool> result = dlg.ShowDialog();

            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                loadPath = dlg.FileName;
                statusMsg.Text = string.Empty;
            }

        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (loadPath != string.Empty && statusMsg.Text != "Invalid operation")
            {
                var fi = new FileInfo(@loadPath);

                using (var p = new ExcelPackage(fi))
                {

                    try
                    {
                        //Get the Worksheet created in the previous codesample. 
                        var ws = p.Workbook.Worksheets["MySheet"];
                        var lastRow = ws.Dimension.End.Row;
                        var lastColumn = ws.Dimension.End.Column;

                        //Set the next last cell in the row to sheetDate.Text
                        ws.Cells[lastRow + 1, 1].Value = sheetDay;
                        //Set the next last cell in the row to txtTaskInput.Text;
                        ws.Cells[lastRow + 1, 2].Value = sheetTime;
                        ws.Cells[lastRow + 1, 3].Value = txtTaskInput.Text;

                        p.Save();
                        statusMsg.Text = string.Empty;

                        this.WindowState = WindowState.Minimized;
                    }
                    catch
                    {
                        statusMsg.Foreground = System.Windows.Media.Brushes.Red;
                        statusMsg.Text = "Invalid operation";
                    }
                }
            }
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            //txtTaskInput.Text = String.Empty;
            this.WindowState = WindowState.Minimized;
        }

/*        public class Tasks
        {
            public string CalcTask { get; set; }
            public double CalcTime { get; set; }

            public Tasks(string calcTask = "No Name", double calcTime = 0) //constructor
            {
                CalcTask = calcTask;
                CalcTime = calcTime;
            }    
        }*/


        private void btnCalc_Click(object sender, RoutedEventArgs e)
        {
            if (loadPath != string.Empty && statusMsg.Text != "Invalid operation")
            {
                try
                {
                    var fi = new FileInfo(@loadPath);
                    using (var p = new ExcelPackage(fi))
                    {


                        List<Object> instanceList = new List<Object>();

                        var wsTimes = p.Workbook.Worksheets["MySheet"];
                        var wsTimesStart = wsTimes.Dimension.Start;
                        var wsTimesEnd = wsTimes.Dimension.End;

                        //this needs refactoring to import straight to the datatable
                        for (int row = wsTimesStart.Row + 1; row <= wsTimesEnd.Row; row++)
                        {
                            instanceList.Add(wsTimes.Cells[row, 1].Value);
                            instanceList.Add(wsTimes.Cells[row, 2].Value);
                            instanceList.Add(wsTimes.Cells[row, 3].Value);
                        }


                        //mainly used for debug not output to excel file
                        //List<Object> calcList = new List<Object>();
                        /*                    string taskDate;
                                            string nextTaskDate;
                                            string taskTime;
                                            string nextTaskTime;
                                            TimeSpan calcTaskTime;
                                            string taskDesc;
                                            string nextTaskDesc;
                                            */
                        /*                    for (int order = 0; order <= instanceList.Count; order += 3)
                                            {


                                                if (instanceList.Count >= order + 5)
                                                {
                                                    taskDate = (string)instanceList[order];
                                                    taskTime = (string)instanceList[order + 1];
                                                    taskDesc = (string)instanceList[order + 2];
                                                    nextTaskDate = (string)instanceList[order + 3];
                                                    nextTaskTime = (string)instanceList[order + 4];
                                                    nextTaskDesc = (string)instanceList[order + 5];

                                                    if (nextTaskDate != null && nextTaskTime != null && nextTaskDesc != null)
                                                    {
                                                        //if  (taskDesc == nextTaskDesc && taskDate == nextTaskDate)
                                                        if (taskDate == nextTaskDate)
                                                        {
                                                            calcList.Add(taskDesc);
                                                            calcTaskTime = DateTime.Parse(nextTaskTime).Subtract(DateTime.Parse(taskTime));
                                                            calcList.Add(calcTaskTime);
                                                            Console.WriteLine("taskDecs =" + taskDesc + " calcTaskTime =" + calcTaskTime);
                                                        }
                                                    }
                                                }
                                            }*/

                        DataTable inputTable = new DataTable("inputTable1");
                        inputTable.Clear();
                        inputTable.Columns.Add("Date");
                        inputTable.Columns.Add("Time");
                        inputTable.Columns.Add("TaskDesc");
                        for (int order = 0; order <= instanceList.Count - 2; order += 3)
                        {
                            DataRow _task = inputTable.NewRow();
                            _task["Date"] = Convert.ToString(instanceList[order]);
                            _task["Time"] = Convert.ToString(instanceList[order + 1]);
                            _task["TaskDesc"] = Convert.ToString(instanceList[order + 2]);
                            inputTable.Rows.Add(_task);
                        }
                        //inputTable.WriteXml(@"C:\Users\Ben\Desktop\Timesheet test\calcTest\inputTable.xml");

                        DataTable calcTable = new DataTable("calcTable1");
                        calcTable.Clear();
                        calcTable.Columns.Add("Date");
                        calcTable.Columns.Add("Task");
                        calcTable.Columns.Add("Duration");
                        for (int i = 0; i < inputTable.Rows.Count - 1; i++)
                        {
                            // Compare with previous row using index
                            if (inputTable.Rows[i]["Date"] == inputTable.Rows[i + 1]["Date"])
                            {
                                DataRow _calcRow = calcTable.NewRow();
                                _calcRow["Date"] = inputTable.Rows[i]["Date"];
                                _calcRow["Task"] = inputTable.Rows[i]["TaskDesc"];
                                _calcRow["Duration"] = (DateTime.Parse((string)inputTable.Rows[i + 1]["Time"]).Subtract(DateTime.Parse((string)inputTable.Rows[i]["Time"]))).TotalMinutes;
                                calcTable.Rows.Add(_calcRow);
                            }
                        }
                        //sort table in terms of date then tasks using linq
                        DataTable sortCalcTable = new DataTable();
                        sortCalcTable.Clear();
                        sortCalcTable = calcTable.AsEnumerable()
                            .OrderBy(r => r.Field<string>("Date"))
                            .ThenBy(r => r.Field<string>("Task"))
                            .CopyToDataTable();
                        sortCalcTable.TableName = "sortCalcTable1";
                        sortCalcTable.WriteXml(@"C:\Users\Ben\Desktop\Timesheet test\calcTest\calcTable.xml");

                        //output table
                        DataTable outTable = new DataTable("OutputTable1");
                        outTable.Clear();
                        outTable.Columns.Add("Date");
                        outTable.Columns.Add("Task");
                        outTable.Columns.Add("TotalTime");
                        outTable.Columns.Add("TotalDuration", typeof(int));

                        //generate rows with unique tasks each 
                        DataRow _outRow1st = outTable.NewRow();
                        _outRow1st["Date"] = sortCalcTable.Rows[0]["Date"];
                        _outRow1st["Task"] = sortCalcTable.Rows[0]["Task"];
                        _outRow1st["TotalDuration"] = 0;
                        outTable.Rows.Add(_outRow1st);
                        //int j = 0; //iterator for outtable
                        int sum = 0;
                        TimeSpan ts;
                        for (int i = 0; i < sortCalcTable.Rows.Count; i++) //iterator for sortCalctable
                        {
                            if (sortCalcTable.Rows[i]["Date"] == outTable.Rows[outTable.Rows.Count - 1]["Date"])
                            {
                                if (sortCalcTable.Rows[i]["Task"] != outTable.Rows[outTable.Rows.Count - 1]["Task"])
                                {
                                    DataRow _outRow = outTable.NewRow();
                                    _outRow["Date"] = sortCalcTable.Rows[i]["Date"];
                                    _outRow["Task"] = sortCalcTable.Rows[i]["Task"];
                                    _outRow["TotalDuration"] = 0;
                                    _outRow["TotalTime"] = 0;
                                    outTable.Rows.Add(_outRow);
                                }
                            }
                            else
                            {
                                DataRow _outRow = outTable.NewRow();
                                _outRow["Date"] = sortCalcTable.Rows[i]["Date"];
                                _outRow["Task"] = sortCalcTable.Rows[i]["Task"];
                                _outRow["TotalDuration"] = 0;
                                _outRow["TotalTime"] = 0;
                                outTable.Rows.Add(_outRow);
                            }

                            if (sortCalcTable.Rows[i]["Date"] == outTable.Rows[outTable.Rows.Count - 1]["Date"]
                                && sortCalcTable.Rows[i]["Task"] == outTable.Rows[outTable.Rows.Count - 1]["Task"])
                            {
                                sum = Convert.ToInt32(outTable.Rows[outTable.Rows.Count - 1]["TotalDuration"])
                                    + Convert.ToInt32(sortCalcTable.Rows[i]["Duration"]);
                                ts = TimeSpan.FromMinutes(sum);
                                DataRow _outRow = outTable.Rows[outTable.Rows.Count - 1];
                                _outRow["TotalDuration"] = sum;
                                _outRow["TotalTime"] = ts.ToString("hh\\:mm");

                                sum = 0;
                            }
                        }

                        //outTable.WriteXml(@"C:\Users\Ben\Desktop\Timesheet test\calcTest\outTable.xml");

                        //export datatable to excel sheet


                        var wsCalc = p.Workbook.Worksheets.Add("Calculated_Times");
                        wsCalc.Cells["A1"].LoadFromDataTable(outTable, true);
                        wsCalc.DeleteColumn(4);
                        p.Save();

                        statusMsg.Foreground = System.Windows.Media.Brushes.Green;
                        statusMsg.Text = "Calculation completed";
                    }

                }
                catch
                {
                    statusMsg.Foreground = System.Windows.Media.Brushes.Red;
                    statusMsg.Text = "Invalid operation, calculation failed";
                }

            }            
             else
            {
                statusMsg.Foreground = System.Windows.Media.Brushes.Red;
                statusMsg.Text = "Invalid operation";
            }
        }
    }

}


