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
using System.Data;


namespace Timesheet_remainder
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string _fileLoadPath = string.Empty; //new file file path
        private DateTime _sheetDateTime;
        private int _popupCounter = 0;
        private readonly ExcelManager _excelManager = new ExcelManager();

        public MainWindow()
        {
            InitializeComponent();


            //Sets sheetDate text block to the current time
            sheetDate.Text = DateTime.Now.ToString("dd/MM/yy   HH:mm");
            _sheetDateTime = DateTime.Now;


            //Sets status to no file load

            statusMsg.Foreground = System.Windows.Media.Brushes.Red;
            statusMsg.Text = "Timesheet not loaded";


            NewTimer();

            TrayManagement();
        }

        protected override void OnStateChanged(EventArgs e)
        {
            //State change events minimise to hide window
            if (WindowState == System.Windows.WindowState.Minimized)
                this.Hide();

            base.OnStateChanged(e);
        }

        #region Auto Popup
        private void NewTimer()
        {
            // Create a timer and set a 20 s interval.
            var timer = new System.Timers.Timer();
            const int intervalSeconds = 20 * 1000;
            
            timer.Interval = intervalSeconds;

            // Hook up the Elapsed event for the timer. 
            timer.Elapsed += TimedPopupEvent;

            // Have the timer fire repeated events (true is the default)
            timer.AutoReset = true;

            // Start the timer
            timer.Enabled = true;
        }

        private void TimedPopupEvent(Object source, System.Timers.ElapsedEventArgs e)
        {
            //Console.WriteLine("The Elapsed event was raised at {0}", e.SignalTime);  //use for debug
            DateTime currentTime = DateTime.Now;
            
            this.Dispatcher.Invoke(() =>
            {
                sheetDate.Text = DateTime.Now.ToString("dd/MM/yy   HH:mm");
                _sheetDateTime = DateTime.Now;

                //Autopopup at set minutes
                if (currentTime.Minute == 30 | currentTime.Minute == 00)
                {                    
                    if (_popupCounter == 0)
                    {                        
                        this.Show();
                        this.WindowState = WindowState.Normal;                       
                        _popupCounter += 1;

                        var debugString =
                            $"timed popup triggered at Time = {currentTime.Minute}, counter = {_popupCounter}"; 
                        Console.WriteLine(debugString);
                    }                                                           
                    //Console.WriteLine(counter.ToString()); //debug output for 
                } 
                else
                {
                    _popupCounter = 0;
                }             
            });
        }


        #endregion

        private void TrayManagement()
        {
            //Minimise to tray
            var notifyIcon = new System.Windows.Forms.NotifyIcon
            {
                Icon = System.Drawing.Icon.ExtractAssociatedIcon(System.Windows.Forms.Application.ExecutablePath),
                Visible = true
            };

            //Maximise when double click            
            notifyIcon.Click +=
                delegate (object sender, EventArgs args)
                {
                    this.Show();
                    _sheetDateTime = DateTime.Now;
                    this.WindowState = WindowState.Normal;
                };
        }

        #region Buttons
        private void btnNewSheet_Click(object sender, RoutedEventArgs e)
        {
            var newExcelSheetDialogue = new Microsoft.Win32.SaveFileDialog {DefaultExt = ".xlsx"};

            var newExcelSheetDialogueSuccessful = newExcelSheetDialogue.ShowDialog();
            if (newExcelSheetDialogueSuccessful == true)
            {
                _fileLoadPath = newExcelSheetDialogue.FileName;
            }

            if (newExcelSheetDialogue.FileName != string.Empty)
            {
                try
                {
                    _excelManager.NewExcelFile(_fileLoadPath);
                    statusMsg.Text = string.Empty;
                }
                catch
                {
                    statusMsg.Foreground = System.Windows.Media.Brushes.Red;
                    statusMsg.Text = "Invalid operation";
                }
            }
        }

        private void btnLoadSheet_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog {DefaultExt = ".xlsx"};

            bool? showDialogSuccessful = openFileDialog.ShowDialog();

            if (showDialogSuccessful == true)
            {
                _fileLoadPath = openFileDialog.FileName;
                statusMsg.Text = string.Empty;
            }

        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (_fileLoadPath != string.Empty && statusMsg.Text != "Invalid operation")
            {
                var fileInfoPath = new FileInfo(_fileLoadPath);

                using (var excelPackage = new ExcelPackage(fileInfoPath))
                {

                    try
                    {
                        _excelManager.AddNewEntryToWorkSheet(_fileLoadPath, _sheetDateTime, txtTaskInput.Text);
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

        private void btnCalc_Click(object sender, RoutedEventArgs e)
        {
            if (_fileLoadPath != string.Empty && statusMsg.Text != "Invalid operation")
            {
                try
                {
                    var fi = new FileInfo(_fileLoadPath);
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
        #endregion
    }

}


