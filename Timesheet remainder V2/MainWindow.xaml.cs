using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Controls;


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
        private readonly ExcelController _excelController;
        private readonly DataTableController _dataTableController;

        public MainWindow()
        {
            InitializeComponent();

            sheetDate.Text = DateTime.Now.ToString("dd/MM/yy   HH:mm");
            _sheetDateTime = DateTime.Now;
            _excelController = new ExcelController();
            _dataTableController = new DataTableController();

            statusMsg.Foreground = System.Windows.Media.Brushes.Red;
            statusMsg.Text = "Timesheet not loaded";

            NewTimer();
            TrayManagement();
        }

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
        private void ExcelCalculations()
        {
            try
            {
                var fi = new FileInfo(_fileLoadPath);
                using (var excelPackage = new ExcelPackage(fi))
                {
                    List<Object> instanceList = new List<Object>();

                    var wsTimes = excelPackage.Workbook.Worksheets["MySheet"];
                    var wsTimesStart = wsTimes.Dimension.Start;
                    var wsTimesEnd = wsTimes.Dimension.End;

                    //TODO: this needs refactoring to import straight to the datatable
                    for (int row = wsTimesStart.Row + 1; row <= wsTimesEnd.Row; row++)
                    {
                        instanceList.Add(wsTimes.Cells[row, 1].Value);
                        instanceList.Add(wsTimes.Cells[row, 2].Value);
                        instanceList.Add(wsTimes.Cells[row, 3].Value);
                    }

                    var outTable = new DataTableController().GetOutputDataTable(instanceList);

                    //export datatable to excel sheet
                    var wsCalc = excelPackage.Workbook.Worksheets.Add("Calculated_Times");
                    wsCalc.Cells["A1"].LoadFromDataTable(outTable, true);
                    wsCalc.DeleteColumn(4);
                    excelPackage.Save();

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

        #region Timed Management
        private void NewTimer()
        {
            var timer = new System.Timers.Timer();
            const int intervalSeconds = 20 * 1000;
            timer.Interval = intervalSeconds;
            timer.Elapsed += TimedEvents;
            timer.AutoReset = true;
            timer.Enabled = true;
        }

        private void TimedEvents(Object source, System.Timers.ElapsedEventArgs e)
        {
            //Console.WriteLine("The Elapsed event was raised at {0}", e.SignalTime);  //use for debug
            this.Dispatcher.Invoke(() =>
            {
                UpdateTimes();
                AutoPopupAtTime();
            });
        }

        private void UpdateTimes()
        {
            sheetDate.Text = DateTime.Now.ToString("dd/MM/yy    HH:mm");
            _sheetDateTime = DateTime.Now;
        }

        private void AutoPopupAtTime()
        {
            DateTime currentTime = DateTime.Now;
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
            }
            else
            {
                _popupCounter = 0;
            }
        }

        protected override void OnStateChanged(EventArgs e)
        {
            //State change events minimise to hide window
            if (WindowState == System.Windows.WindowState.Minimized)
                this.Hide();

            base.OnStateChanged(e);
        }
        #endregion

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
                    _excelController.NewExcelFile(_fileLoadPath);
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

        private void btnBreak_Click(object sender, RoutedEventArgs e)
        {
            if (_fileLoadPath != string.Empty && statusMsg.Text != "Invalid operation")
            {
                try
                {
                    _excelController.AddNewEntryToWorkSheet(_fileLoadPath, _sheetDateTime, "Break");
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

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (_fileLoadPath != string.Empty && statusMsg.Text != "Invalid operation")
            {
                try
                {
                    _excelController.AddNewEntryToWorkSheet(_fileLoadPath, _sheetDateTime, ComboBoxTaskInput.Text);
                    DropDownListModel.AddToDropDownList(ComboBoxTaskInput.Text);
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

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            //txtTaskInput.Text = String.Empty;
            this.WindowState = WindowState.Minimized;
        }

        private void btnCalc_Click(object sender, RoutedEventArgs e)
        {
            if (_fileLoadPath != string.Empty && statusMsg.Text != "Invalid operation")
            {
                ExcelCalculations();
            }            
            else
            {
                statusMsg.Foreground = System.Windows.Media.Brushes.Red;
                statusMsg.Text = "Invalid operation";
            }
        }

        private void ComboBoxTaskInput_DropDownOpened(object sender, EventArgs e)
        {
            var comboBoxTaskInputController = new ComboBoxTaskInputController(ComboBoxTaskInput);
            comboBoxTaskInputController.PopulateDropDownBox();
        }
        #endregion
    }
}


