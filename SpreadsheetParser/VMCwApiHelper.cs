using Microsoft.Office.Interop.Excel;
using SpreadsheetParser.ConnectWise;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace SpreadsheetParser
{
    public class VMCwApiHelper : INotifyPropertyChanged
    {
        #region Commands

        #region Submit Button

        private ICommand _clickCommand;
        public ICommand ClickCommand
        {
            get
            {
                return _clickCommand ?? (_clickCommand = new CommandHandler(() => MyAction(), CanExecute()));
            }
        }

        private bool CanExecute()
        {
            //return !string.IsNullOrEmpty(Column) && !string.IsNullOrEmpty(StartRow) &&
            //    !string.IsNullOrEmpty(Column) && File.Exists(FileName);
            return true;
        }

        //private bool _canExecute;
        public void MyAction()
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application xlApp;
                Microsoft.Office.Interop.Excel.Workbook xlWorkbook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorksheet;
                object misValue = System.Reflection.Missing.Value;


                xlApp = new Microsoft.Office.Interop.Excel.Application();
                xlWorkbook = xlApp.Workbooks.Open(FileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);

                //int startColumn = Header.Cells.Column;
                //int startRow = header.Cells.Row + 1;
                int col = ExcelColumnNameToNumber(Column);
                Microsoft.Office.Interop.Excel.Range startCell = xlWorksheet.Cells[StartRow, col];
                Microsoft.Office.Interop.Excel.Range endCell = xlWorksheet.Cells[EndRow, col];
                Microsoft.Office.Interop.Excel.Range myRange = xlWorksheet.Range[startCell, endCell];
                System.Array myvalues = (System.Array)myRange.Cells.Value;
                string[] strArray = myvalues.OfType<object>().Select(o => o.ToString()).ToArray();

                foreach (var ticketId in strArray)
                {
                    try
                    {
                        IConnectWiseService _connectWiseService = new ConnectWiseService(Company, BaseUrl, SiteUrl, SiteSuffix, PublicKey, PrivateKey);
                        //var res = _connectWiseService.ChangeCompany(Convert.ToInt32(ticketId), Value).Result;
                        var res = _connectWiseService.ChangeGenerically(Convert.ToInt32(ticketId), Value, Op, Path).Result;
                    }
                    catch (Exception ex)
                    {
                    }

                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        public int ExcelColumnNameToNumber(string columnName)
        {
            if (string.IsNullOrEmpty(columnName)) throw new ArgumentNullException("columnName");

            columnName = columnName.ToUpperInvariant();

            int sum = 0;

            for (int i = 0; i < columnName.Length; i++)
            {
                sum *= 26;
                sum += (columnName[i] - 'A' + 1);
            }

            return sum;
        }

        #endregion Submit Button

        #region Browse Button

        private ICommand _browseCommand;
        public ICommand BrowseCommand
        {
            get
            {
                return _browseCommand ?? (_browseCommand = new CommandHandler(() => BrowseAction(), true));
            }
        }
        //private bool _canExecute;
        public void BrowseAction()
        {
            // Create OpenFileDialog 
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();



            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "Excel Files (*.xlsx)|*.xlsx|Old Excel Files (*.xls)|*.xls";


            // Display OpenFileDialog by calling ShowDialog method 
            Nullable<bool> result = dlg.ShowDialog();


            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                string filename = dlg.FileName;
                FileName = filename;
            }
        }

        #endregion Browse Button

        #endregion Commands

        #region Properties

        private string _fileName;

        public string FileName
        {
            get { return _fileName; }
            set
            {
                if (value != _fileName)
                {
                    _fileName = value;
                    OnPropertyChanged("FileName");
                }
            }
        }

        private string _column = "e";

        public string Column
        {
            get { return _column; }
            set
            {
                if (value != _column)
                {
                    _column = value;
                    OnPropertyChanged("Column");
                }
            }
        }

        private string _startRow = "2";

        public string StartRow
        {
            get { return _startRow; }
            set
            {
                if (value != _startRow)
                {
                    _startRow = value;
                    OnPropertyChanged("StartRow");
                }
            }
        }

        private string _endRow = "4";

        public string EndRow
        {
            get { return _endRow; }
            set
            {
                if (value != _endRow)
                {
                    _endRow = value;
                    OnPropertyChanged("EndRow");
                }
            }
        }

        private string _result;

        public string Result
        {
            get { return _result; }
            set
            {
                if (value != _result)
                {
                    _result = value;
                    OnPropertyChanged("Result");
                }
            }
        }

        private string _baseUrl = "https://connectwiselab.yourcompany.com/v4_6_release/apis/3.0/";

        public string BaseUrl
        {
            get { return _baseUrl; }
            set
            {
                if (value != _baseUrl)
                {
                    _baseUrl = value;
                    OnPropertyChanged("BaseUrl");
                }
            }
        }

        private string _siteUrl = "service/tickets";

        public string SiteUrl
        {
            get { return _siteUrl; }
            set
            {
                if (value != _siteUrl)
                {
                    _siteUrl = value;
                    OnPropertyChanged("SiteUrl");
                }
            }
        }

        private string _siteSuffix = "";

        public string SiteSuffix
        {
            get { return _siteSuffix; }
            set
            {
                if (value != _siteSuffix)
                {
                    _siteSuffix = value;
                    OnPropertyChanged("SiteSuffix");
                }
            }
        }

        private string _path = "company";

        public string Path
        {
            get { return _path; }
            set
            {
                if (value != _path)
                {
                    _path = value;
                    OnPropertyChanged("Path");
                }
            }
        }

        private string _value = "5134";

        public string Value
        {
            get { return _value; }
            set
            {
                if (value != _value)
                {
                    _value = value;
                    OnPropertyChanged("Value");
                }
            }
        }

        private string _op = "replace";

        public string Op
        {
            get { return _op; }
            set
            {
                if (value != _op)
                {
                    _op = value;
                    OnPropertyChanged("Op");
                }
            }
        }

        private string _company = "yourcompany";

        public string Company
        {
            get { return _company; }
            set
            {
                if (value != _company)
                {
                    _company = value;
                    OnPropertyChanged("Company");
                }
            }
        }

        private string _publicKey = "";

        public string PublicKey
        {
            get { return _publicKey; }
            set
            {
                if (value != _publicKey)
                {
                    _publicKey = value;
                    OnPropertyChanged("PublicKey");
                }
            }
        }

        private string _privateKey = "";

        public string PrivateKey
        {
            get { return _privateKey; }
            set
            {
                if (value != _privateKey)
                {
                    _privateKey = value;
                    OnPropertyChanged("PrivateKey");
                }
            }
        }

        #endregion Properties

        #region Events

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion Events
    }
}
