using Microsoft.Office.Interop.Excel;
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
    public class ViewModel : INotifyPropertyChanged
    {
        #region Commands

        #region SpreadSheetParserCommand

        private ICommand _spreadSheetParserCommand;
        public ICommand SpreadSheetParserCommand
        {
            get
            {
                return _spreadSheetParserCommand ?? (_spreadSheetParserCommand = new CommandHandler(() => SpreadSheetAction(), true));
            }
        }
        //private bool _canExecute;
        public void SpreadSheetAction()
        {
            SpreadsheetParserHelper parser = new SpreadsheetParserHelper();
            parser.ShowDialog();
        }

        #endregion SpreadSheetParserCommand

        #region CwHelperCommand

        private ICommand _cwHelperCommand;
        public ICommand CwHelperCommand
        {
            get
            {
                return _cwHelperCommand ?? (_cwHelperCommand = new CommandHandler(() => HelperAction(), true));
            }
        }
        //private bool _canExecute;
        public void HelperAction()
        {
            CwApiHelper cw = new CwApiHelper();
            cw.ShowDialog();
        }

        #endregion CwHelperCommand

        #endregion Commands

        #region Events

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion Events
    }

    #region Helper class

    public class CommandHandler : ICommand
    {
        private System.Action _action;
        private bool _canExecute;
        public CommandHandler(System.Action action, bool canExecute)
        {
            _action = action;
            _canExecute = canExecute;
        }

        public bool CanExecute(object parameter)
        {
            return _canExecute;
        }

        public event EventHandler CanExecuteChanged;

        public void Execute(object parameter)
        {
            _action();
        }
    }

    #endregion Helper class
}
