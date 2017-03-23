﻿using System;
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
using System.Windows.Shapes;

namespace SpreadsheetParser
{
    /// <summary>
    /// Interaction logic for SpreadsheetParser.xaml
    /// </summary>
    public partial class SpreadsheetParserHelper : Window
    {
        public SpreadsheetParserHelper()
        {
            InitializeComponent();
            this.DataContext = _vm;
        }

        private VMSpreadsheetParser _vm = new VMSpreadsheetParser();

        private void TextBox_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 3)
            {
                myTextBox.SelectAll();
            }
        }
    }
}
