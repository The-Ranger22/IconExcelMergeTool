using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Win32;

namespace ExcelMerge
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        public MainWindow()
        {
            InitializeComponent();
            
        }

        private void _selectFiles(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Multiselect = true;
            
            if (ofd.ShowDialog() == true)
            {
                TBlock.Height = LGrid.RowDefinitions[1].ActualHeight;
                TBlock.Background = Brushes.Gainsboro;
                string s = "";
                foreach (string fname in ofd.FileNames)
                {
                    s += $"{fname}\n";
                }
                TBlock.Text = s;
            }

        }

        void _mergeFiles(object sender, RoutedEventArgs e)
        {
            
        }

        private void DisplaySelectedFiles()
        {
            
        }

    }
}