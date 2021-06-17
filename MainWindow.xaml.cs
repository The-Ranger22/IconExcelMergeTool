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
    public partial class MainWindow {
        private ExcelMerger em;
        public MainWindow()
        {
            InitializeComponent();
            em = new ExcelMerger();
            
        }

        private void _selectFiles(object o, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Multiselect = true;
            ofd.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if (ofd.ShowDialog() == true)
            {
                TBlock.Height = LGrid.RowDefinitions[1].ActualHeight;
                TBlock.Background = Brushes.Gainsboro;
                string s = "";
                foreach (string fname in ofd.FileNames)
                {
                    em.AddWorkbook(fname);
                    s += $"{fname}\n";
                }
                TBlock.Text = s;
            }

        }

        void _mergeFiles(object sender, RoutedEventArgs e)
        {
            
            SaveFileDialog sfd = new SaveFileDialog(); //declare and init save file dialog for use.
            sfd.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

            if (sfd.ShowDialog() == true) {
                em.Filename = sfd.FileName;
                em.Merge();
                em.Export();
            }
        }

        private void DisplaySelectedFiles()
        {
            
        }

        

        private void MainWindow_OnClosed(object sender, EventArgs e) {
            em.Quit();
        }
    }
}