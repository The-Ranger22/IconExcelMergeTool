using System;
using System.Collections;
using System.Windows;
using System.Windows.Input;
using Microsoft.Win32;

namespace ExcelMerge {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow {
        private const string NO_KEY_SELECTED = "No Key Selected";
        private ExcelMerger em;

        public MainWindow() {
            InitializeComponent();
            em = new ExcelMerger();
        }

        private void _selectFiles(object o, EventArgs e) {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Multiselect = true;
            ofd.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if (ofd.ShowDialog() == true) {
                Mouse.OverrideCursor = Cursors.Wait;
                //SelectedFiles.Height = LGrid.RowDefinitions[1].ActualHeight;
                foreach (string fname in ofd.FileNames) {
                    em.AddWorkbook(fname);
                }
                SelectedFiles.ItemsSource = ofd.FileNames;
                em.Merge();
                ArrayList keyFields = (ArrayList) em.Header.Clone();
                keyFields.Add(NO_KEY_SELECTED);
                PrimaryKeyComBox.ItemsSource = keyFields;
                PrimaryKeyComBox.SelectedItem = NO_KEY_SELECTED;
                SecondaryKeyComBox.ItemsSource = keyFields;
                SecondaryKeyComBox.SelectedItem = NO_KEY_SELECTED;
                SumFields.ItemsSource = em.Header;
                Mouse.OverrideCursor = Cursors.Arrow;
            }
        }

        private void _mergeFiles(object sender, RoutedEventArgs e) {
            SaveFileDialog sfd = new SaveFileDialog(); //declare and init save file dialog for use.
            sfd.Filter = "Excel Files|*.xlsx";

            if (sfd.ShowDialog() == true) {
                em.Filename = sfd.FileName;
                
                try {
                    Mouse.OverrideCursor = Cursors.Wait;
                    if (!(CBox.IsChecked is null) && (bool) CBox.IsChecked && SumFields.SelectedItems.Count > 0) {
                        int[] selectedFieldIndices = new int[SumFields.SelectedItems.Count];
                        int primaryKey = (!PrimaryKeyComBox.SelectionBoxItem.Equals(NO_KEY_SELECTED))
                            ? em.Header.IndexOf(PrimaryKeyComBox.SelectionBoxItem)
                            : ExcelMerger.NULL_KEY;
                        int secondaryKey = (!SecondaryKeyComBox.SelectionBoxItem.Equals(NO_KEY_SELECTED))
                            ? em.Header.IndexOf(SecondaryKeyComBox.SelectionBoxItem)
                            : ExcelMerger.NULL_KEY;
                        int counter = 0;
                        foreach (var field in SumFields.SelectedItems) {
                            selectedFieldIndices[counter++] = em.Header.IndexOf(field);
                        }
                        em.Slim(primaryKey, secondaryKey, selectedFieldIndices);
                    }


                    em.Export();
                    Mouse.OverrideCursor = Cursors.Arrow;
                    MessageBox.Show("Merge complete!");
                    
                }
                catch (Exception exception) {
                    Console.WriteLine(exception);
                    throw;
                }
            }
        }

        


        private void MainWindow_OnClosed(object sender, EventArgs e) {
            em.Quit();
        }
    }
}