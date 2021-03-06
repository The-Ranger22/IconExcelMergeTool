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
        #region Private Fields
        /// <summary>
        /// The default string constant that is displayed when no keys have been selected by the user.
        /// </summary>
        private const string NO_KEY_SELECTED = "No Key Selected";
        /// <summary>
        /// Used to delimit the string given by the ignorable key field in the view
        /// </summary>
        private const char DELIMITER = ';';
        /// <summary>
        /// Used to manipulate excel files.
        /// </summary>
        private ExcelMerger _em;
        /// <summary>
        /// Used to prevent any merge attempt before files have been selected.
        /// </summary>
        private bool _filesAreSelected = false;
        #endregion
        public MainWindow() {
            InitializeComponent();
            _em = new ExcelMerger();
        }
        /// <summary>
        /// Selects files and reads them into the ExcelMerger object's internal data structure for processing.
        /// </summary>
        /// <param name="o"></param>
        /// <param name="e"></param>
        private void _selectFiles(object o, EventArgs e) {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Multiselect = true;
            ofd.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if (ofd.ShowDialog() == true) {
                _resetView();
                _filesAreSelected = true;
                Mouse.OverrideCursor = Cursors.Wait;
                foreach (string fname in ofd.FileNames) {
                    _em.AddWorkbook(fname);
                }
                SelectedFiles.ItemsSource = ofd.FileNames;
                _em.Merge();
                ArrayList keyFields = (ArrayList) _em.Header.Clone();
                keyFields.Add(NO_KEY_SELECTED);
                PrimaryKeyComBox.ItemsSource = keyFields;
                PrimaryKeyComBox.SelectedItem = NO_KEY_SELECTED;
                SecondaryKeyComBox.ItemsSource = keyFields;
                SecondaryKeyComBox.SelectedItem = NO_KEY_SELECTED;
                SumFields.ItemsSource = _em.Header;
                Mouse.OverrideCursor = Cursors.Arrow;
            }
        }
        /// <summary>
        /// Merges the selected files together based on arguments provided by the user through the view.
        /// </summary>
        /// <remarks>
        /// The arguments the user can provide via the UI are thusly:
        ///     Primary Key:
        ///         The primary key is the first field that is matched against during the slim operation.
        ///     Secondary Key:
        ///         The secondary key is a redundancy that is matched against if, and only if, the primary key produces
        ///         a null or ignorable result.
        ///     Summation Fields:
        ///         The fields to be summed during the slim operation.
        ///     Ignorable Key Values:
        ///         A string containing characters/phrases/words that are deemed ignorable that are separated by semicolon.
        ///     Slim Results:
        ///         Enables the slim operation during the merge process.
        /// </remarks>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void _mergeFiles(object sender, RoutedEventArgs e) {
            SaveFileDialog sfd = new SaveFileDialog(); //declare and init save file dialog for use.
            sfd.Filter = "Excel Files|*.xlsx";
            if (!_filesAreSelected) {
                MessageBox.Show("No files selected!");
                return;
            }
            if (sfd.ShowDialog() == true && _filesAreSelected) {
                _em.Filename = sfd.FileName;

                try {
                    Mouse.OverrideCursor = Cursors.Wait;
                    if (!(CBox.IsChecked is null) && (bool) CBox.IsChecked && SumFields.SelectedItems.Count > 0) {
                        int[] selectedFieldIndices = new int[SumFields.SelectedItems.Count];
                        int primaryKey = (!PrimaryKeyComBox.SelectionBoxItem.Equals(NO_KEY_SELECTED))
                            ? _em.Header.IndexOf(PrimaryKeyComBox.SelectionBoxItem)
                            : ExcelMerger.NULL_KEY;
                        int secondaryKey = (!SecondaryKeyComBox.SelectionBoxItem.Equals(NO_KEY_SELECTED))
                            ? _em.Header.IndexOf(SecondaryKeyComBox.SelectionBoxItem)
                            : ExcelMerger.NULL_KEY;
                        int counter = 0;
                        foreach (var field in SumFields.SelectedItems) {
                            selectedFieldIndices[counter++] = _em.Header.IndexOf(field);
                        }
                        if (!IgnorableKeyValues.Text.Equals("")) {
                            _em.ParseIgnorables(IgnorableKeyValues.Text, DELIMITER);
                        }
                        _em.Slim(primaryKey, secondaryKey, selectedFieldIndices);
                    }


                    _em.Export();
                    Mouse.OverrideCursor = Cursors.Arrow;
                    MessageBox.Show("Merge complete!");
                    _em.Quit();

                }
                catch (Exception exception) {
                    MessageBox.Show($"Something unexpected has occurred during the merging process: {exception}");
                    Console.WriteLine(exception);
                    _em.Quit();
                }
                finally {
                    _resetView();
                }
            }
        }
        /// <summary>
        /// Resets both the view elements and the ExcelMerger object to their default states.
        /// </summary>
        /// <remarks>
        /// While the method will reset all other control elements in the view, it will leave IgnorableKeyValues alone
        /// so that the user need not reenter the same set of ignorables over and over again.
        /// </remarks>
        private void _resetView() {
            _em.Quit();
            _em = new ExcelMerger();
            SelectedFiles.ItemsSource = null;
            PrimaryKeyComBox.ItemsSource = null;
            SecondaryKeyComBox.ItemsSource = null;
            SumFields.ItemsSource = null;
            CBox.IsChecked = false;
            _filesAreSelected = false;
        }


        private void MainWindow_OnClosed(object sender, EventArgs e) {
            _em.Quit();
        }
    }
}