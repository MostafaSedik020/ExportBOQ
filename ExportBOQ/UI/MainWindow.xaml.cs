using System;
using System.Collections.Generic;
using System.Data.Common;
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
using System.Xml.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using ExportBOQ.Excel;

namespace ExportBOQ.UI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        Document doc;
        public MainWindow(Document doc)
        {
            InitializeComponent();
            this.doc = doc;
        }


        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

        }

        private void ExportBtn_Click(object sender, RoutedEventArgs e)
        {
            // Define an array of CheckBox controls
            CheckBox[] elementCheckBoxes = new CheckBox[]
            {
                 cbFoundation, cbColumns, cbRwall, cbSwall, cbFloors,
                 cbSog, cbStairs, cbTanks, cbStand, cbBitPaint, cbPolySheet, cbMembrane
            };

            // Create an array to hold the selection states
            int[] selectionStates = new int[elementCheckBoxes.Length];

            // Populate the selectionStates array based on the CheckBox states
            for (int i = 0; i < elementCheckBoxes.Length; i++)
            {
                selectionStates[i] = elementCheckBoxes[i].IsChecked == true ? 1 : 0;
            }

            // Call the method to export data to Excel
            ExportDataToExcel.SendColumnsDataToExcel(doc, "dsd", @"E:\", selectionStates);
        }
        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                DragMove();
            }
        }
        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void ListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void cbSelectAll_Checked(object sender, RoutedEventArgs e)
        {
            bool newVal = (cbSelectAll.IsChecked == true);
           
            cbFoundation.IsChecked = newVal;
            cbColumns.IsChecked = newVal;
            cbRwall.IsChecked = newVal;
            cbSwall.IsChecked = newVal;
            cbStairs.IsChecked = newVal;
            cbSog.IsChecked = newVal;
            cbTanks.IsChecked = newVal;
            cbStand.IsChecked = newVal;
            cbBitPaint.IsChecked = newVal;
            cbPolySheet.IsChecked = newVal;
            cbMembrane.IsChecked = newVal;
            cbFloors.IsChecked = newVal;

        }
        private void cbOtherCheckedChanged(object sender, RoutedEventArgs e)
        {
           cbSelectAll.IsChecked = null;

            if ((cbFoundation.IsChecked == true) && (cbColumns.IsChecked == true) && (cbRwall.IsChecked == true)
                && (cbSwall.IsChecked == true)  && (cbStairs.IsChecked == true)
                && ( cbSog.IsChecked == true) && (cbTanks.IsChecked == true) && (cbStand.IsChecked == true)
                && (cbBitPaint.IsChecked == true) && (cbPolySheet.IsChecked == true) && (cbMembrane.IsChecked == true)
                && (cbFloors.IsChecked == true))
                {
                    cbSelectAll.IsChecked = true;
                }
            if ((cbFoundation.IsChecked == false ) && (cbColumns.IsChecked == false ) && (cbRwall.IsChecked == false )
                && (cbSwall.IsChecked ==  false ) && (cbStairs.IsChecked == false )
                && (cbSog.IsChecked == false ) && (cbTanks.IsChecked == false ) && (cbStand.IsChecked == false )
                && (cbBitPaint.IsChecked == false ) && (cbPolySheet.IsChecked == false ) && (cbMembrane.IsChecked == false )
                && ( cbFloors.IsChecked == false ))
            {
                cbSelectAll.IsChecked = false ;
            }
        }

        private void ListBox_SelectionChanged_1(object sender, SelectionChangedEventArgs e)
        {

        }
    }
}
