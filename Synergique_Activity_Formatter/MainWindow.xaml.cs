using System.Windows;
using System.Windows.Navigation;
using System;
using System.Collections.Generic;
using System.Text;
using Bytescout.Spreadsheet;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Media.Animation;
using Bytescout.Spreadsheet.Constants;
using Synergique_Activity_Formatter.Core;
using System.Text.Json;
using System.Windows.Controls;

namespace Synergique_Activity_Formatter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        private readonly FileManagement _fileManagement = new FileManagement();
        private readonly ExcelReader _excelReader = new ExcelReader();
        private readonly ExcelWriter _excelWriter = new ExcelWriter();
        private readonly JsonManager _jsonManager = new JsonManager();
        private readonly MonthRerouter _monthRerouter = new MonthRerouter();
        private List<Item> _recentDataItems = new List<Item>();
        private List<Item> _oldDataItems = new List<Item>();
        private int currentMonth = 0;

        public MainWindow()
        {
            InitializeComponent();
            _oldDataItems = _jsonManager.ReadData("oldItems.txt");
        }

        private void Browse_OnClick(object sender, RoutedEventArgs e)
        {
            var newSheetPath = _fileManagement.BrowseNewSheet();
            _recentDataItems = _excelReader.ReadData(_jsonManager.ReadData("oldItems.txt"), newSheetPath,_jsonManager,true);
        }

        private void Save_OnClick(object sender, RoutedEventArgs e)
        {
            var updatedRecentDataItems = _excelWriter.WriteOutput(currentMonth, _recentDataItems, _oldDataItems);
            if (updatedRecentDataItems != null)
            {
                _recentDataItems = updatedRecentDataItems;
            }
            _jsonManager.SerializeToJson(_recentDataItems, "savedItems.txt");
        }

        private void BrowseData2_OnClick(object sender, RoutedEventArgs e)
        {
            var oldDataPath = _fileManagement.BrowseOldData();
            _oldDataItems = _excelReader.ReadData(_oldDataItems, oldDataPath,_jsonManager,false);
            _jsonManager.SerializeToJson(_oldDataItems, "oldItems.txt");
        }

        private void Month_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var inputMonth = ((ComboBoxItem)(((ComboBox)sender).SelectedItem)).Content.ToString();
            currentMonth = _monthRerouter.RerouteMonth(inputMonth);
        }
    }
}