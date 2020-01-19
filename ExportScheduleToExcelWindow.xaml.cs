using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using Autodesk.Revit.DB;
using Microsoft.WindowsAPICodePack.Dialogs;
using OfficeOpenXml;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;

namespace QApps
{
    /// <summary>
    /// Interaction logic for Window.xaml
    /// </summary>
    public partial class ExportScheduleToExcelWindow
    {
        private ExportScheduleToExcelViewModel _viewModel;

        public ExportScheduleToExcelWindow(ExportScheduleToExcelViewModel viewModel)
        {
            InitializeComponent();
            _viewModel = viewModel;
            DataContext = _viewModel;
        }

        private void btnOk_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_viewModel.ExportExcelFolderPath))
                _viewModel.ExportExcelFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            #region Code

            List<ViewScheduleExtension> viewScheduleExport = new List<ViewScheduleExtension>();
            foreach (ViewScheduleExtension viewSchedule in _viewModel.AllViewSchedules)
            {
                if (viewSchedule.IsExport) viewScheduleExport.Add(viewSchedule);
            }

            if (viewScheduleExport.Count == 0)
            {
                MessageBox.Show("Please select the Schedules which will be exported!");
                return;
            }

            if (_viewModel.IsMultiWorkbooks)
            {
                foreach (var vse in viewScheduleExport)
                {
                    ViewSchedule selectedSchedule = vse.ViewSchedule;
                    string filePath = Path.Combine(_viewModel.ExportExcelFolderPath, selectedSchedule.Name + ".xlsx");

                    try
                    {
                        string filePathTxt = Path.Combine(_viewModel.ExportExcelFolderPath,
                            string.Concat(selectedSchedule.Name, ".txt"));

                        List<string> scheduleExportFieldName = selectedSchedule.GetNameOfFields();

                        ViewScheduleExportOptions viewScheduleExportOption = new ViewScheduleExportOptions();
                        selectedSchedule.Export(_viewModel.ExportExcelFolderPath, Path.GetFileName(filePathTxt),
                            viewScheduleExportOption);

                        DataTable dataTable = ExcelUtils.ReadFile(filePathTxt, scheduleExportFieldName.Count);

                        FileInfo existingFile = new FileInfo(filePath);
                        using (var package = new ExcelPackage(existingFile))
                        {
                            ExcelUtils.ExportDataTableToExcel(package, dataTable,
                                selectedSchedule.Name, filePath);
                        }

                        File.Delete(filePathTxt);
                    }
                    catch (Exception)
                    {
                    }
                }
            }
            else
            {
                string title = _viewModel.Doc.Title;
                title = title.Substring(0, title.Length - 4);
                string filePath = Path.Combine(_viewModel.ExportExcelFolderPath,
                    string.Concat(title, ".xlsx"));
                FileInfo existingFile = new FileInfo(filePath);
                using (var package = new ExcelPackage(existingFile))
                {
                    foreach (var vse in viewScheduleExport)
                    {
                        ViewSchedule selectedSchedule = vse.ViewSchedule;
                        try
                        {
                            string filePathTxt = Path.Combine(_viewModel.ExportExcelFolderPath,
                                string.Concat(selectedSchedule.Name, ".txt"));

                            List<string> scheduleExportFieldName = selectedSchedule.GetNameOfFields();

                            ViewScheduleExportOptions viewScheduleExportOption = new ViewScheduleExportOptions();
                            selectedSchedule.Export(_viewModel.ExportExcelFolderPath, Path.GetFileName(filePathTxt),
                                viewScheduleExportOption);

                            DataTable dataTable = ExcelUtils.ReadFile(filePathTxt, scheduleExportFieldName.Count);

                            ExcelUtils.ExportDataTableToExcel(package,dataTable,
                                selectedSchedule.Name, filePath);

                            File.Delete(filePathTxt);
                        }
                        catch (Exception)
                        {
                        }
                    }
                }

            }

            if (MessageBox.Show("Do you want to open folder that store exported files?", "Open folder",
                    MessageBoxButton.YesNo)
                    == MessageBoxResult.Yes)
            {
                try
                {
                    Process.Start(_viewModel.ExportExcelFolderPath);
                }
                catch (Exception)
                {
                }
            }

            #endregion

            DialogResult = true;
            Close();
        }


        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void TextBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            string initialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments);
            dialog.InitialDirectory = initialDirectory;
            dialog.IsFolderPicker = true;
            dialog.RestoreDirectory = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                _viewModel.ExportExcelFolderPath = dialog.FileName;
            }
        }

        private void IsExportClick(object sender, RoutedEventArgs e)
        {
            ViewScheduleExtension first = _viewModel.SelectedViewSchedules
                .FirstOrDefault();
            bool selected = first.IsExport;
            foreach (var vs in _viewModel.SelectedViewSchedules)
                vs.IsExport = !selected;

        }
    }
}
