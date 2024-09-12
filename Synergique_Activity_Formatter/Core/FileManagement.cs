namespace Synergique_Activity_Formatter.Core
{
    public class FileManagement
    {
        public string BrowseNewSheet()
        {
            // Configure open file dialog box
            var dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.FileName = "Document"; // Default file name
            dialog.DefaultExt = ".xlsx"; // Default file extension
            dialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"; // Filter files by extension

            // Show open file dialog box
            bool? result = dialog.ShowDialog();

            // Process open file dialog box results
            if (result == true)
            {
                return  dialog.FileName;
            }

            return null;
        }

        public string BrowseOldData()
        {
            var dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.FileName = "Document"; // Default file name
            dialog.DefaultExt = ".xlsx"; // Default file extension
            dialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"; // Filter files by extension

// Show open file dialog box
            bool? result = dialog.ShowDialog();

// Process open file dialog box results
            if (result == true)
            {
                // Open document
                return dialog.FileName;
            }

            return null;
        }
    }
}