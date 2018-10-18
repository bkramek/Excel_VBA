using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelVBA
{
    class View
    {
        public void ShowView(string sourceDirectory)
        {
            
            var txtFiles = Directory.EnumerateFiles(sourceDirectory, "*.xls*", SearchOption.AllDirectories);
            Excel.Application app = new Excel.Application();
            app.Visible = true;
            foreach (string currentFile in txtFiles)
            {
                app.Workbooks.Open(currentFile);
            }
            
            
          

        }
    }
}
