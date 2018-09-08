using Microsoft.Win32;
using System.Activities;
using System.Activities.Presentation;
using System.Activities.Presentation.Model;
using System.Windows;

namespace BalaReva.Excel.Design
{
    /// <summary>
    /// Interaction logic for LineChartDesigner.xaml
    /// </summary>
    public partial class LineChartDesigner : ActivityDesigner
    {
        public LineChartDesigner()
        {
            InitializeComponent();
        }

        private void btnFile_Click(object sender, RoutedEventArgs e)
        {
             OpenFileDialog _openFileDialog = new OpenFileDialog();
            _openFileDialog.Title = "Open XLSX File";
            _openFileDialog.Filter = "Excel Files|*.xl*;*.xlsx;*.xlsm";
            //_openFileDialog.InitialDirectory = @"C:\";

            if (_openFileDialog.ShowDialog() == true)
            {
                ModelProperty property = this.ModelItem.Properties["FilePath"];
                //property
                property.SetValue(new InArgument<string>(_openFileDialog.FileName));

            }
        }
    }
}
