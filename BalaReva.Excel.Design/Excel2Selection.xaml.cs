namespace BalaReva.Excel.Design
{
    using Microsoft.Win32;
    using System.Activities;
    using System.Activities.Presentation;
    using System.Activities.Presentation.Model;

    // Interaction logic for Excel2Selection.xaml
    public partial class Excel2Selection: ActivityDesigner
    {
        public Excel2Selection()
        {
            InitializeComponent();
        }

        private void btnFile_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            OpenFileDialog _openFileDialog = new OpenFileDialog();
            _openFileDialog.Title = "Open XLSX File";
            _openFileDialog.Filter = "Excel Files|*.xl*;*.xlsx;*.xlsm";

            if (_openFileDialog.ShowDialog() == true)
            {
                ModelProperty property = this.ModelItem.Properties["FilePath"];
                //property
                property.SetValue(new InArgument<string>(_openFileDialog.FileName));
            }
        }

        private void btnNewFilePath_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            OpenFileDialog _openFileDialog = new OpenFileDialog();
            _openFileDialog.Title = "Open XLSX File";
            _openFileDialog.Filter = "Excel Files|*.xl*;*.xlsx;*.xlsm";

            if (_openFileDialog.ShowDialog() == true)
            {
                ModelProperty property = this.ModelItem.Properties["NewFilePath"];
                //property
                property.SetValue(new InArgument<string>(_openFileDialog.FileName));
            }
        }
    }
}
