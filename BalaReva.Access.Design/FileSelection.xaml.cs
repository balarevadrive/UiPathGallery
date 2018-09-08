namespace BalaReva.Access.Design
{
    using Microsoft.Win32;
    using System.Activities;
    using System.Activities.Presentation;
    using System.Activities.Presentation.Model;
    using System.Windows;

    // Interaction logic for FileSelection.xaml
    public partial class FileSelection : ActivityDesigner
    {
        public FileSelection()
        {
            InitializeComponent();
        }

        private void btnFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog _openFileDialog = new OpenFileDialog();
            _openFileDialog.Title = "Open Access File";
            _openFileDialog.Filter = "Microsoft Office Access|*.accdb;*.mdb;*.adp;*.mda;*.accda;*.mde;*.accde;*.ade";

            if (_openFileDialog.ShowDialog() == true)
            {
                ModelProperty property = this.ModelItem.Properties["FilePath"];
                //property
                property.SetValue(new InArgument<string>(_openFileDialog.FileName));
            }

        }
    }
}
