

namespace BalaReva.Externals.Design
{
    using System.Activities;
    using System.Activities.Presentation;
    using System.Activities.Presentation.Model;
    using System.Windows;
    using System.Windows.Forms;

    // Interaction logic for UnZipFileDesign.xaml
    public partial class UnZipFileDesign : ActivityDesigner
    {
        public UnZipFileDesign()
        {
            InitializeComponent();
        }

        private void btnDirectory_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                ModelProperty modelProperty = this.ModelItem.Properties["ExtractFolderPath"];
                modelProperty.SetValue(new InArgument<string>(dialog.SelectedPath));
            }
        }

        private void btnFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog _openFileDialog = new OpenFileDialog();
            if (_openFileDialog.ShowDialog() == DialogResult.OK)
            {
                ModelProperty modelProperty = this.ModelItem.Properties["ZipFile"];
                modelProperty.SetValue(new InArgument<string>(_openFileDialog.FileName));
            }
        }
    }
}
