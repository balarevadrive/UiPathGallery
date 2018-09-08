namespace BalaReva.PowerPoint.Design
{
    using System.Activities;
    using System.Activities.Presentation;
    using System.Activities.Presentation.Model;
    using System.Windows;
    using System.Windows.Forms;

    /// <summary>
    /// Interaction logic for InsertPictureDesign.xaml
    /// </summary>
    public partial class InsertPictureDesign : ActivityDesigner
    {
        public InsertPictureDesign()
        {
            InitializeComponent();
        }

        private void btnFile_Click(object sender, RoutedEventArgs e)
        {
            
            OpenFileDialog _openFileDialog = new OpenFileDialog();
            _openFileDialog.Title = "Open PowerPiont File";
            _openFileDialog.Filter = "All PowerPoint Presentations|*.pp*";


            if (_openFileDialog.ShowDialog() ==  DialogResult.OK )
            {
                ModelProperty property = this.ModelItem.Properties["FilePath"];
                //property
                property.SetValue(new InArgument<string>(_openFileDialog.FileName));
            }
        }

        private void btnImage_Click(object sender, RoutedEventArgs e)
        {
            string strfilter = string.Empty;
            OpenFileDialog _openFileDialog = new OpenFileDialog();
            _openFileDialog.Title = "Open Image File";

            strfilter = "All Pictures(*.emf;*.jpg;*.jpeg;*.jfif;*.png;*.bmp;*.dib;*.rle;*.bmz;*.gif;*.gfa;*.emz;*.wmz;*.pcz;*.tif;*.cgm;*.eps;*.pct;*.pict;*.wpg;)";
            strfilter = "|*.emf;*.jpg;*.jpeg;*.jfif;*.png;*.bmp;*.dib;*.rle;*.bmz;*.gif;*.gfa;*.emz;*.wmz;*.pcz;*.tif;*.cgm;*.eps;*.pct;*.pict;*.wpg";

            _openFileDialog.Filter = strfilter;

            if (_openFileDialog.ShowDialog() == DialogResult.OK)
            {
                ModelProperty property = this.ModelItem.Properties["ImagePath"];
                //property
                property.SetValue(new InArgument<string>(_openFileDialog.FileName));
            }

        }
    }
}
