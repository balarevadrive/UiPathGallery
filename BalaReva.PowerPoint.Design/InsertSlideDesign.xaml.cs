using System;
using System.Activities;
using System.Activities.Presentation;
using System.Activities.Presentation.Model;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace BalaReva.PowerPoint.Design
{
    /// <summary>
    /// Interaction logic for InsertSlideDesign.xaml
    /// </summary>
    public partial class InsertSlideDesign : ActivityDesigner
    {
        public InsertSlideDesign()
        {
            InitializeComponent();
        }

        private void btnFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog _openFileDialog = new OpenFileDialog();
            _openFileDialog.Title = "Open PowerPiont File";
            _openFileDialog.Filter = "All PowerPoint Presentations|*.pp*";

            if (_openFileDialog.ShowDialog() ==  DialogResult.OK)
            {
                ModelProperty property = this.ModelItem.Properties["FilePath"];
                //property
                property.SetValue(new InArgument<string>(_openFileDialog.FileName));
            }
        }
    }
}
