using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Forms;

namespace GetAllMetadataFromPdf
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, IMainWindow<string>
    {
        public MainWindow()
        {
            InitializeComponent();

            IMessageService messageService = new MessageService();

            MainPresenter mainPresenter = new MainPresenter(this, messageService);
        }

        public string PdfPath
        {
            get
            {
                return fldPdfPath.Text;
            }
        }

        public event EventHandler ImportMetaDataToExcelClick;

        private void btnChoosePdfClick(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "PDF file|*.pdf";
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                fldPdfPath.Text = dlg.FileName;
            }
        }

        private void btnInportMetaDataClick(object sender, RoutedEventArgs e)
        {
            if (ImportMetaDataToExcelClick != null) ImportMetaDataToExcelClick(this, EventArgs.Empty);
        }
    }
}
