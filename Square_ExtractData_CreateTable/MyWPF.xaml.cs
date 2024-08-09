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
using System.ComponentModel;
using System.Windows.Forms;
using Square_ExtractData_CreateTable.ViewModel;

namespace Square_ExtractData_CreateTable
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class MyWPF : Window, INotifyPropertyChanged
    {
        private ViewModelObject viewModelObject;

        public MyWPF(ViewModelObject view)
        {
            InitializeComponent();

            this.viewModelObject = view;
            this.DataContext = view;

            OnPropertyChanged(nameof(view));

            //this.Closed += MyWPF_Closed;

            //toolStripStatusLabel1.Text = viewModelObject.progressMessage;
            //toolStripProgressBar1.Value = viewModelObject.progressValue;

            //OnPropertyChanged(nameof(toolStripStatusLabel1));
            //OnPropertyChanged(nameof(toolStripProgressBar1));
        }

        private void MyWPF_Closed(object sender, EventArgs e)
        {
            //this.Close();
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
