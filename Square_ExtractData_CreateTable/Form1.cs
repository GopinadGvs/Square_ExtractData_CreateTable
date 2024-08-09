using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Square_ExtractData_CreateTable.ViewModel;

namespace Square_ExtractData_CreateTable
{
    public partial class Form1 : Form, INotifyPropertyChanged
    {
        private ViewModelObject viewModelObject;

        public Form1(ViewModelObject view)
        {
            InitializeComponent();
            this.viewModelObject = view;
            toolStripStatusLabel1.Text = viewModelObject.progressMessage;
            toolStripProgressBar1.Value = viewModelObject.progressValue;

            OnPropertyChanged(nameof(toolStripStatusLabel1));
            OnPropertyChanged(nameof(toolStripProgressBar1));
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
