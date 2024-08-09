using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Threading;

namespace Square_ExtractData_CreateTable.ViewModel
{
    public class ViewModelObject : INotifyPropertyChanged
    {
        private int _progressValue;
        private string _progressMessage = "";

        //Thread cadThread1;

        public event PropertyChangedEventHandler PropertyChanged;

        public int progressValue
        {
            get => _progressValue;
            set
            {
                _progressValue = value;
                OnPropertyChanged(nameof(progressValue));
            }
        }

        public string progressMessage
        {
            get => _progressMessage;
            set
            {
                _progressMessage = value;
                OnPropertyChanged(nameof(progressMessage));
            }
        }

        public void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        //public Action ShowUI;

        //public void ShowUIInCAD()
        //{
        //    //Form1 frm = new Form1();
        //    //frm.toolStripProgressBar1.Value = progressValue;
        //    //frm.ShowDialog();

        //    if (ShowUI != null)
        //    {
        //        ShowUI.Invoke();
        //    }
        //}

        //public ViewModel()
        //{
        //    //cadThread1 = new Thread(ShowUIInCAD);
        //    //cadThread1.Start();
        //}
    }
}
