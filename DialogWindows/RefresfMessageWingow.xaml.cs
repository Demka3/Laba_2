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
using System.Windows.Shapes;

namespace Laba2.DialogWindows
{
    /// <summary>
    /// Логика взаимодействия для RefresfMessageWingow.xaml
    /// </summary>
    public partial class RefresfMessageWingow : Window
    {        
        public RefresfMessageWingow(string refreshed, string added, string deleted)
        {
            InitializeComponent();
            messageTextBlock.Text += refreshed;
            messageTextBlock.Text += added;
            messageTextBlock.Text += deleted;            
        }

        private void okButton_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
        }
    }
}
