using Microsoft.Win32;
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
using пр51.Context;

namespace пр51
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Report(object sender, RoutedEventArgs e)
        {
            // Создаём файловый диалог
            SaveFileDialog sfd = new SaveFileDialog();
            // Указываем формат файла
            sfd.Filter = "Word Files (*.docx)|*.docx";
            // Открываем файловый диалог
            sfd.ShowDialog();
            // Если было выбрано имя
            if (sfd.FileName != "")
            {
                // Создаём отчёт
                OwnerContext.Report(sfd.FileName);
            }
        }
        public void LoadRooms()
        {
            for (int i = 1; i < 20; i++)
                Parent.Children.Add(new Elements.Room(i));
        }
    }
}
