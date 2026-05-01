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

namespace пр51.Elements
{
    /// <summary>
    /// Логика взаимодействия для Room.xaml
    /// </summary>
    public partial class Room : UserControl
    {
        // Ссылка: 1
        public Room(int Room)
        {
            InitializeComponent();
            // Указываем номер квартиры
            NameRoom.Content = "Квартира №" + Room;
            // Вызываем загрузку данных
            LoadOwner(Room);
        }

        /// <summary> Загрузка данных
        // Ссылка: 1
        public void LoadOwner(int Room)
        {
            // Получаем жильцов квартиры
            List<OwnerContext> roomOwners = OwnerContext.AllOwners().FindAll(x => x.NumberRoom == Room);
            // Добавляем элементы в stack panel
            foreach (OwnerContext roomOwner in roomOwners)
            {
                Parent.Children.Add(new Elements.Owner(roomOwner));
            }
        }
    }
}
