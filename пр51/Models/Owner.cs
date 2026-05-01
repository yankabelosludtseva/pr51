using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace пр51.Models
{
    public class Owner
    {
        /// <summary> Имя
        public string FirstName { get; set; }
        /// <summary> Фамилия
        public string LastName { get; set; }
        /// <summary> Отчество
        public string SurName { get; set; }
        /// <summary> Номер квартиры
        public int NumberRoom { get; set; }
        /// <summary> Конструктор класса
        public Owner(string FirstName, string LastName, string SurName, int NumberRoom)
        {
            this.FirstName = FirstName;
            this.LastName = LastName;
            this.SurName = SurName;
            this.NumberRoom = NumberRoom;
        }
    }
}
