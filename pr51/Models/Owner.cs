using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace pr51.Models
{
    public class Owner
    {
        /// <summary> Имя</summary>
        public string FirstName { get; set; }

        /// <summary> Фамилия</summary>
        public string LastName { get; set; }

        /// <summary> Отчество</summary>
        public string SurName { get; set; }

        /// <summary> Номер квартиры</summary>
        public int NumberRoom { get; set; }

        /// <summary> Конструктор класса</summary>
        public Owner(string FirstName, string LastName, string SurName, int NumberRoom)
        {
            this.FirstName = FirstName;
            this.LastName = LastName;
            this.SurName = SurName;
            this.NumberRoom = NumberRoom;
        }
    }
}
