using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading.Tasks;
using pr51.Models;

namespace pr51.Context
{
    public class OwnerContext : DbContext
    {
        // Таблица владельцев
        public DbSet<Owner> Owners { get; set; }

        // Строка подключения к базе данных
        private static string connectionString = "Server=localhost;Database=DocumentsDB;User Id=root;Password=;";

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseMySql(connectionString, ServerVersion.AutoDetect(connectionString));
        }

        /// <summary>
        /// Получить всех владельцев
        /// </summary>
        public async Task<List<Owner>> GetAllOwnersAsync()
        {
            return await Owners.ToListAsync();
        }

        /// <summary>
        /// Получить владельца по номеру квартиры
        /// </summary>
        public async Task<Owner> GetOwnerByRoomAsync(int roomNumber)
        {
            return await Owners.FirstOrDefaultAsync(o => o.NumberRoom == roomNumber);
        }

        /// <summary>
        /// Получить владельцев по фамилии
        /// </summary>
        public async Task<List<Owner>> GetOwnersByLastNameAsync(string lastName)
        {
            return await Owners.Where(o => o.LastName == lastName).ToListAsync();
        }

        /// <summary>
        /// Сгенерировать отчёт по всем владельцам
        /// </summary>
        public async Task<string> GenerateReportAsync()
        {
            var owners = await GetAllOwnersAsync();

            var report = new System.Text.StringBuilder();
            report.AppendLine("=== ОТЧЁТ ПО ВЛАДЕЛЬЦАМ КВАРТИР ===\n");
            report.AppendLine($"Всего записей: {owners.Count}\n");
            report.AppendLine("№\tФИО\t\t\t\tКвартира");
            report.AppendLine(new string('-', 60));

            int index = 1;
            foreach (var owner in owners)
            {
                string fullName = $"{owner.LastName} {owner.FirstName} {owner.SurName}";
                report.AppendLine($"{index++}\t{fullName,-30}\t{owner.NumberRoom}");
            }

            report.AppendLine(new string('-', 60));
            report.AppendLine($"\nОтчёт сгенерирован: {DateTime.Now:dd.MM.yyyy HH:mm:ss}");

            return report.ToString();
        }

        /// <summary>
        /// Сгенерировать отчёт в формате JSON
        /// </summary>
        public async Task<string> GenerateJsonReportAsync()
        {
            var owners = await GetAllOwnersAsync();

            var jsonReport = new System.Text.StringBuilder();
            jsonReport.AppendLine("{");
            jsonReport.AppendLine("  \"reportTitle\": \"Отчёт по владельцам квартир\",");
            jsonReport.AppendLine($"  \"generatedAt\": \"{DateTime.Now:yyyy-MM-dd HH:mm:ss}\",");
            jsonReport.AppendLine($"  \"totalCount\": {owners.Count},");
            jsonReport.AppendLine("  \"owners\": [");

            for (int i = 0; i < owners.Count; i++)
            {
                var owner = owners[i];
                jsonReport.AppendLine("    {");
                jsonReport.AppendLine($"      \"firstName\": \"{owner.FirstName}\",");
                jsonReport.AppendLine($"      \"lastName\": \"{owner.LastName}\",");
                jsonReport.AppendLine($"      \"surName\": \"{owner.SurName}\",");
                jsonReport.AppendLine($"      \"numberRoom\": {owner.NumberRoom}");
                jsonReport.Append("    }");

                if (i < owners.Count - 1)
                    jsonReport.AppendLine(",");
                else
                    jsonReport.AppendLine();
            }

            jsonReport.AppendLine("  ]");
            jsonReport.AppendLine("}");

            return jsonReport.ToString();
        }

        /// <summary>
        /// Добавить нового владельца
        /// </summary>
        public async Task AddOwnerAsync(Owner owner)
        {
            await Owners.AddAsync(owner);
            await SaveChangesAsync();
        }

        /// <summary>
        /// Обновить данные владельца
        /// </summary>
        public async Task UpdateOwnerAsync(Owner owner)
        {
            Owners.Update(owner);
            await SaveChangesAsync();
        }

        /// <summary>
        /// Удалить владельца
        /// </summary>
        public async Task DeleteOwnerAsync(int roomNumber)
        {
            var owner = await GetOwnerByRoomAsync(roomNumber);
            if (owner != null)
            {
                Owners.Remove(owner);
                await SaveChangesAsync();
            }
        }
    }
}
