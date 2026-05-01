using PdfSharp.Drawing;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using пр51.Models;
using Word = Microsoft.Office.Interop.Word;

namespace пр51.Context
{
    /// <summary>
    /// Класс контекста владельца для работы с документами
    /// </summary>
    public class OwnerContext : Owner
    {
        /// <summary>
        /// Конструктор класса
        /// </summary>
        public OwnerContext(string FirstName, string LastName, string SurName, int NumberRoom) :
            base(FirstName, LastName, SurName, NumberRoom)
        { }

        /// <summary> Получение данных
        public static List<OwnerContext> AllOwners()
        {
            // Создаём новый список
            List<OwnerContext> allOwners = new List<OwnerContext>();
            // Добавляем данные в список
            allOwners.Add(new OwnerContext("Елена", "Иванова", "Петровна", 1));
            allOwners.Add(new OwnerContext("Алексей", "Смирнов", "Владимирович", 2));
            allOwners.Add(new OwnerContext("Анна", "Кузнецова", "Сергеевна", 3));
            allOwners.Add(new OwnerContext("Дмитрий", "Павлов", "Александрович", 3));
            allOwners.Add(new OwnerContext("Ольга", "Михайлова", "Ивановна", 4));
            allOwners.Add(new OwnerContext("Артем", "Козлов", "Олегович", 5));
            allOwners.Add(new OwnerContext("Наталья", "Соколова", "Викторовна", 6));
            allOwners.Add(new OwnerContext("Игорь", "Лебедев", "Андреевич", 6));
            allOwners.Add(new OwnerContext("Екатерина", "Федорова", "Дмитриевна", 7));
            allOwners.Add(new OwnerContext("Андрей", "Александров", "Игоревич", 7));
            allOwners.Add(new OwnerContext("Оксана", "Степанова", "Николаевна", 8));
            allOwners.Add(new OwnerContext("Сергей", "Никитин", "Васильевич", 9));
            allOwners.Add(new OwnerContext("Мария", "Ковалева", "Александровна", 10));
            allOwners.Add(new OwnerContext("Павел", "Фролов", "Михайлович", 11));
            allOwners.Add(new OwnerContext("Елена", "Белова", "Александровна", 12));
            allOwners.Add(new OwnerContext("Илья", "Поляков", "Данилович", 13));
            allOwners.Add(new OwnerContext("Анастасия", "Гаврилова", "Валерьевна", 14));
            allOwners.Add(new OwnerContext("Денис", "Орлов", "Владимирович", 15));
            allOwners.Add(new OwnerContext("Алина", "Киселева", "Сергеевна", 16));
            allOwners.Add(new OwnerContext("Артем", "Ткаченко", "Викторович", 16));
            allOwners.Add(new OwnerContext("Валерия", "Романова", "Павловна", 16));
            allOwners.Add(new OwnerContext("Александр", "Максимов", "Юрьевич", 17));
            allOwners.Add(new OwnerContext("Евгения", "Сидорова", "Игоревна", 17));
            allOwners.Add(new OwnerContext("Никита", "Антонов", "Алексеевич", 18));
            allOwners.Add(new OwnerContext("Юлия", "Дмитриева", "Владимировна", 19));
            // Возвращаем список обратно
            return allOwners;
        }

        /// <summary> Генерация отчёта
        public static void Report(string fileName)
        {
            // Создаём приложение
            Word.Application app = new Word.Application();
            // Создаём документ
            Word.Document doc = app.Documents.Add();
            // Создаём заголовок
            Word.Paragraph paraHeader = doc.Paragraphs.Add();
            // Указываем шрифт для заголовка
            paraHeader.Range.Font.Size = 16;
            // Задаём текст для заголовка
            paraHeader.Range.Text = "Список жильцов дома";
            // Указываем положение на странице
            paraHeader.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            // Убираем отступ
            paraHeader.Range.ParagraphFormat.SpaceAfter = 0;
            // Добавляем жирность
            paraHeader.Range.Font.Bold = 1;
            // Добавляем на документ
            paraHeader.Range.InsertParagraphAfter();

            // Создаём подзаголовок
            Word.Paragraph paraAddress = doc.Paragraphs.Add();
            // Указываем шрифт
            paraAddress.Range.Font.Size = 14;
            // Задаём текст
            paraAddress.Range.Text = "по адресу: г. Пермь, ул. Луначарского, д. 24";
            // Указываем отступ
            paraHeader.Range.ParagraphFormat.SpaceAfter = 20;
            // Убираем жирность
            paraHeader.Range.Font.Bold = 0;
            // Добавляем на документ
            paraAddress.Range.InsertParagraphAfter();

            // Создаём надпись
            Word.Paragraph paraCount = doc.Paragraphs.Add();
            // Указываем шрифт
            paraCount.Range.Font.Size = 14;
            // Задаём текст
            paraCount.Range.Text = $"Всего жильцов: {AllOwners().Count}";
            // Указываем положение на странице
            paraCount.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            // Указываем отступ
            paraHeader.Range.ParagraphFormat.SpaceAfter = 0;
            // Добавляем на документ
            paraCount.Range.InsertParagraphAfter();

            // Создаём таблицу
            Word.Paragraph tableParagraph = doc.Paragraphs.Add();
            // Добавляем на документ
            Word.Table paymentsTable = doc.Tables.Add(tableParagraph.Range, AllOwners().Count + 1, 4);
            // Указываем границы таблицы
            paymentsTable.Borders.InsideLineStyle = paymentsTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            // Указываем положение таблицы
            paymentsTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            // Создаём заголовки в таблице
            Cell("№", paymentsTable.Cell(1, 1).Range);
            Cell("Фамилия", paymentsTable.Cell(1, 2).Range);
            Cell("Имя", paymentsTable.Cell(1, 3).Range);
            Cell("Отчество", paymentsTable.Cell(1, 4).Range);

            // Перебираем жильцов
            for (int i = 0; i < AllOwners().Count; i++)
            {
                // Получаем жильца
                OwnerContext owner = AllOwners()[i];
                // Добавляем границы о жильцах
                Cell((i + 1).ToString(), paymentsTable.Cell(1 + 1 + i, 1).Range);
                Cell(owner.LastName, paymentsTable.Cell(1 + 1 + i, 2).Range, Word.WdParagraphAlignment.wdAlignParagraphLeft);
                Cell(owner.FirstName, paymentsTable.Cell(1 + 1 + i, 3).Range, Word.WdParagraphAlignment.wdAlignParagraphLeft);
                Cell(owner.SurName, paymentsTable.Cell(1 + 1 + i, 4).Range, Word.WdParagraphAlignment.wdAlignParagraphLeft);
            }
            // Сохраняем документ
            doc.SaveAs2(fileName);
            // Закрываем документ
            doc.Close();
            // Закрываем приложение
            app.Quit();
        }

        /// <summary>
        /// Добавление текста в ячейку
        /// </summary>
        /// <param name="Text">Текст в ячейке</param>
        /// <param name="Cell">Ячейка</param>
        /// <param name="Alignment">Положение в ячейке</param>
        public static void Cell(string Text, Word.Range Cell,
            Word.WdParagraphAlignment Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter)
        {
            // Указываем текст
            Cell.Text = Text;
            // Указываем положение текста в ячейке
            Cell.ParagraphFormat.Alignment = Alignment;
        }

        /// <summary>
        /// Генерация отчёта PDF
        /// </summary>
        /// <param name="fileName">Наименование файла</param>
        public static void ReportPDF(string fileName)
        {
            // Создаём документ PDF
            PdfDocument document = new PdfDocument();
            // Указываем заголовок документа
            document.Info.Title = "Отчёт по жильцам дома";
            // Добавляем страницу в документ
            PdfPage page = document.AddPage();
            // Получаем графику для созданной страницы страницы
            XGraphics gfx = XGraphics.FromPdfPage(page);
            // Присваиваем отступ сверху
            int MarginTop = 20;
            // Присваиваем отступ слева
            int MarginLeft = 50;
            // Задаём используемые шрифты
            XFont fontHeader = new XFont("Arial", 16);
            XFont font = new XFont("Arial", 12);
            // Указываем заголовок
            gfx.DrawString("Список жильцов дома", fontHeader, XBrushes.Black,
                new XRect(0, MarginTop, page.Width, 15),
                XStringFormats.Center);

            // Указываем подзаголовок
            gfx.DrawString("по адресу: г. Пермь, ул. Луначарского, д. 24", font, XBrushes.Black,
                new XRect(0, MarginTop + 30, page.Width, 10),
                XStringFormats.Center);

            // Указываем текст
            gfx.DrawString("Всего жильцов: " + AllOwners().Count, font, XBrushes.Black,
                new XRect(MarginLeft, MarginTop + 70, page.Width, 10),
                XStringFormats.CenterLeft);

            // Расчитываем ширину ячейки в таблице
            int Width = (Convert.ToInt32(page.Width.Value) - MarginLeft * 2 - 30) / 4;
            // Рисуем квадраты, которые будут обозначать ячейки таблицы
            gfx.DrawRectangle(new XSolidBrush(XColors.LightGray), MarginLeft, MarginTop + 100, Width, 20);
            gfx.DrawRectangle(new XSolidBrush(XColors.LightGray), MarginLeft + Width + 10, MarginTop + 100, Width, 20);
            gfx.DrawRectangle(new XSolidBrush(XColors.LightGray), MarginLeft + (Width + 10) * 2, MarginTop + 100, Width, 20);
            gfx.DrawRectangle(new XSolidBrush(XColors.LightGray), MarginLeft + (Width + 10) * 3, MarginTop + 100, Width, 20);
            // Вставляем текст на места ячеек
            gfx.DrawString("№" + AllOwners().Count, font, XBrushes.Black,
                new XRect(MarginLeft, MarginTop + 100, Width, 20),
                XStringFormats.Center);

            // Вставляем текст на места ячеек
            gfx.DrawString("Фамилия", font, XBrushes.Black,
                new XRect(MarginLeft + Width + 10, MarginTop + 100, Width, 20),
                XStringFormats.Center);

            // Вставляем текст на места ячеек
            gfx.DrawString("Имя", font, XBrushes.Black,
                new XRect(MarginLeft + (Width + 10) * 2, MarginTop + 100, Width, 20),
                XStringFormats.Center);

            // Вставляем текст на места ячеек
            gfx.DrawString("Отчество", font, XBrushes.Black,
                new XRect(MarginLeft + (Width + 10) * 3, MarginTop + 100, Width, 20),
                XStringFormats.Center);

            // Перебираем жильцов квартир
            for (int i = 0; i < AllOwners().Count; i++)
            {
                // Рисуем квадраты, которые будут обозначать ячейки таблицы
                gfx.DrawRectangle(new XSolidBrush(XColors.LightGray), MarginLeft, MarginTop + 100 + 25 * (i + 1), Width, 20);
                gfx.DrawRectangle(new XSolidBrush(XColors.LightGray), MarginLeft + Width + 10, MarginTop + 100 + 25 * (i + 1), Width, 20);
                gfx.DrawRectangle(new XSolidBrush(XColors.LightGray), MarginLeft + (Width + 10) * 2, MarginTop + 100 + 25 * (i + 1), Width, 20);
                gfx.DrawRectangle(new XSolidBrush(XColors.LightGray), MarginLeft + (Width + 10) * 3, MarginTop + 100 + 25 * (i + 1), Width, 20);
                // Вставляем текст на места ячеек
                gfx.DrawString((i + 1).ToString(), font, XBrushes.Black,
                    new XRect(MarginLeft, MarginTop + 100 + 25 * (i + 1), Width, 20),
                    XStringFormats.Center);

                // Вставляем текст на места ячеек
                gfx.DrawString(AllOwners()[i].LastName, font, XBrushes.Black,
                    new XRect(MarginLeft + Width + 10, MarginTop + 100 + 25 * (i + 1), Width, 20),
                    XStringFormats.Center);

                // Вставляем текст на места ячеек
                gfx.DrawString(AllOwners()[i].FirstName, font, XBrushes.Black,
                    new XRect(MarginLeft + (Width + 10) * 2, MarginTop + 100 + 25 * (i + 1), Width, 20),
                    XStringFormats.Center);

                // Вставляем текст на места ячеек
                gfx.DrawString(AllOwners()[i].SurName, font, XBrushes.Black,
                    new XRect(MarginLeft + (Width + 10) * 3, MarginTop + 100 + 25 * (i + 1), Width, 20),
                    XStringFormats.Center);
            }

            // Сохраняем документ
            document.Save(fileName);
            // Открываем документ
            Process.Start(fileName);
        }
    }
}
