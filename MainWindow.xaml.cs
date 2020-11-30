using System;
using System.Collections.Generic;
using System.Net;
using System.Windows;
using System.Windows.Forms;
using MessageBox = System.Windows.MessageBox;
using System.IO;
using Laba2.DialogWindows;
using System.Threading;

namespace Laba2
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>    

    public partial class MainWindow : Window
    {
        private static Thread autoRefresh = new Thread(AutoRefresh);
        public MainWindow()
        {
            InitializeComponent();
            StartProgram();
            autoRefresh.Start();
        }

        //Обновляем раз в час
        private static void AutoRefresh()
        {
            while (true)
            {
                try
                {
                    Thread.Sleep(3600 * 1000);
                    Refresh();
                }
                catch (Exception)
                {

                }

            }
        }
        private static void StartProgram()
        {
            //Проверка на существование файла с таким же именем
            bool isFileExist = File.Exists(@"thrlist.xlsx");

            //Проверка существования локальной базы
            if (!File.Exists(@"LocalData.xlsx"))
            {
                //Если не существует, спрашиваем, надо ли скачать
                LoadConfirmWindow loadConfirmWindow = new LoadConfirmWindow();
                if (loadConfirmWindow.ShowDialog() == true)
                {
                    //Скачиваем файл, если нет интернета и пользователь отказался повторить скачивание - выходим
                    if (!Download())
                    {
                        return;
                    }
                }
                //На нет и суда нет
                else
                {
                    return;
                }
            }
            //Есть, ну и хорошо
            else
            {
                return;
            }

            CreateLocalData();

            //Удаляем, если не было файла с таким именем
            if (!isFileExist)
            {
                File.Delete("thrlist.xlsx");
            }

        }

        private static void CreateLocalData()
        {
            //Создаем локальнцю базу            
            Excel sourceFile = new Excel();
            Excel localFile = new Excel();
            sourceFile.FileOpen(@"thrlist.xlsx");

            //Перезаписываем таблицу, пока есть не пустые строки
            int j = 2;
            while (true)
            {
                string[] rowData = new string[8];
                try
                {
                    for (int i = 0; i < 8; i++)
                    {
                        rowData[i] = (sourceFile.Rows[j][i]);
                    }
                }
                catch (ArgumentOutOfRangeException)
                {
                    break;
                }
                j++;
                localFile.AddRow(rowData[0], rowData[1], rowData[2], rowData[3], rowData[4], rowData[5], rowData[6], rowData[7]);
            }
            //Записываем локальную базу в файл            
            localFile.FileSave(@"LocalData.xlsx");
        }

        private static bool Download()
        {
            WebClient webClient = new WebClient();
            while (true)
            {
                //Ловим ошибки соединения    
                try
                {
                    webClient.DownloadFile(@"https://bdu.fstec.ru/files/documents/thrlist.xlsx", @"thrlist.xlsx");
                    //Если скачался - выходим
                    if (File.Exists(@"thrlist.xlsx"))
                    {
                        return true;
                    }
                }
                catch (WebException)
                {
                    //Если произошла ошибка - спрашиваем о повторном скачивании
                    RepeatDownloadWindow repeatDownloadWindow = new RepeatDownloadWindow();
                    //Если пользователь отказался - выходим
                    if (repeatDownloadWindow.ShowDialog() == false)
                    {
                        return false;
                    }
                }
            }

        }

        private static void Refresh()
        {
            //Проверка на существование файла с таким же именем
            bool isFileExist = File.Exists(@"thrlist.xlsx");
            //Скачиваем файл, если нет интернета и пользователь отказался повторить скачивание - выходим
            if (!Download())
            {
                return;
            }

            Excel sourceFile = new Excel();
            sourceFile.FileOpen(@"thrlist.xlsx");
            //Проверяем существование локальной базы
            Excel localFile;
            try
            {
                localFile = DoesFileExist();
            }
            catch (IOException)
            {
                MessageBox.Show("Пожалуйста, закройте файл локальной базы.");
                return;
            }
            if (localFile == null)
            {
                //Если ее нет, делаем новую                
                CreateLocalData();
                MessageBox.Show("Локальная база создана!");
                //Удаляем, если не было файла с таким именем
                if (!isFileExist)
                {
                    File.Delete("thrlist.xlsx");
                }
                return;
            }

            //Переменные для хранения измененй
            List<string> id = new List<string>();
            List<string> idAdded = new List<string>();
            List<string> idDeleted = new List<string>();
            List<string> itWas = new List<string>();
            List<string> itBecame = new List<string>();
            int changedRowsCount = 0;

            //Получаем данные из скачаного и локального файлов
            int RowCountInSource = 2;
            int RowCountInLocal = 0;
            List<List<string>> sourceData = new List<List<string>>();
            List<List<string>> localData = new List<List<string>>();
            while (true)
            {
                List<string> sourceRow = new List<string>();
                try
                {
                    for (int i = 0; i < 8; i++)
                    {
                        sourceRow.Add(sourceFile.Rows[RowCountInSource][i]);
                    }
                }
                catch (ArgumentOutOfRangeException)
                {
                    break;
                }
                sourceData.Add(sourceRow);
                RowCountInSource++;
            }
            int l = 0;
            while (true)
            {
                List<string> localRow = new List<string>();
                try
                {
                    for (l = 0; l < 8; l++)
                    {
                        localRow.Add(localFile.Rows[RowCountInLocal][l]);
                    }
                }
                catch (ArgumentOutOfRangeException)
                {
                    //Если один или несколько последних столбцов полностью удалены, придется полностью обновить базу                    
                    if (l > 0 && RowCountInLocal != 0)
                    {
                        CreateLocalData();
                        MessageBox.Show("Локальная база полностью обновлена!");
                        return;
                    }
                    //Замаливаем грехи класса Excel
                    else if (RowCountInLocal == 0)
                    {
                        CreateLocalData();
                        MessageBox.Show("Локальная база полностью обновлена!");
                        return;
                    }
                    else
                    {
                        break;
                    }

                }
                localData.Add(localRow);
                RowCountInLocal++;
            }

            //Проверяем локальную базу на лишние записи
            for (int i = 0; i < localData.Count; i++)
            {
                bool isThisRow = false;
                for (int j = 0; j < sourceData.Count; j++)
                {
                    if (sourceData[j][0] == localData[i][0])
                    {
                        isThisRow = true;
                    }
                }
                if (!isThisRow)
                {
                    //Если есть полностью пустая строка между другими, не считаем ее в изменениях
                    if (localData[i][0] == "" && localData[i][1] == "" && localData[i][2] == "" && localData[i][3] == "" && localData[i][4] == "" && localData[i][5] == "" && localData[i][6] == "" && localData[i][7] == "")
                    {
                        continue;
                    }
                    idDeleted.Add(localFile.Rows[i][0]);
                    //Делаем лишнюю строку пустой
                    for (int k = 0; k < 8; k++)
                    {
                        localFile.Rows[i][k] = "";
                        localData[i][k] = "";
                    }
                }
            }

            //Проверяем локальную базу на дублирующиеся записи и удаляем их
            for (int i = 0; i < localData.Count; i++)
            {
                for (int j = i + 1; j < localData.Count; j++)
                {

                    if (localData[i][0] == localData[j][0] && localData[i][0] != "")
                    {
                        idDeleted.Add(localData[j][0]);
                        for (int k = 0; k < 8; k++)
                        {
                            localFile.Rows[j][k] = "";
                            localData[j][k] = "";
                        }

                    }

                }
            }


            for (int i = 0; i < sourceData.Count; i++)
            {
                bool isThisRow = false;
                for (int j = 0; j < localData.Count; j++)
                {
                    bool wasChanged = false;

                    //Проверяем наличие угрозы с идентификатором
                    if (sourceData[i][0] == localData[j][0])
                    {
                        for (int k = 0; k < 8; k++)
                        {
                            //Ищем изменения и записываем их, если есть
                            if (sourceData[i][k] != localData[j][k])
                            {
                                id.Add(sourceData[i][0]);
                                itWas.Add(localData[j][k]);
                                localFile.Rows[j][k] = sourceData[i][k];
                                itBecame.Add(localFile.Rows[j][k]);
                                wasChanged = true;
                            }

                        }
                        isThisRow = true;
                    }
                    //Считаем количество измененных записей
                    if (wasChanged)
                    {
                        changedRowsCount++;
                    }
                }
                //Если такой угрозы нет, добавляем ее (к сожалению, в конец)
                if (!isThisRow)
                {
                    localFile.AddRow(sourceData[i][0], sourceData[i][1], sourceData[i][2], sourceData[i][3], sourceData[i][4], sourceData[i][5], sourceData[i][6], sourceData[i][7]);
                    idAdded.Add(sourceData[i][0]);
                }
            }
            //Сохраняем изменения
            try
            {
                localFile.FileSave(@"LocalData.xlsx");
            }
            catch (IOException)
            {
                MessageBox.Show("Пожалуйста, закройте файл локальной базы.");
                return;
            }

            //Локальная база не изменилась
            if (changedRowsCount == 0 && idDeleted.Count == 0 && idAdded.Count == 0)
            {
                MessageBox.Show("Ошибка!\nЛокальная база уже обновлена до последней версии.");
                //Удаляем, если не было файла с таким именем
                if (!isFileExist)
                {
                    File.Delete("thrlist.xlsx");
                }
                return;
            }

            //Сообщение об обновлении
            string messageRefresh = $"Успешно!\nКоличество обновленных записей - {changedRowsCount}\n";
            for (int i = 0; i < id.Count; i++)
            {
                messageRefresh += $"Угроза {id[i]}: {itWas[i].Replace("_x000D_", "")} -> {itBecame[i].Replace("_x000D_", "")}\n";
            }
            string messageAdded = $"Добавлено записей - {idAdded.Count}\n";
            for (int i = 0; i < idAdded.Count; i++)
            {
                messageAdded += $"Угроза {idAdded[i]}\n";
            }
            string messageDeleted = $"Удалено записей - {idDeleted.Count}\n";
            for (int i = 0; i < idDeleted.Count; i++)
            {
                messageDeleted += $"Угроза {idDeleted[i]}\n";
            }
            RefresfMessageWingow rms = new RefresfMessageWingow(messageRefresh, messageAdded, messageDeleted);
            rms.ShowDialog();

            //Удаляем, если не было файла с таким именем
            if (!isFileExist)
            {
                File.Delete("thrlist.xlsx");
            }

        }

        private void refreshButton_Click(object sender, RoutedEventArgs e)
        {
            Refresh();
        }

        private void fastShowButton_Click(object sender, RoutedEventArgs e)
        {
            Excel localFile;
            try
            {
                localFile = DoesFileExist();
            }
            catch (IOException)
            {
                MessageBox.Show("Пожалуйста, закройте файл локальной базы.");
                return;
            }
            if (localFile == null)
            {
                MessageBox.Show("Ошибка!\nЛокальный файл не найден.");
                return;
            }
            startPosition = 0;
            List<ShortNote> data = new List<ShortNote>();
            int skipedRowCount = 0;
            for (int i = 0; i < 15 + skipedRowCount; i++)
            {
                //Проверяем наличие пустых строк
                bool isRowEmpty = true; ;
                for (int j = 0; j < 8; j++)
                {
                    if (localFile.Rows[i][j] != "")
                    {
                        isRowEmpty = false;
                    }
                }
                //Если не пустая
                if (!isRowEmpty)
                {
                    //И если в первой ячейке число, а вторая не пуста, то добавляем инфу в список на вывод
                    if (localFile.Rows[i][0] != "" && int.TryParse(localFile.Rows[i][0], out int l) && localFile.Rows[i][1] != "")
                    {
                        data.Add(new ShortNote("УБИ." + localFile.Rows[i][0], localFile.Rows[i][1]));
                    }
                    //Иначе выводим сообщение об ошибке с базой
                    else
                    {
                        MessageBox.Show("Что-то пошло не так :(\nПопробуйте обновить базу.");
                        return;
                    }
                }
                //Учитываем, что пустую строку надо пропустить и все равно вывести 15 строк
                else
                {
                    skipedRowCount++;
                    startPosition++;
                }
            }
            quickShowDataGrid.Visibility = Visibility.Visible;
            backButton.Content = @"<";
            nextButton.Visibility = Visibility.Visible;
            backButton.Visibility = Visibility.Hidden;
            noteTextBlock.Visibility = Visibility.Hidden;
            //Вывод
            quickShowDataGrid.ItemsSource = data;
            quickShowDataGrid.Columns[0].Header = "Идентификатор УБИ";
            quickShowDataGrid.Columns[1].Header = "Наименование УБИ";
            quickShowDataGrid.Columns[0].Width = 130;
            quickShowDataGrid.Columns[1].Width = 1200;

        }

        private void showNoteButton_Click(object sender, RoutedEventArgs e)
        {
            //Проверяем, ввели ли что-то
            if (noteIdBox.Text.Trim() == "")
            {
                MessageBox.Show("Идентификатор не введен!");
                return;
            }
            //Проверяем, ввели число или нет
            if (!int.TryParse(noteIdBox.Text, out int noteId))
            {
                MessageBox.Show("Введите число!");
                return;
            }

            Excel localFile;
            try
            {
                localFile = DoesFileExist();
            }
            catch (IOException)
            {
                MessageBox.Show("Пожалуйста, закройте файл локальной базы.");
                return;
            }
            if (localFile == null)
            {
                MessageBox.Show("Ошибка!\nЛокальный файл не найден.");
                return;
            }

            //Обнуляем блок текста
            noteTextBlock.Text = "";

            string[] note = new string[8];
            int localId = 0;
            bool wasFound = false;
            //Поиск угрозы с введенным идентификатором
            while (!wasFound)
            {
                try
                {
                    int.TryParse(localFile.Rows[localId][0], out int a);
                    if (a == noteId)
                    {
                        for (int i = 0; i < 8; i++)
                        {
                            if (localFile.Rows[localId][i] == "")
                            {
                                MessageBox.Show("Что-то пошло не так :(\nПопробуйте обновить базу.");
                                return;
                            }
                        }
                        note = new string[8] { localFile.Rows[localId][0], localFile.Rows[localId][1].Replace("_x000D_", ""), localFile.Rows[localId][2].Replace("_x000D_", ""), localFile.Rows[localId][3].Replace("_x000D_", ""), localFile.Rows[localId][4].Replace("_x000D_", ""), localFile.Rows[localId][5], localFile.Rows[localId][6], localFile.Rows[localId][7] };
                        //Останавливаем поиск, если нашли
                        wasFound = true;
                    }
                }
                catch (ArgumentOutOfRangeException)
                {
                    //Не нашли
                    MessageBox.Show($"Угроза с идентификатором {noteId} не найдена или записана не полностью!");
                    return;
                }
                catch (FormatException)
                {
                    MessageBox.Show("Что-то пошло не так :(\nПопробуйте обновить базу.");
                    return;
                }
                localId++;
            }
            //Вывод угрозы
            noteTextBlock.Text += "Идентификатор УБИ: " + note[0] + "\n";
            noteTextBlock.Text += "Наименование УБИ: " + note[1] + "\n";
            noteTextBlock.Text += "Описание: " + note[2] + "\n";
            noteTextBlock.Text += "Источник угрозы: " + note[3] + "\n";
            noteTextBlock.Text += "Объект воздействия: " + note[4] + "\n";
            noteTextBlock.Text += "Нарушение конфиденциальности: ";
            //Меняем 0 и 1 на нет и да
            try
            {
                if (int.Parse(note[5]) == 1)
                {
                    noteTextBlock.Text += "Да\n";
                }
                else
                {
                    noteTextBlock.Text += "Нет\n";
                }
                noteTextBlock.Text += "Нарушение целостности: ";
                if (int.Parse(note[6]) == 1)
                {
                    noteTextBlock.Text += "Да\n";
                }
                else
                {
                    noteTextBlock.Text += "Нет\n";
                }
                noteTextBlock.Text += "Нарушение доступности: ";
                if (int.Parse(note[7]) == 1)
                {
                    noteTextBlock.Text += "Да\n";
                }
                else
                {
                    noteTextBlock.Text += "Нет\n";
                }
            }
            catch (FormatException)
            {
                MessageBox.Show("Что-то пошло не так :(\nПопробуйте обновить базу.");
                return;
            }
            nextButton.Visibility = Visibility.Hidden;
            backButton.Visibility = Visibility.Hidden;
            quickShowDataGrid.Visibility = Visibility.Hidden;
            noteTextBlock.Visibility = Visibility.Visible;
        }

        //Сохраняем в файл
        private void saveButton_Click(object sender, RoutedEventArgs e)
        {
            Excel localFile;
            try
            {
                localFile = DoesFileExist();
            }
            catch (IOException)
            {
                MessageBox.Show("Пожалуйста, закройте файл локальной базы.");
                return;
            }
            if (localFile == null)
            {
                MessageBox.Show("Ошибка!\nЛокальный файл не найден.");
                return;
            }
            Excel saveFile = new Excel();
            //
            saveFile.AddRow("Идентификатор УБИ", "Наименование УБИ", "Описание", "Источник угрозы", "Объект воздействия", "Нарушение конфиденциальности", "Нарушение целостности", "Нарушение доступности");

            int localEnum = 0;
            while (true)
            {
                List<string> localRow = new List<string>();
                try
                {
                    for (int i = 0; i < 8; i++)
                    {
                        localRow.Add(localFile.Rows[localEnum][i]);
                    }
                }
                catch (ArgumentOutOfRangeException)
                {
                    break;
                }
                //Меняем 0 и 1 на нет и да
                int.TryParse(localRow[5], out int a);
                int.TryParse(localRow[6], out int b);
                int.TryParse(localRow[7], out int c);
                if (a == 1)
                {
                    localRow[5] = "Да";
                }
                else if (a == 0 && localRow[5] != "")
                {
                    localRow[5] = "Нет";
                }
                if (b == 1)
                {
                    localRow[6] = "Да";
                }
                else if (b == 0 && localRow[6] != "")
                {
                    localRow[6] = "Нет";
                }
                if (c == 1)
                {
                    localRow[7] = "Да";
                }
                else if (c == 0 && localRow[7] != "")
                {
                    localRow[7] = "Нет";
                }
                //Добавляем строку в файл
                saveFile.AddRow(localRow[0], localRow[1].Replace("_x000D_", ""), localRow[2].Replace("_x000D_", ""), localRow[3].Replace("_x000D_", ""), localRow[4].Replace("_x000D_", ""), localRow[5], localRow[6], localRow[7]);
                localEnum++;
            }
            //Спаршиваем путь у пользователя
            FolderBrowserDialog folderBrowser = new FolderBrowserDialog();
            folderBrowser.ShowDialog();
            if (string.IsNullOrEmpty(folderBrowser.SelectedPath))
            {
                return;
            }
            if (File.Exists($@"{folderBrowser.SelectedPath}\LocalData.xlsx"))
            {
                MessageBox.Show("В данной директории файл с именем LocalData.xlsx уже существует!");
                return;
            }
            //Сохраняем            
            saveFile.FileSave($@"{folderBrowser.SelectedPath}\LocalData.xlsx");
            MessageBox.Show("Сохранено!");
        }

        //Переменные для прокрутки страниц
        private static int startPosition = 0;
        private static bool isOver = false;

        private void nextButton_Click(object sender, RoutedEventArgs e)
        {
            //Проверка на последнюю страницу
            if (isOver)
            {
                return;
            }
            Excel localFile;
            try
            {
                localFile = DoesFileExist();
            }
            catch (IOException)
            {
                MessageBox.Show("Пожалуйста, закройте файл локальной базы.");
                return;
            }
            if (localFile == null)
            {
                MessageBox.Show("Ошибка!\nЛокальный файл не найден.");
                return;
            }

            //Позиция первого элемента следующей страницы
            startPosition += 15;
            List<ShortNote> data = new List<ShortNote>();
            int i = 0;
            int skipedRowCount = 0;
            try
            {
                
                for (i = startPosition; i < startPosition + 15 + skipedRowCount; i++)
                {
                    //Проверяем наличие пустых строк
                    bool isRowEmpty = true; ;
                    for (int j = 0; j < 8; j++)
                    {
                        if (localFile.Rows[i][j] != "")
                        {
                            isRowEmpty = false;
                        }
                    }
                    //Если не пустая
                    if (!isRowEmpty)
                    {
                        //И если в первой ячейке число, а вторая не пуста, то добавляем инфу в список на вывод
                        if (localFile.Rows[i][0] != "" && int.TryParse(localFile.Rows[i][0], out int l) && localFile.Rows[i][1] != "")
                        {
                            data.Add(new ShortNote("УБИ." + localFile.Rows[i][0], localFile.Rows[i][1]));
                        }
                        //Иначе выводим сообщение об ошибке с базой
                        else
                        {
                            startPosition -= 15;
                            MessageBox.Show("Что-то пошло не так :(\nПопробуйте обновить базу.");
                            return;
                        }
                    }
                    //Учитываем, что пустую строку надо пропустить и все равно вывести 15 строк
                    else
                    {
                        skipedRowCount++;                        
                    }
                }
            }
            catch (ArgumentOutOfRangeException)
            {
                quickShowDataGrid.ItemsSource = data;
                quickShowDataGrid.Columns[0].Header = "Идентификатор УБИ";
                quickShowDataGrid.Columns[1].Header = "Наименование УБИ";
                quickShowDataGrid.Columns[0].Width = 130;
                quickShowDataGrid.Columns[1].Width = 1200;
                nextButton.Visibility = Visibility.Hidden;
                isOver = true;
                return;
            }
            startPosition += skipedRowCount;
            //Вывод
            backButton.Visibility = Visibility.Visible;
            quickShowDataGrid.ItemsSource = data;
            quickShowDataGrid.Columns[0].Header = "Идентификатор УБИ";
            quickShowDataGrid.Columns[1].Header = "Наименование УБИ";
            quickShowDataGrid.Columns[0].Width = 130;
            quickShowDataGrid.Columns[1].Width = 1200;
        }

        private void backButton_Click(object sender, RoutedEventArgs e)
        {
            //Уже точно не последняя страница
            isOver = false;
            Excel localFile;
            try
            {
                localFile = DoesFileExist();
            }
            catch (IOException)
            {
                MessageBox.Show("Пожалуйста, закройте файл локальной базы.");
                return;
            }
            if (localFile == null)
            {
                MessageBox.Show("Ошибка!\nЛокальный файл не найден.");
                return;
            }

            //Позиция первого элемента предыдущей страницы
            startPosition -= 15;
            List<ShortNote> data;
            
            
            int notEmptyRowCount = 0;
            int skipedRowCount = 0;
            //Немного магии кнопки назад
            try
            {
                while (true)
                {
                    data = new List<ShortNote>();  
                    notEmptyRowCount = 0;
                    for (int i = startPosition; i < startPosition + 15 + skipedRowCount; i++)
                    {                        
                        bool isRowEmpty = true;
                        for (int j = 0; j < 8; j++)
                        {
                            if (localFile.Rows[i][j] != "")
                            {
                                isRowEmpty = false;
                            }
                        }
                        //Если не пустая
                        if (!isRowEmpty)
                        {
                            notEmptyRowCount++;
                            //И если в первой ячейке число, а вторая не пуста, то добавляем инфу в список на вывод
                            if (localFile.Rows[i][0] != "" && int.TryParse(localFile.Rows[i][0], out int l) && localFile.Rows[i][1] != "")
                            {
                                data.Add(new ShortNote("УБИ." + localFile.Rows[i][0], localFile.Rows[i][1]));
                            }
                            //Иначе выводим сообщение об ошибке с базой
                            else
                            {
                                startPosition += 15 + skipedRowCount;
                                MessageBox.Show("Что-то пошло не так :(\nПопробуйте обновить базу.");
                                return;
                            }
                        }
                    }                    
                    if (notEmptyRowCount < 15)
                    {                        
                        skipedRowCount++;
                        startPosition--;
                        if (startPosition < 0)
                        {                            
                            break;
                        }
                    }
                    else
                    {
                        break;
                    }
                }                
            }
            catch (ArgumentOutOfRangeException)
            {
                //Обработка на первой странице
                startPosition += 15 + skipedRowCount;
                return;
            }            
            //Вывод
            if (startPosition <= 0)
            {
                backButton.Visibility = Visibility.Hidden;
            }
            startPosition += skipedRowCount;
            nextButton.Visibility = Visibility.Visible;
            quickShowDataGrid.ItemsSource = data;
            quickShowDataGrid.Columns[0].Header = "Идентификатор УБИ";
            quickShowDataGrid.Columns[1].Header = "Наименование УБИ";
            quickShowDataGrid.Columns[0].Width = 130;
            quickShowDataGrid.Columns[1].Width = 1200;
        }

        //Проверяем, не удалили ли файл локальной базы
        private static Excel DoesFileExist()
        {
            Excel localFile = new Excel();
            try
            {
                localFile.FileOpen(@"LocalData.xlsx");
            }
            catch (FileNotFoundException)
            {
                return null;
            }
            return localFile;
        }

        //Прерываем поток автообновления, если программа закрыта
        private void Window_Closed(object sender, EventArgs e)
        {
            autoRefresh.Abort();
        }
    }

}
