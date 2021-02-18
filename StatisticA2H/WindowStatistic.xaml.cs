using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Media.TextFormatting;
using Excel = Microsoft.Office.Interop.Excel;

namespace StatisticA2H
{
    /// <summary>
    /// Логика взаимодействия для WindowStatistic.xaml
    /// </summary>
    public partial class WindowStatistic : Window
    {
        public List<OneValue> Values { get; set; }
        private List<Stat> myList = new List<Stat>();
        public WindowStatistic()
        {
            InitializeComponent();
        }

        private void ButtonBack_OnClick(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Dictionary<KeyStat, ValueStat> dictionary = new Dictionary<KeyStat, ValueStat>();
            if (Values != null)
            {
                foreach (var el in Values)
                {
                    var tempKey = new KeyStat
                    {
                        NumberPath = el.NumberPath,
                        Operation = el.Operation,
                    };
                    if (dictionary.Keys.Contains(tempKey))
                    {
                        var tmpVal = dictionary[tempKey];
                        dictionary[tempKey] = new ValueStat
                        {
                            Count = tmpVal.Count + 1,
                            SecondsTotal = tmpVal.SecondsTotal + el.Span,
                        };
                    }
                    else
                    {
                        dictionary.Add(tempKey, new ValueStat
                        {
                            Count = 1,
                            SecondsTotal = el.Span,
                        });
                    }
                }
                
                foreach (var el in dictionary)
                {
                    var averSec = Convert.ToDouble(el.Value.SecondsTotal) / Convert.ToDouble(el.Value.Count);
                    myList.Add(new Stat
                    {
                        NumOper = el.Key.Operation,
                        Operation = Statistic.GetNameOperation(el.Key.Operation),
                        NumPath = el.Key.NumberPath,
                        Count = el.Value.Count,
                        AverageSec = averSec,
                        StrAverage = $"{((int)averSec / 60)} м. {(int)(averSec % 60)} с."
                    });
                }
                ListViewStatistic.ItemsSource = myList;
                if (myList.Count > 0)
                {
                    var averAll = Values.Average(a => a.Span);
                    var sum = Values.Where(l => l.Operation >= 1 && l.Operation <= 14).Sum(a => a.Span);
                    var count = Values.Count(l => l.Operation >= 1 && l.Operation <= 14);
                    double averAuto = (double)sum / count;
                    LabelAverAll.Content = $"Среднее арифметическое затраченное время по всем операциям: {averAll:F2} сек.\nСреднее арифметическое затраченное время только по операциям автоматической работы: {averAuto:f2} сек.";
                }
                else
                {
                    LabelAverAll.Content = $"Статистика недоступна при отсуствующих данных.";
                }

            }
            else
            {
                MessageBox.Show("Статистика отсутствует!");
            }

        }

        struct KeyStat
        {
            public int Operation { get; set; } //операция
            public int NumberPath { get; set; } //место
        }
        struct ValueStat
        {
            public int Count { get; set; } //количество
            public int SecondsTotal { get; set; } //время в секундах сумма
        }
        struct Stat
        {
            public int NumOper { get; set; }
            public string Operation { get; set; }
            public int NumPath { get; set; }
            public int Count { get; set; }
            public double AverageSec { get; set; }
            public string StrAverage { get; set; }
        }
        private void ButtonSortOper_OnClick(object sender, RoutedEventArgs e)
        {
            myList.Sort(((stat, stat1) => string.Compare(stat.Operation, stat1.Operation, true)));
            ListViewStatistic.ItemsSource = null;
            ListViewStatistic.ItemsSource = myList;
        }
        private void ButtonSortPath_OnClick(object sender, RoutedEventArgs e)
        {
            myList.Sort((((stat, stat1) =>
            {
                if (stat.NumPath > stat1.NumPath)
                    return 1;
                else if (stat.NumPath < stat1.NumPath)
                    return -1;
                else
                    return 0;
            })));
            ListViewStatistic.ItemsSource = null;
            ListViewStatistic.ItemsSource = myList;
        }
        private void ButtonSortCount_OnClick(object sender, RoutedEventArgs e)
        {
            myList.Sort((((stat, stat1) =>
            {
                if (stat.Count < stat1.Count)
                    return 1;
                else if (stat.Count > stat1.Count)
                    return -1;
                else
                    return 0;
            })));
            ListViewStatistic.ItemsSource = null;
            ListViewStatistic.ItemsSource = myList;
        }
        private void ButtonSortTime_OnClick(object sender, RoutedEventArgs e)
        {
            myList.Sort((((stat, stat1) =>
            {
                if (stat.AverageSec < stat1.AverageSec)
                    return 1;
                else if (stat.AverageSec > stat1.AverageSec)
                    return -1;
                else
                    return 0;
            })));
            ListViewStatistic.ItemsSource = null;
            ListViewStatistic.ItemsSource = myList;
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {

        }
        /// <summary> Вывод в Excel </summary>
        private void ButtonExcel_OnClick(object sender, RoutedEventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            excel.Workbooks.Add();
            Excel.Worksheet worksheet = excel.ActiveSheet;
            worksheet.Cells[1, 1] = "Операция";
            worksheet.Cells[1, 2] = "Место";
            worksheet.Cells[1, 3] = "Количество";
            worksheet.Cells[1, 4] = "Среднее время, сек.";
            excel.Columns.ColumnWidth = 16;
            worksheet.Columns[1].ColumnWidth = 50;
            for (int i = 0; i < myList.Count; i++)
            {
                worksheet.Cells[i + 2, 1] = myList[i].Operation;
                worksheet.Cells[i + 2, 2] = myList[i].NumPath;
                worksheet.Cells[i + 2, 3] = myList[i].Count;
                worksheet.Cells[i + 2, 4] = myList[i].AverageSec;
            }
            worksheet.Range[$"A1:D{myList.Count + 1}"].AutoFormat(Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic1);
            worksheet.Cells[myList.Count + 2, 1] = "Итого среднее время выполнения всех автоматических операций, сек:";
            var sum = Values.Where(l => l.Operation >= 1 && l.Operation <= 14).Sum(a => a.Span);
            var count = Values.Count(l => l.Operation >= 1 && l.Operation <= 14);
            double averAuto = (double)sum / count;
            worksheet.Cells[myList.Count + 2, 4] = averAuto;
            excel.Visible = true;
        }
    }
}
