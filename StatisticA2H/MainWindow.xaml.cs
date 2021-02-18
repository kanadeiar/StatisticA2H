using System;
using System.IO;
using System.Reflection;
using System.Timers;
using System.Windows;
using Microsoft.Win32;
using Sharp7;
using Excel = Microsoft.Office.Interop.Excel;

namespace StatisticA2H
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        static Statistic _statistic = new Statistic();
        static Timer timer = new Timer();
        static S7Client s7client = new S7Client();
        static Random random = new Random();
        private static string _fileName = default;

        NUM_OPER currentOperation = default; //текущая операция выполняемая
        int numberPath; //номер цели
        DateTime startTime = new DateTime(); //время начала текущей операции
        int oldPath = default; //старый номер цели
        bool simulationOn = false; //включение симуляции
        int simStep = default; //шаг симуляции
        private bool startCalculate; //начало подсчета - пропуск подсчета этой операции

        int rez = default;
        bool fromMixToVT1 = default;
        bool fromMixToVT1Work = default;
        bool fromMixToVT2 = default;
        bool fromMixToVT2Work = default;
        bool fromMixToKam = default;
        bool fromMixToKamWork = default;
        bool fromVT1ToKam = default;
        bool fromVT1ToKamWork = default;
        bool fromVT2ToKam = default;
        bool fromVT2ToKamWork = default;
        bool fromKamToOutFull = default;
        bool fromKamToOutFullWork = default;
        bool fromKamToOut = default;
        bool fromKamToOutWork = default;
        bool readyWork = default;
        bool manualWork = default;
        int valuePath = default;
        public MainWindow()
        {
            InitializeComponent();
            timer.Interval = 1000;
            timer.Elapsed += Timer_Elapsed;
        }
        private void Timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            timer.Enabled = false;
            Calculate();
            timer.Enabled = true;
        }
        /// <summary>
        /// Чтение новых параметров и их запоминание по необходимости
        /// </summary>
        private void Calculate()
        {
            byte[] buffer = new byte[16];

            if (!simulationOn)
            {
                rez = s7client.ReadArea(S7Consts.S7AreaMK, 0, 282, 4, S7Consts.S7WLByte, buffer);
                fromMixToVT1 = (buffer[1] & 0b100000) != 0;
                fromMixToVT1Work = (buffer[2] & 0b10) != 0;
                fromMixToVT2 = (buffer[2] & 0b100) != 0;
                fromMixToVT2Work = (buffer[2] & 0b1000) != 0;
                fromMixToKam = (buffer[0] & 0b10) != 0;
                fromMixToKamWork = (buffer[0] & 0b100) != 0;
                fromVT1ToKam = (buffer[2] & 0b10000) != 0;
                fromVT1ToKamWork = (buffer[2] & 0b100000) != 0;
                fromVT2ToKam = (buffer[2] & 0b1000000) != 0;
                fromVT2ToKamWork = (buffer[2] & 0b10000000) != 0;
                fromKamToOutFull = (buffer[0] & 0b1000) != 0;
                fromKamToOutFullWork = (buffer[0] & 0b10000) != 0;
                fromKamToOut = (buffer[0] & 0b1000000) != 0;
                fromKamToOutWork = (buffer[0] & 0b10000000) != 0;
                readyWork = (buffer[0] & 0b1) != 0;
                manualWork = !fromMixToVT1 && !fromMixToVT1Work && !fromMixToVT2 && !fromMixToVT2Work
                    && !fromMixToKam && !fromMixToKamWork && !fromVT1ToKam && !fromVT1ToKamWork && !fromVT2ToKam
                    && !fromVT2ToKamWork && !fromKamToOutFull && !fromKamToOutFullWork && !fromKamToOut && !fromKamToOutWork && !readyWork;
                s7client.DBRead(201, 100, 2, buffer);
                valuePath = S7.GetIntAt(buffer, 0);
            }
            else
            {
                if (random.Next(20) == 5)
                {
                    if (simStep == 0)
                    {
                        manualWork = false;
                        fromMixToVT1 = true;
                        valuePath = 400;
                        simStep = 1;
                    } 
                    else if (simStep == 1)
                    {
                        fromMixToVT1Work = true;
                        valuePath = 210;
                        simStep = 2;
                    }
                    else if (simStep == 2)
                    {
                        fromMixToVT1 = false;
                        fromMixToVT1Work = false;
                        fromVT2ToKam = true;
                        valuePath = 220;
                        simStep = 3;
                    }
                    else if (simStep == 3)
                    {
                        fromVT2ToKamWork = true;
                        valuePath = random.Next(20) + 2;
                        simStep = 4;
                    }
                    else if (simStep == 4)
                    {
                        fromVT2ToKam = false;
                        fromVT2ToKamWork = false;
                        fromKamToOutFull = true;
                        valuePath = random.Next(20) + 2;
                        simStep = 5;
                    }
                    else if (simStep == 5)
                    {
                        fromKamToOutFullWork = true;
                        valuePath = 300;
                        simStep = 6;
                    }
                    else if (simStep == 6)
                    {
                        fromKamToOutFull = false;
                        fromKamToOutFullWork = false;
                        manualWork = true;
                        valuePath = random.Next(20) + 2;
                        simStep = 0;
                    }
                }
            }
 
            //опеределение старта записи данных, начало операции
            if (currentOperation == 0)
            {
                if (fromMixToVT1)
                {
                    currentOperation = NUM_OPER.FROM_MIX_TO_VT1;
                    RedataCalc(valuePath);
                } 
                else if (fromMixToVT2)
                {
                    currentOperation = NUM_OPER.FROM_MIX_TO_VT2;
                    RedataCalc(valuePath);
                } 
                else if (fromMixToKam)
                {
                    currentOperation = NUM_OPER.FROM_MIX_TO_KAM;
                    RedataCalc(valuePath);
                } 
                else if (fromVT1ToKam)
                {
                    currentOperation = NUM_OPER.FROM_VT1_TO_KAM;
                    RedataCalc(valuePath);
                } 
                else if (fromVT2ToKam)
                {
                    currentOperation = NUM_OPER.FROM_VT2_TO_KAM;
                    RedataCalc(valuePath);
                }
                else if (fromKamToOutFull)
                {
                    currentOperation = NUM_OPER.FROM_KAM_TO_OUT_FULL;
                    RedataCalc(valuePath);
                }
                else if (fromKamToOut)
                {
                    currentOperation = NUM_OPER.FROM_KAM_TO_OUT;
                    RedataCalc(valuePath);
                }
                else if (manualWork)
                {
                    currentOperation = NUM_OPER.MANUAL_WORK;
                    RedataCalc(valuePath);
                }
            }
            //добавление записанных данных от одной операции к другой
            if (currentOperation == NUM_OPER.FROM_MIX_TO_VT1 && fromMixToVT1Work)
            {
                SaveCalc(oldPath, startTime, currentOperation);
                RedataCalc(valuePath);
                currentOperation = NUM_OPER.FROM_MIX_TO_VT1_WORK;
            }
            if (currentOperation == NUM_OPER.FROM_MIX_TO_VT2 && fromMixToVT2Work)
            {
                SaveCalc(oldPath, startTime, currentOperation);
                RedataCalc(valuePath);
                currentOperation = NUM_OPER.FROM_MIX_TO_VT2_WORK;
            }
            if (currentOperation == NUM_OPER.FROM_MIX_TO_KAM && fromMixToKamWork)
            {
                SaveCalc(oldPath, startTime, currentOperation);
                RedataCalc(valuePath);
                currentOperation = NUM_OPER.FROM_MIX_TO_KAM_WORK;
            }
            if (currentOperation == NUM_OPER.FROM_VT1_TO_KAM && fromVT1ToKamWork)
            {
                SaveCalc(oldPath, startTime, currentOperation);
                RedataCalc(valuePath);
                currentOperation = NUM_OPER.FROM_VT1_TO_KAM_WORK;
            }
            if (currentOperation == NUM_OPER.FROM_VT2_TO_KAM && fromVT2ToKamWork)
            {
                SaveCalc(oldPath, startTime, currentOperation);
                RedataCalc(valuePath);
                currentOperation = NUM_OPER.FROM_VT2_TO_KAM_WORK;
            }
            if (currentOperation == NUM_OPER.FROM_KAM_TO_OUT_FULL && fromKamToOutFullWork)
            {
                SaveCalc(oldPath, startTime, currentOperation);
                RedataCalc(valuePath);
                currentOperation = NUM_OPER.FROM_KAM_TO_OUT_FULL_WORK;
            }
            if (currentOperation == NUM_OPER.FROM_KAM_TO_OUT && fromKamToOutWork)
            {
                SaveCalc(oldPath, startTime, currentOperation);
                RedataCalc(valuePath);
                currentOperation = NUM_OPER.FROM_KAM_TO_OUT_WORK;
            }
            //добавление записанных данных выполненной операции
            if (currentOperation == NUM_OPER.FROM_MIX_TO_VT1_WORK && !fromMixToVT1Work)
            {
                SaveCalc(oldPath, startTime, currentOperation);
                currentOperation = default;
            }
            if (currentOperation == NUM_OPER.FROM_MIX_TO_VT2_WORK && !fromMixToVT2Work)
            {
                SaveCalc(oldPath, startTime, currentOperation);
                currentOperation = default;
            }
            if (currentOperation == NUM_OPER.FROM_MIX_TO_KAM_WORK && !fromMixToKamWork)
            {
                SaveCalc(oldPath, startTime, currentOperation);
                currentOperation = default;
            }
            if (currentOperation == NUM_OPER.FROM_VT1_TO_KAM_WORK && !fromVT1ToKamWork)
            {
                SaveCalc(oldPath, startTime, currentOperation);
                currentOperation = default;
            }
            if (currentOperation == NUM_OPER.FROM_VT2_TO_KAM_WORK && !fromVT2ToKamWork)
            {
                SaveCalc(oldPath, startTime, currentOperation);
                currentOperation = default;
            }
            if (currentOperation == NUM_OPER.FROM_KAM_TO_OUT_FULL_WORK && !fromKamToOutFullWork)
            {
                SaveCalc(oldPath, startTime, currentOperation);
                currentOperation = default;
            }
            if (currentOperation == NUM_OPER.FROM_KAM_TO_OUT_WORK && !fromKamToOutWork)
            {
                SaveCalc(oldPath, startTime, currentOperation);
                currentOperation = default;
            }
            //добавление записи - отмененных вручную операций
            if (currentOperation == NUM_OPER.FROM_MIX_TO_VT1 && !fromMixToVT1)
            {
                SaveCalc(oldPath, startTime, currentOperation);
                currentOperation = default;
            }
            if (currentOperation == NUM_OPER.FROM_MIX_TO_VT2 && !fromMixToVT2)
            {
                SaveCalc(oldPath, startTime, currentOperation);
                currentOperation = default;
            }
            if (currentOperation == NUM_OPER.FROM_MIX_TO_KAM && !fromMixToKam)
            {
                SaveCalc(oldPath, startTime, currentOperation);
                currentOperation = default;
            }
            if (currentOperation == NUM_OPER.FROM_VT1_TO_KAM && !fromVT1ToKam)
            {
                SaveCalc(oldPath, startTime, currentOperation);
                currentOperation = default;
            }
            if (currentOperation == NUM_OPER.FROM_VT2_TO_KAM && !fromVT2ToKam)
            {
                SaveCalc(oldPath, startTime, currentOperation);
                currentOperation = default;
            }
            if (currentOperation == NUM_OPER.FROM_KAM_TO_OUT_FULL && !fromKamToOutFull)
            {
                SaveCalc(oldPath, startTime, currentOperation);
                currentOperation = default;
            }
            if (currentOperation == NUM_OPER.FROM_KAM_TO_OUT && !fromKamToOut)
            {
                SaveCalc(oldPath, startTime, currentOperation);
                currentOperation = default;
            }
            //изменение задания при ручной операции
            if (currentOperation == NUM_OPER.MANUAL_WORK && manualWork && valuePath != oldPath)
            {
                SaveCalc(oldPath, startTime, currentOperation);
            }
            //конец ручной операции
            if (currentOperation == NUM_OPER.MANUAL_WORK && !manualWork)
            {
                SaveCalc(oldPath, startTime, currentOperation);
                currentOperation = default;
            }

            //вывод данных
            LabelResult.Dispatcher.Invoke(() =>
            {
                LabelResult.Content = $"Полученные данные:\n" +
                $"Из смесителя на трансбордер (на вибростол 1) = {fromMixToVT1}\n" +
                $"С трансбордера на вибростол 1 (со смесителя) = {fromMixToVT1Work}\n" +
                $"Из смесителя на трансбордер (на вибростол 2) = {fromMixToVT2}\n" +
                $"С трансбордера на вибростол 2 (со смесителя) = {fromMixToVT2Work}\n" +
                $"Из смесителя на трансбордер (на место созревания) = {fromMixToKam}\n" +
                $"С трансбордера на место созревания (со смесителя) = {fromMixToKamWork}\n" +
                $"Из вибростола 1 на трансбордер (на место созревания) = {fromVT1ToKam}\n" +
                $"С транбордера на место созревания (с вибростола 1) = {fromVT1ToKamWork}\n" +
                $"Из вибростола 2 на трансбордер (на место созревания) = {fromVT2ToKam}\n" +
                $"С транбордера на место созревания (с вибростола 1) = {fromVT2ToKamWork}\n" +
                $"С места созревания на трансбордер полная форма (на подачу) = {fromKamToOutFull}\n" +
                $"С трансбордера на подачу полная форма (с места созревания) = {fromKamToOutFullWork}\n" +
                $"С места созревания на трансбордер пустая форма (на подачу) = {fromKamToOut}\n" +
                $"С трансбордера на подачу пустая форма (с места созревания) = {fromKamToOutWork}\n" +
                $"Ожидание = {readyWork}\n" +
                $"Управление вручную = {manualWork}\n" +
                $"Цель место = {valuePath}";
            });
            ListBoxResult.Dispatcher.Invoke(() => 
            {
                PrintDataToListBox();
            });

            oldPath = valuePath;
        }
        //перекалькуляция данных статистики текущей операции
        void RedataCalc(int valuePath)
        {
            numberPath = valuePath;
            startTime = DateTime.Now;
        }
        //запоминание данных статистики прошедшей операции
        void SaveCalc(int valuePath, DateTime startTime, NUM_OPER numberOperation)
        {
            if (startCalculate)
                startCalculate = false;
            else
            {
                DateTime endTime = DateTime.Now;
                string s = Statistic.GetNameOperation((int)numberOperation);
                _statistic.Values.Add(
                    new OneValue
                    {
                        NumberPath = valuePath,
                        Operation = (int)numberOperation,
                        StrOperation = s,
                        Time = DateTime.Now,
                        Span = (int)(endTime - startTime).TotalSeconds,
                    });
            }
        }
        //вывод данных в лист бокс формы
        void PrintDataToListBox()
        {
            int selectNum = ListBoxResult.SelectedIndex;
            ListBoxResult.Items.Clear();
            foreach (var el in _statistic.Values)
            {
                ListBoxResult.Items.Add($"{el.NumberPath} Место.  |  Дата: {el.Time}  |   Время: {el.Span / 60} м. {el.Span % 60} с.  |  Операция: {Statistic.GetNameOperation(el.Operation)}");
            }
            if (selectNum < ListBoxResult.Items.Count)
            {
                ListBoxResult.SelectedIndex = selectNum;
            }
        }
        //начало подсчета
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (!simulationOn)
            {
                //инит соединения
                int rez = s7client.ConnectTo("10.0.57.10", 0, 2);
                if (rez == 0)
                {
                    s7client.Connect();
                    if (rez != 0)
                        MessageBox.Show($"Ошибка соединения c ПЛК 10.0.57.10\nОшибка: {rez}");
                }
                else
                    MessageBox.Show($"Ошибка установки соединения c ПЛК 10.0.57.10\nОшибка: {rez}");
                if (rez == 0)
                    timer.Start();
                CheckBoxSim.IsEnabled = false;
            }
            else
            {
                timer.Start();
                CheckBoxSim.IsEnabled = false;
            }
            startCalculate = true;
        }
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (!simulationOn)
            {
                timer.Stop();
                int rez = s7client.Disconnect();
                if (rez == 0)
                    MessageBox.Show("Благополучное отсоединение от ПЛК");
                else
                    MessageBox.Show($"Ошибка отсоединения от ПЛК 10.0.57.10\nОшибка: {rez}");
                CheckBoxSim.IsEnabled = true;
            }
            else
            {
                timer.Stop();
                CheckBoxSim.IsEnabled = true;
            }

        }
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (MessageBox.Show("Действительно выйти из программы?", "Предупреждение о выходе", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.No)
                e.Cancel = true;
            else
            {
                if (!string.IsNullOrEmpty(_fileName))
                {
                    if (MessageBox.Show("Сохранить собранные данные в файл?", "Сохранить ли данные в файл", MessageBoxButton.YesNo,
                        MessageBoxImage.Question) == MessageBoxResult.Yes)
                    {
                        _statistic.SaveListToXml(_fileName);
                    }
                }
                timer.Stop();
            }
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            WindowStatistic window = new WindowStatistic();
            window.Owner = Window.GetWindow(this);
            window.Values = _statistic.Values;
            window.ShowDialog();
        }
        //новый файл
        private void NewItem_Click(object sender, RoutedEventArgs e)
        {
            _statistic = new Statistic();
            PrintDataToListBox();
        }

        private void OpenItem_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            string initDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            dialog.InitialDirectory = initDir;
            dialog.Filter = "Файлы статстики (*.xml)|*.xml|Все файлы (*.*)|*.*";
            dialog.FileName = $"Статистика.xml";
            if (dialog.ShowDialog() == true)
            {
                _fileName = dialog.FileName;
                LabelFilePath.Content = $"Путь к файлу: {_fileName}";
                _statistic.LoadListFromXml(_fileName);
                PrintDataToListBox();
            }
        }

        private void SaveItem_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_fileName))
            {
                SaveAsItem_Click(sender, e);
            }
            else
            {
                _statistic.SaveListToXml(_fileName);
            }
        }

        private void SaveAsItem_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog();
            string initDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            dialog.InitialDirectory = initDir;
            dialog.Filter = "Файлы статстики (*.xml)|*.xml|Все файлы (*.*)|*.*";
            dialog.FileName = $"Статистика.xml";
            if (dialog.ShowDialog() == true)
            {
                _fileName = Path.Combine(initDir, dialog.FileName);
                LabelFilePath.Content = $"Путь к файлу: {_fileName}";
                _statistic.SaveListToXml(_fileName);
            }
        }

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            Button_Click_1(sender, e);
            Close();
        }

        private void CheckBoxSim_Checked(object sender, RoutedEventArgs e)
        {
            simulationOn = (bool)CheckBoxSim.IsChecked;
        }

        /// <summary> Вывод данных в Excel </summary>
        private void ButtonExcel_OnClick(object sender, RoutedEventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            excel.Workbooks.Add();
            Excel.Worksheet worksheet = excel.ActiveSheet;
            worksheet.Cells[1, 1] = "Дата";
            worksheet.Cells[1, 2] = "Место";
            worksheet.Cells[1, 3] = "Операция";
            worksheet.Cells[1, 4] = "Время, сек.";
            excel.Columns.ColumnWidth = 16;
            worksheet.Columns[3].ColumnWidth = 50;
            for (int i = 0; i < _statistic.Values.Count; i++)
            {
                worksheet.Cells[i + 2, 1] = _statistic.Values[i].Time;
                worksheet.Cells[i + 2, 2] = _statistic.Values[i].NumberPath;
                worksheet.Cells[i + 2, 3] = Statistic.GetNameOperation(_statistic.Values[i].Operation);
                worksheet.Cells[i + 2, 4] = _statistic.Values[i].Span;
            }
            worksheet.Range[$"A1:D{_statistic.Values.Count + 1}"].AutoFormat(Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic1);
            excel.Visible = true;
        }
    }
}
