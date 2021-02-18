using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Xml.Serialization;

namespace StatisticA2H
{
    [Serializable]
    public class Statistic
    {
        public List<OneValue> Values { get; set; }
        public Statistic()
        {
            Values = new List<OneValue>();
        }
        public void SaveListToXml(string fileName)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(List<OneValue>));
            using (FileStream steam = new FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                serializer.Serialize(steam, Values);
            }
            MessageBox.Show($"Файл {fileName} успешно сохранен");
        }

        public void LoadListFromXml(string fileName)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(List<OneValue>));
            using (FileStream steam = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.None))
            {
                List<OneValue> list = serializer.Deserialize(steam) as List<OneValue>;
                Values = list;
                MessageBox.Show($"Файл {fileName} успешно прочитан");
            }
        }

        public static string GetNameOperation(int operation)
        {
            string s = default;
            switch (operation)
            {
                case 1:
                    s = "Из смесителя на трансбордер (на вибростол 1)";
                    break;
                case 2:
                    s = "С трансбордера на вибростол 1 (со смесителя)";
                    break;
                case 3:
                    s = "Из смесителя на трансбордер (на вибростол 2)";
                    break;
                case 4:
                    s = "С трансбордера на вибростол 2 (со смесителя)";
                    break;
                case 5:
                    s = "Из смесителя на трансбордер (на место созревания)";
                    break;
                case 6:
                    s = "С трансбордера на место созревания (со смесителя)";
                    break;
                case 7:
                    s = "Из вибростола 1 на трансбордер (на место созревания)";
                    break;
                case 8:
                    s = "С трансбордера на место созревания (с вибростола 1)";
                    break;
                case 9:
                    s = "Из вибростола 2 на трансбордер (на место созревания)";
                    break;
                case 10:
                    s = "С трансбордера на место созревания (с вибростола 2)";
                    break;
                case 11:
                    s = "С места созревания на трансбордер полная форма (на подачу)";
                    break;
                case 12:
                    s = "С трансбордера на подачу полная форма (с места созревания)";
                    break;
                case 13:
                    s = "С места созревания на трансбордер пустая форма (на подачу)";
                    break;
                case 14:
                    s = "С трансбордера на подачу пустая форма (с места созревания)";
                    break;
                case 15:
                    s = "Управление вручную";
                    break;
                default:
                    break;
            }
            return s;
        }


    }
    [Serializable]
    public class OneValue
    {
        public int NumberPath { get; set; } //номер места
        [XmlAttribute]
        public int Operation { get; set; } //номер операции
        [XmlAttribute]
        public string StrOperation { get; set; } //название операции
        [XmlAttribute]
        public DateTime Time { get; set; } //время и дата записи
        [XmlAttribute]
        public int Span { get; set; } //время выполнения операции в секундах
    }
    public enum NUM_OPER
    {
        FROM_MIX_TO_VT1 = 1,
        FROM_MIX_TO_VT1_WORK,
        FROM_MIX_TO_VT2,
        FROM_MIX_TO_VT2_WORK,
        FROM_MIX_TO_KAM,
        FROM_MIX_TO_KAM_WORK,
        FROM_VT1_TO_KAM,
        FROM_VT1_TO_KAM_WORK,
        FROM_VT2_TO_KAM,
        FROM_VT2_TO_KAM_WORK,
        FROM_KAM_TO_OUT_FULL,
        FROM_KAM_TO_OUT_FULL_WORK,
        FROM_KAM_TO_OUT,
        FROM_KAM_TO_OUT_WORK,
        MANUAL_WORK
    };
}
