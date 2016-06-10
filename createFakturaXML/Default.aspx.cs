using System;
using System.IO;

using System.Text;
using System.Diagnostics;
using System.Xml;
using System.Net;
using System.Web;

namespace createFakturaXML
{
    public partial class _Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }


        protected void Button3_Click(object sender, EventArgs e)
        {
            String savePath = @"c:\temp\";
            if (FileUpload1.HasFile && FileUpload1.FileName.Substring(FileUpload1.FileName.Length - 4).ToLower() == "xlsx")
            {
                string fileNameShort = FileUpload1.FileName.Substring(0, FileUpload1.FileName.Length - 5);

                string[] separator = new string[] { "ЛС № " };
                string[] separator11 = new string[] { "за" };
                string[] separator1 = new string[] { "ФИО (наименование плательщика) " };
                string[] separator2 = new string[] { "Адрес помещения " };
                string[] separator3 = new string[] { "S помещения:" };
                string[] separator31 = new string[] { "Площадь помещения:" };
                string[] separator4 = new string[] { "Проп./прож.: " };
                string[] separator41 = new string[] { "Прописано/проживает: " };
                string[] separator5 = new string[] { "S дома" };
                string[] separator51 = new string[] { "Площадь дома" };
                string[] separator6 = new string[] { ", в т.ч. МОП" };
                string[] separator7 = new string[] { "Прож. в доме" };
                string[] separator71 = new string[] { "Проживает в доме" };
                string[] separator711 = new string[] { "проживает в доме" };
                string[] separatorM = new string[] { "м2" };
                string[] separatorComma = new string[] { "," };
                string[] separatorP = new string[] { "чел." };

                string[] separator8 = new string[] { "исполнителя услуг" };

                //string[] separator9 = new string[] { "ИНН" };

                string[] separatorLs = new string[] { "Л/счет" };

                
                StreamWriter errorRow = new StreamWriter(savePath + @"error.txt", false, Encoding.Default);
                String fileName = FileUpload1.FileName;
                //savePath += fileName;
                FileUpload1.SaveAs(savePath + fileName);
                string path1 = savePath + fileName;
                var wb = new ClosedXML.Excel.XLWorkbook(path1);
                XmlTextWriter myXml = new XmlTextWriter(savePath + fileNameShort + ".xml", Encoding.Default);
                myXml.Formatting = Formatting.Indented;
                myXml.WriteStartDocument(true);
                myXml.WriteStartElement("Фаил");
                try
                {
                    #region ТЛТ
                    if (isTlt.Checked)
                    {
                        for (int i = 1; i <= 40000; i = i + 45)
                        {
                            if (
                                Convert.ToString(wb.Worksheet(1).Row(i+1).Cell(1).Value)
                                    .Contains("ФИО (наименование"))
                            {
                                if (i != 1)
                                    myXml.WriteEndElement();
                                myXml.WriteStartElement("ЛицевойСчет");

                                #region Раздел1

                                myXml.WriteStartElement("Раздел1");
                                myXml.WriteElementString("Период",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 1).Cell(1).Value)
                                        .Trim()
                                        .Split(separator11, StringSplitOptions.None)[1].Trim().Split(' ')[0] + " " +
                                    Convert.ToString(wb.Worksheet(1).Row(i + 1).Cell(1).Value)
                                        .Trim()
                                        .Split(separator11, StringSplitOptions.None)[1].Trim().Split(' ')[1]);

                                if (
                                    Convert.ToString(wb.Worksheet(1).Row(i + 1).Cell(1).Value)
                                        .ToUpper()
                                        .Contains("НЕ ПРИВАТИЗИРОВАНА"))
                                    myXml.WriteElementString("Приватизирована", "Нет");
                                else
                                    myXml.WriteElementString("Приватизирована", "Да");
                                myXml.WriteElementString("НомерЛС",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 1).Cell(1).Value)
                                        .Split(separator, StringSplitOptions.None)[1].Trim()
                                        .Split(separator1, StringSplitOptions.None)[0].Trim());
                                myXml.WriteElementString("ФИО",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 1).Cell(1).Value)
                                        .Trim()
                                        .Split(separator1, StringSplitOptions.None)[1]
                                        .Split(separator2, StringSplitOptions.None)[0].Trim());
                                myXml.WriteElementString("АдресПомещения",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 1).Cell(1).Value)
                                        .Trim()
                                        .Split(separator2, StringSplitOptions.None)[1].Trim().
                                        Split(separator3, StringSplitOptions.None)[0]);

                                myXml.WriteElementString("ПлощадьПомещения",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 1).Cell(1).Value)
                                        .Split(separator3, StringSplitOptions.None)[1]
                                        .Split(separatorM, StringSplitOptions.None)[0].Trim());
                                myXml.WriteElementString("ПрописаноПроживает",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 1).Cell(1).Value)
                                        .Split(separator4, StringSplitOptions.None)[1]
                                        .Split(separatorP, StringSplitOptions.None)[0].Trim());
                                myXml.WriteElementString("ПлощадьДома",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 1).Cell(1).Value)
                                        .Split(separator5, StringSplitOptions.None)[1]
                                        .Split(separatorM, StringSplitOptions.None)[0].Trim());
                                myXml.WriteElementString("ПлощадьМОП",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 1).Cell(1).Value)
                                        .Split(separator6, StringSplitOptions.None)[1]
                                        .Split(separatorM, StringSplitOptions.None)[0].Trim());
                                myXml.WriteElementString("Проживает",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 1).Cell(1).Value)
                                        .Split(separator7, StringSplitOptions.None)[1]
                                        .Split(separatorP, StringSplitOptions.None)[0].Trim());
                                myXml.WriteElementString("Организация",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 3).Cell(1).Value)
                                        .Split(separator8, StringSplitOptions.None)[1].Trim());
                                myXml.WriteEndElement();

                                #endregion Рвздел1

                                #region Раздел2

                                myXml.WriteStartElement("Раздел2");
                                string paymentTo = Convert.ToString(wb.Worksheet(1).Row(i + 2).Cell(22).Value).Trim();
                                myXml.WriteElementString("ПолучательПлатежа", paymentTo);
                                //string bankInfo = Convert.ToString(wb.Worksheet(1).Row(i + 2).Cell(25).Value).Trim();
                                string bankInfo =
                                    "Р/с - 40702810754400005587   Кор/счет-30101810200000000607    БИК 043601607 Поволжский банк ОАО «Сбербанк России» г. Самара";
                                myXml.WriteElementString("БанковскийСчет", bankInfo.Trim());

                                myXml.WriteElementString("ПлатежныйКод",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 2).Cell(33).Value)
                                        .Trim()
                                        .Replace('*', ' ')
                                        .Trim());

                                myXml.WriteElementString("ВидПлаты",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 2).Cell(39).Value)
                                        .Trim()
                                        .Replace("(ООО «УК «Ассоциация Управляющих Компаний»)", " ")
                                        .Trim());

                                myXml.WriteElementString("СуммаКОплате",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 2).Cell(43).Value).Trim());

                                myXml.WriteEndElement();

                                #endregion Рвздел2

                                #region Раздел3

                                myXml.WriteStartElement("Раздел3");
                                Decimal domTotal = 0;
                                for (int t = i + 10; t <= i + 31; t++)
                                {
                                    if (wb.Worksheet(1).Row(t).Cell(1).Value != null &&
                                        Convert.ToString(wb.Worksheet(1).Row(t).Cell(1).Value).Trim() != "")
                                    {
                                        myXml.WriteStartElement("Услуга");
                                        myXml.WriteElementString("ВидУслуги",
                                            Convert.ToString(wb.Worksheet(1).Row(t).Cell(1).Value).Trim());
                                        myXml.WriteElementString("ЕдиницаИзмерения",
                                            Convert.ToString(wb.Worksheet(1).Row(t).Cell(7).Value).Trim());
                                        myXml.WriteStartElement("ОбъемКоммунальныхУслуг");
                                        myXml.WriteElementString("ИндивидуальноеПотребление",
                                            Convert.ToString(wb.Worksheet(1).Row(t).Cell(9).Value).Trim());
                                        myXml.WriteElementString("ОбщедомовыеНужды",
                                            Convert.ToString(wb.Worksheet(1).Row(t).Cell(13).Value).Trim());
                                        myXml.WriteEndElement();
                                        myXml.WriteElementString("Тариф",
                                            Convert.ToString(wb.Worksheet(1).Row(t).Cell(15).Value).Trim());
                                        myXml.WriteStartElement("РазмерПлатыЗаКоммунальныеУслуги");
                                        myXml.WriteElementString("ИндивидуальноеПотребление",
                                            Convert.ToString(wb.Worksheet(1).Row(t).Cell(18).Value).Trim());
                                        myXml.WriteElementString("ОбщедомовыеНужды",
                                            Convert.ToString(wb.Worksheet(1).Row(t).Cell(21).Value).Trim());
                                        myXml.WriteEndElement();
                                        myXml.WriteElementString("ВсегоНачислено",
                                            Convert.ToString(wb.Worksheet(1).Row(t).Cell(23).Value).Trim());
                                        myXml.WriteElementString("Перерасчеты",
                                            Convert.ToString(wb.Worksheet(1).Row(t).Cell(25).Value).Trim());
                                        //myXml.WriteElementString("Льготы", Convert.ToString(wb.Worksheet(1).Row(t).Cell(27).Value).Trim());
                                        myXml.WriteStartElement("ИтогоКОплате");
                                        myXml.WriteElementString("Всего",
                                            Convert.ToString(wb.Worksheet(1).Row(t).Cell(27).Value).Trim());
                                        myXml.WriteStartElement("ЗаКоммунальныеУслуги");
                                        myXml.WriteElementString("ИндивидуальноеПотребление",
                                            Convert.ToString(wb.Worksheet(1).Row(t).Cell(30).Value).Trim());
                                        myXml.WriteElementString("ОбщедомовыеНужды",
                                            Convert.ToString(wb.Worksheet(1).Row(t).Cell(33).Value).Trim());
                                        if (Convert.ToString(wb.Worksheet(1).Row(t).Cell(33).Value).Trim() != "")
                                            domTotal +=
                                                Convert.ToDecimal(
                                                    Convert.ToString(wb.Worksheet(1).Row(t).Cell(33).Value).Trim());
                                        myXml.WriteEndElement();
                                        myXml.WriteEndElement();
                                        myXml.WriteEndElement();
                                    }
                                    if (t != i + 10)
                                        t++;
                                }
                                myXml.WriteStartElement("ИтогоКОплате");
                                myXml.WriteElementString("РазмерПлатыИндивид",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 33).Cell(18).Value).Trim());
                                myXml.WriteElementString("РазмерПлатыДом",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 33).Cell(21).Value).Trim());
                                myXml.WriteElementString("ВсегоНачислено",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 33).Cell(23).Value).Trim());
                                myXml.WriteElementString("Перерасчеты",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 33).Cell(25).Value).Trim());
                                myXml.WriteElementString("Всего",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 33).Cell(27).Value).Trim());
                                myXml.WriteElementString("ИтогоИндивид",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 33).Cell(30).Value));
                                myXml.WriteElementString("ИтогоДом",
                                    domTotal.ToString());
                                myXml.WriteElementString("Долг",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 34).Cell(42).Value));
                                myXml.WriteElementString("Оплачено",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 35).Cell(42).Value));
                                myXml.WriteElementString("Пени",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 36).Cell(42).Value));
                                myXml.WriteEndElement();
                                myXml.WriteEndElement();

                                #endregion Рвздел3

                                #region Раздел4

                                myXml.WriteStartElement("Раздел4");
                                for (int r = i + 10; r <= i + 31; r++)
                                {
                                    if (wb.Worksheet(1).Row(r).Cell(1).Value != null &&
                                        Convert.ToString(wb.Worksheet(1).Row(r).Cell(1).Value).Trim() != "")
                                    {
                                        myXml.WriteStartElement("СправочнаяИнформация");
                                        string serv = Convert.ToString(wb.Worksheet(1).Row(r).Cell(1).Value).Trim();
                                        myXml.WriteElementString("ВидУслуги", serv);
                                        //myXml.WriteStartElement("НормативПотребления");
                                        //myXml.WriteElementString("Индивидуальное", Convert.ToString(wb.Worksheet(1).Row(r).Cell(29).Value).Trim());
                                        //myXml.WriteElementString("Общедомовое", Convert.ToString(wb.Worksheet(1).Row(r).Cell(31).Value).Trim());
                                        //myXml.WriteEndElement();
                                        myXml.WriteStartElement("Показания");
                                        myXml.WriteElementString("Индивидуальные",
                                            Convert.ToString(wb.Worksheet(1).Row(r).Cell(36).Value).Trim());
                                        myXml.WriteElementString("Общедомовые",
                                            Convert.ToString(wb.Worksheet(1).Row(r).Cell(38).Value).Trim());
                                        myXml.WriteEndElement();
                                        myXml.WriteStartElement("ОбъемКоммунальныхУслуг");
                                        myXml.WriteElementString("ПомещенияДома",
                                            Convert.ToString(wb.Worksheet(1).Row(r).Cell(42).Value).Trim());
                                        myXml.WriteElementString("ОбщедомовыеНуждыДома",
                                            Convert.ToString(wb.Worksheet(1).Row(r).Cell(44).Value).Trim());
                                        myXml.WriteEndElement();
                                        myXml.WriteEndElement();
                                    }
                                    if (r != i + 10)
                                        r++;
                                }
                                myXml.WriteEndElement();

                                #endregion Рвздел4

                                #region Раздел5

                                myXml.WriteStartElement("Раздел5");
                                //for (int j = z; j <= z + 3; j++)
                                //{
                                //    colPart5 = 1;
                                //    for (int r = 2; r <= 18; r++)
                                //    {
                                //        if (Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim() != "" && Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim() != "Раздел 6")
                                //        {
                                //            if (Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim() == "Виды услуг" || Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim() == "Вид услуг"
                                //                || Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim().Contains("Рассрочка платежа")
                                //                || Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim().Contains("учетом рассрочки"))
                                //            {
                                //                if (j != z + 3)
                                //                    j++;
                                //                else
                                //                    break;
                                //            }
                                //            else
                                //            {
                                //                if (r >= 14)
                                //                    colPart5 = 3;
                                //                switch (colPart5)
                                //                {
                                //                    case 1:
                                //                        {
                                //                            myXml.WriteStartElement("Услуга");
                                //                            myXml.WriteElementString("ВидУслуги", Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim());
                                //                            break;
                                //                        }
                                //                    case 2:
                                //                        {
                                //                            myXml.WriteElementString("ОснованиеПерерасчета", Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim());
                                //                            break;
                                //                        }
                                //                    case 3:
                                //                        {
                                //                            myXml.WriteElementString("Сумма", Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim());
                                //                            myXml.WriteEndElement();
                                //                            break;
                                //                        }
                                //                }
                                //                colPart5++;
                                //            }
                                //        }
                                //    }
                                //}
                                myXml.WriteEndElement();

                                #endregion

                                #region Примечание

                                //string t1 = Convert.ToString(wb.Worksheet(1).Row(i + 36).Cell(12).Value).Trim();
                                myXml.WriteElementString("Примечание", "");

                                #endregion Прмечание
                            }
                        }
                    }
#endregion
                    #region Радужный элит
                    else if (isRaduga.Checked)
                    {
                        for (int i = 1; i <= 40000; i = i + 62)
                        {
                            if (
                                Convert.ToString(wb.Worksheet(1).Row(i).Cell(7).Value)
                                    .Contains("ПЛАТЕЖНЫЙ ДОКУМЕНТ (СЧЕТ)"))
                            {
                                if (i != 1)
                                    myXml.WriteEndElement();
                                myXml.WriteStartElement("ЛицевойСчет");

                                #region Раздел1

                                myXml.WriteStartElement("Раздел1");
                                myXml.WriteElementString("Период",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 7).Cell(1).Value)
                                        .Trim()
                                        .Split(separator11, StringSplitOptions.None)[1].Trim().Split(' ')[0] + " " +
                                    Convert.ToString(wb.Worksheet(1).Row(i + 7).Cell(1).Value)
                                        .Trim()
                                        .Split(separator11, StringSplitOptions.None)[1].Trim().Split(' ')[1]);

                                if (
                                    Convert.ToString(wb.Worksheet(1).Row(i + 7).Cell(1).Value)
                                        .ToUpper()
                                        .Contains("НЕ ПРИВАТИЗИРОВАНА"))
                                    myXml.WriteElementString("Приватизирована", "Нет");
                                else
                                    myXml.WriteElementString("Приватизирована", "Да");
                                myXml.WriteElementString("НомерЛС",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 10).Cell(37).Value)
                                        .Split(separatorLs, StringSplitOptions.None)[1].Trim());
                                myXml.WriteElementString("ФИО",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 8).Cell(1).Value)
                                        .Trim()
                                        .Split(separator1, StringSplitOptions.None)[1]
                                        .Split(separator2, StringSplitOptions.None)[0].Trim());
                                myXml.WriteElementString("АдресПомещения",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 8).Cell(1).Value)
                                        .Trim()
                                        .Split(separator2, StringSplitOptions.None)[1].Trim().
                                        Split(separator31, StringSplitOptions.None)[0]);

                                myXml.WriteElementString("ПлощадьПомещения",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 8).Cell(1).Value)
                                        .Split(separator31, StringSplitOptions.None)[1]
                                        .Split(separatorM, StringSplitOptions.None)[0].Trim());
                                myXml.WriteElementString("ПрописаноПроживает",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 8).Cell(1).Value)
                                        .Split(separator41, StringSplitOptions.None)[1]
                                        .Split(separatorP, StringSplitOptions.None)[0].Trim());
                                myXml.WriteElementString("ПлощадьДома",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 8).Cell(1).Value)
                                        .Split(separator51, StringSplitOptions.None)[1]
                                        .Split(separatorM, StringSplitOptions.None)[0].Trim());
                                myXml.WriteElementString("ПлощадьМОП",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 8).Cell(1).Value)
                                        .Split(separator6, StringSplitOptions.None)[1]
                                        .Split(separatorComma, StringSplitOptions.None)[0].Trim());
                                myXml.WriteElementString("Проживает",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 8).Cell(1).Value)
                                        .Split(separator711, StringSplitOptions.None)[1]
                                        .Split(separatorP, StringSplitOptions.None)[0].Trim());
                                myXml.WriteElementString("Организация",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 12).Cell(1).Value)
                                        .Split(separator8, StringSplitOptions.None)[1].Trim());
                                myXml.WriteEndElement();

                                #endregion Рвздел1

                                #region Раздел2

                                myXml.WriteStartElement("Раздел2");
                                string paymentTo = Convert.ToString(wb.Worksheet(1).Row(i + 9).Cell(11).Value).Trim();
                                myXml.WriteElementString("ПолучательПлатежа", paymentTo);
                                string bankInfo = Convert.ToString(wb.Worksheet(1).Row(i + 9).Cell(17).Value).Trim();
                                //string bankInfo =
                                    //"Р/с - 40702810754400005587   Кор/счет-30101810200000000607    БИК 043601607 Поволжский банк ОАО «Сбербанк России» г. Самара";
                                myXml.WriteElementString("БанковскийСчет", bankInfo.Trim());

                                myXml.WriteElementString("ПлатежныйКод",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 9).Cell(24).Value)
                                        .Trim()
                                        .Replace('*', ' ')
                                        .Trim());

                                myXml.WriteElementString("ВидПлаты",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 9).Cell(37).Value)
                                        .Trim()
                                        .Replace("(ООО «УК «Ассоциация Управляющих Компаний»)", " ")
                                        .Trim());

                                myXml.WriteElementString("СуммаКОплате",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 9).Cell(43).Value).Trim());

                                myXml.WriteEndElement();

                                #endregion Рвздел2

                                #region Раздел3

                                myXml.WriteStartElement("Раздел3");
                                for (int t = i + 25; t <= i + 38; t++)
                                {
                                    if (wb.Worksheet(1).Row(t).Cell(1).Value != null &&
                                        Convert.ToString(wb.Worksheet(1).Row(t).Cell(1).Value).Trim() != "")
                                    {
                                        myXml.WriteStartElement("Услуга");
                                        myXml.WriteElementString("ВидУслуги",
                                            Convert.ToString(wb.Worksheet(1).Row(t).Cell(1).Value).Trim());
                                        myXml.WriteElementString("ЕдиницаИзмерения",
                                            Convert.ToString(wb.Worksheet(1).Row(t).Cell(5).Value).Trim());
                                        myXml.WriteStartElement("ОбъемКоммунальныхУслуг");
                                        myXml.WriteElementString("ИндивидуальноеПотребление",
                                            Convert.ToString(wb.Worksheet(1).Row(t).Cell(6).Value).Trim());
                                        myXml.WriteElementString("ОбщедомовыеНужды",
                                            Convert.ToString(wb.Worksheet(1).Row(t).Cell(8).Value).Trim());
                                        myXml.WriteEndElement();
                                        myXml.WriteElementString("Тариф",
                                            Convert.ToString(wb.Worksheet(1).Row(t).Cell(9).Value).Trim());
                                        myXml.WriteStartElement("РазмерПлатыЗаКоммунальныеУслуги");
                                        myXml.WriteElementString("ИндивидуальноеПотребление",
                                            Convert.ToString(wb.Worksheet(1).Row(t).Cell(12).Value).Trim());
                                        myXml.WriteElementString("ОбщедомовыеНужды",
                                            Convert.ToString(wb.Worksheet(1).Row(t).Cell(16).Value).Trim());
                                        myXml.WriteEndElement();
                                        myXml.WriteElementString("ВсегоНачислено",
                                            Convert.ToString(wb.Worksheet(1).Row(t).Cell(18).Value).Trim());
                                        myXml.WriteElementString("Перерасчеты",
                                            Convert.ToString(wb.Worksheet(1).Row(t).Cell(20).Value).Trim());
                                        //myXml.WriteElementString("Льготы", Convert.ToString(wb.Worksheet(1).Row(t).Cell(27).Value).Trim());
                                        myXml.WriteStartElement("ИтогоКОплате");
                                        myXml.WriteElementString("Всего",
                                            Convert.ToString(wb.Worksheet(1).Row(t).Cell(22).Value).Trim());
                                        myXml.WriteStartElement("ЗаКоммунальныеУслуги");
                                        myXml.WriteElementString("ИндивидуальноеПотребление",
                                            Convert.ToString(wb.Worksheet(1).Row(t).Cell(26).Value).Trim());
                                        myXml.WriteElementString("ОбщедомовыеНужды",
                                            Convert.ToString(wb.Worksheet(1).Row(t).Cell(33).Value).Trim());
                                        myXml.WriteEndElement();
                                        myXml.WriteEndElement();
                                        myXml.WriteEndElement();
                                    }
                                }
                                myXml.WriteStartElement("ИтогоКОплате");
                                myXml.WriteElementString("РазмерПлатыИндивид",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 39).Cell(12).Value).Trim());
                                myXml.WriteElementString("РазмерПлатыДом",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 39).Cell(16).Value).Trim());
                                myXml.WriteElementString("ВсегоНачислено",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 39).Cell(18).Value).Trim());
                                myXml.WriteElementString("Перерасчеты",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 39).Cell(20).Value).Trim());
                                myXml.WriteElementString("Всего",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 39).Cell(22).Value).Trim());
                                myXml.WriteElementString("ИтогоИндивид",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 39).Cell(26).Value));
                                myXml.WriteElementString("ИтогоДом",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 39).Cell(33).Value));
                                myXml.WriteElementString("Долг",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 43).Cell(45).Value));
                                myXml.WriteElementString("Оплачено",
                                    Convert.ToString(wb.Worksheet(1).Row(i + 44).Cell(45).Value));
                                myXml.WriteEndElement();
                                myXml.WriteEndElement();

                                #endregion Рвздел3

                                #region Раздел4

                                myXml.WriteStartElement("Раздел4");
                                for (int r = i + 25; r <= i + 38; r++)
                                {
                                    if (wb.Worksheet(1).Row(r).Cell(1).Value != null &&
                                        Convert.ToString(wb.Worksheet(1).Row(r).Cell(1).Value).Trim() != "")
                                    {
                                        myXml.WriteStartElement("СправочнаяИнформация");
                                        string serv = Convert.ToString(wb.Worksheet(1).Row(r).Cell(1).Value).Trim();
                                        myXml.WriteElementString("ВидУслуги", serv);
                                        myXml.WriteStartElement("НормативПотребления");
                                        myXml.WriteElementString("Индивидуальное", Convert.ToString(wb.Worksheet(1).Row(r).Cell(40).Value).Trim());
                                        myXml.WriteElementString("Общедомовое", Convert.ToString(wb.Worksheet(1).Row(r).Cell(42).Value).Trim());
                                        myXml.WriteEndElement();
                                        myXml.WriteStartElement("Показания");
                                        myXml.WriteElementString("Индивидуальные",
                                            Convert.ToString(wb.Worksheet(1).Row(r).Cell(46).Value).Trim());
                                        myXml.WriteElementString("Общедомовые",
                                            Convert.ToString(wb.Worksheet(1).Row(r).Cell(48).Value).Trim());
                                        myXml.WriteEndElement();
                                        myXml.WriteStartElement("ОбъемКоммунальныхУслуг");
                                        myXml.WriteElementString("ПомещенияДома",
                                            Convert.ToString(wb.Worksheet(1).Row(r).Cell(55).Value).Trim());
                                        myXml.WriteElementString("ОбщедомовыеНуждыДома",
                                            Convert.ToString(wb.Worksheet(1).Row(r).Cell(57).Value).Trim());
                                        myXml.WriteEndElement();
                                        myXml.WriteEndElement();
                                    }
                                }
                                myXml.WriteEndElement();

                                #endregion Рвздел4

                                #region Раздел5

                                myXml.WriteStartElement("Раздел5");
                                //for (int j = z; j <= z + 3; j++)
                                //{
                                //    colPart5 = 1;
                                //    for (int r = 2; r <= 18; r++)
                                //    {
                                //        if (Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim() != "" && Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim() != "Раздел 6")
                                //        {
                                //            if (Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim() == "Виды услуг" || Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim() == "Вид услуг"
                                //                || Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim().Contains("Рассрочка платежа")
                                //                || Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim().Contains("учетом рассрочки"))
                                //            {
                                //                if (j != z + 3)
                                //                    j++;
                                //                else
                                //                    break;
                                //            }
                                //            else
                                //            {
                                //                if (r >= 14)
                                //                    colPart5 = 3;
                                //                switch (colPart5)
                                //                {
                                //                    case 1:
                                //                        {
                                //                            myXml.WriteStartElement("Услуга");
                                //                            myXml.WriteElementString("ВидУслуги", Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim());
                                //                            break;
                                //                        }
                                //                    case 2:
                                //                        {
                                //                            myXml.WriteElementString("ОснованиеПерерасчета", Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim());
                                //                            break;
                                //                        }
                                //                    case 3:
                                //                        {
                                //                            myXml.WriteElementString("Сумма", Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim());
                                //                            myXml.WriteEndElement();
                                //                            break;
                                //                        }
                                //                }
                                //                colPart5++;
                                //            }
                                //        }
                                //    }
                                //}
                                myXml.WriteEndElement();

                                #endregion

                                #region Примечание

                                string t1 = Convert.ToString(wb.Worksheet(1).Row(i + 60).Cell(1).Value).Trim();
                                myXml.WriteElementString("Примечание", t1);

                                #endregion Прмечание
                            }
                        }
                    }
#endregion
                    else
                    {
                        if (!isProseka.Checked)
                        {
                            for (int i = 1; i <= 40000; i = i + 42)
                            {
                                if (
                                    Convert.ToString(wb.Worksheet(1).Row(i).Cell(8).Value)
                                        .Contains("ПЛАТЕЖНЫЙ ДОКУМЕНТ (СЧЕТ)"))
                                {
                                    if (i != 1)
                                        myXml.WriteEndElement();
                                    myXml.WriteStartElement("ЛицевойСчет");

                                    #region Раздел1

                                    myXml.WriteStartElement("Раздел1");
                                    myXml.WriteElementString("Период",
                                        Convert.ToString(wb.Worksheet(1).Row(i + 3).Cell(1).Value)
                                            .Trim()
                                            .Split(separator11, StringSplitOptions.None)[1].Trim().Split(' ')[0] + " " +
                                        Convert.ToString(wb.Worksheet(1).Row(i + 3).Cell(1).Value)
                                            .Trim()
                                            .Split(separator11, StringSplitOptions.None)[1].Trim().Split(' ')[1]);

                                    if (
                                        Convert.ToString(wb.Worksheet(1).Row(i + 3).Cell(1).Value)
                                            .ToUpper()
                                            .Contains("НЕ ПРИВАТИЗИРОВАНА"))
                                        myXml.WriteElementString("Приватизирована", "Нет");
                                    else
                                        myXml.WriteElementString("Приватизирована", "Да");
                                    myXml.WriteElementString("НомерЛС",
                                        Convert.ToString(wb.Worksheet(1).Row(i + 3).Cell(1).Value)
                                            .Split(separator, StringSplitOptions.None)[1].Trim()
                                            .Split(separator1, StringSplitOptions.None)[0].Trim());
                                    myXml.WriteElementString("ФИО",
                                        Convert.ToString(wb.Worksheet(1).Row(i + 3).Cell(1).Value)
                                            .Trim()
                                            .Split(separator1, StringSplitOptions.None)[1]
                                            .Split(separator2, StringSplitOptions.None)[0].Trim());
                                    myXml.WriteElementString("АдресПомещения",
                                        Convert.ToString(wb.Worksheet(1).Row(i + 3).Cell(1).Value)
                                            .Trim()
                                            .Split(separator2, StringSplitOptions.None)[1].Trim().
                                            Split(separator31, StringSplitOptions.None)[0]);

                                    myXml.WriteElementString("ПлощадьПомещения",
                                        Convert.ToString(wb.Worksheet(1).Row(i + 3).Cell(1).Value)
                                            .Split(separator31, StringSplitOptions.None)[1]
                                            .Split(separatorM, StringSplitOptions.None)[0].Trim());
                                    myXml.WriteElementString("ПрописаноПроживает",
                                        Convert.ToString(wb.Worksheet(1).Row(i + 3).Cell(1).Value)
                                            .Split(separator41, StringSplitOptions.None)[1]
                                            .Split(separatorP, StringSplitOptions.None)[0].Trim());
                                    myXml.WriteElementString("ПлощадьДома",
                                        Convert.ToString(wb.Worksheet(1).Row(i + 3).Cell(1).Value)
                                            .Split(separator51, StringSplitOptions.None)[1]
                                            .Split(separatorM, StringSplitOptions.None)[0].Trim());
                                    myXml.WriteElementString("ПлощадьМОП",
                                        Convert.ToString(wb.Worksheet(1).Row(i + 3).Cell(1).Value)
                                            .Split(separator6, StringSplitOptions.None)[1]
                                            .Split(separatorM, StringSplitOptions.None)[0].Trim());
                                    myXml.WriteElementString("Проживает",
                                        Convert.ToString(wb.Worksheet(1).Row(i + 3).Cell(1).Value)
                                            .Split(separator71, StringSplitOptions.None)[1]
                                            .Split(separatorP, StringSplitOptions.None)[0].Trim());
                                    myXml.WriteElementString("Организация",
                                        Convert.ToString(wb.Worksheet(1).Row(i + 6).Cell(1).Value)
                                            .Split(separator8, StringSplitOptions.None)[1].Trim());
                                    myXml.WriteEndElement();

                                    #endregion Рвздел1

                                    #region Раздел2

                                    myXml.WriteStartElement("Раздел2");
                                    string paymentTo =
                                        Convert.ToString(wb.Worksheet(1).Row(i + 4).Cell(14).Value).Trim();
                                    myXml.WriteElementString("ПолучательПлатежа", paymentTo);
                                    string bankInfo = Convert.ToString(wb.Worksheet(1).Row(i + 4).Cell(18).Value).Trim();
                                    myXml.WriteElementString("БанковскийСчет", bankInfo.Trim());

                                    myXml.WriteElementString("ПлатежныйКод",
                                        Convert.ToString(wb.Worksheet(1).Row(i + 4).Cell(24).Value)
                                            .Trim()
                                            .Replace('*', ' ')
                                            .Trim());

                                    myXml.WriteElementString("ВидПлаты",
                                        Convert.ToString(wb.Worksheet(1).Row(i + 4).Cell(31).Value)
                                            .Trim()
                                            .Replace("(ООО «УК «Ассоциация Управляющих Компаний»)", " ")
                                            .Trim());

                                    myXml.WriteElementString("СуммаКОплате",
                                        Convert.ToString(wb.Worksheet(1).Row(i + 4).Cell(36).Value).Trim());

                                    myXml.WriteEndElement();

                                    #endregion Рвздел2

                                    #region Раздел3

                                    myXml.WriteStartElement("Раздел3");
                                    for (int t = i + 13; t <= i + 25; t++)
                                    {
                                        if (wb.Worksheet(1).Row(t).Cell(1).Value != null &&
                                            Convert.ToString(wb.Worksheet(1).Row(t).Cell(1).Value).Trim() != "")
                                        {
                                            myXml.WriteStartElement("Услуга");
                                            myXml.WriteElementString("ВидУслуги",
                                                Convert.ToString(wb.Worksheet(1).Row(t).Cell(1).Value).Trim());
                                            myXml.WriteElementString("ЕдиницаИзмерения",
                                                Convert.ToString(wb.Worksheet(1).Row(t).Cell(5).Value).Trim());
                                            myXml.WriteStartElement("ОбъемКоммунальныхУслуг");
                                            myXml.WriteElementString("ИндивидуальноеПотребление",
                                                Convert.ToString(wb.Worksheet(1).Row(t).Cell(7).Value).Trim().Replace('.',','));
                                            myXml.WriteElementString("ОбщедомовыеНужды",
                                                Convert.ToString(wb.Worksheet(1).Row(t).Cell(8).Value).Trim().Replace('.', ','));
                                            myXml.WriteEndElement();
                                            myXml.WriteElementString("Тариф",
                                                Convert.ToString(wb.Worksheet(1).Row(t).Cell(10).Value).Trim().Replace('.', ','));
                                            myXml.WriteStartElement("РазмерПлатыЗаКоммунальныеУслуги");
                                            myXml.WriteElementString("ИндивидуальноеПотребление",
                                                Convert.ToString(wb.Worksheet(1).Row(t).Cell(11).Value).Trim().Replace('.', ','));
                                            myXml.WriteElementString("ОбщедомовыеНужды",
                                                Convert.ToString(wb.Worksheet(1).Row(t).Cell(12).Value).Trim().Replace('.', ','));
                                            myXml.WriteEndElement();
                                            myXml.WriteElementString("ВсегоНачислено",
                                                Convert.ToString(wb.Worksheet(1).Row(t).Cell(13).Value).Trim().Replace('.', ','));
                                            myXml.WriteElementString("Перерасчеты",
                                                Convert.ToString(wb.Worksheet(1).Row(t).Cell(14).Value).Trim().Replace('.', ','));
                                            myXml.WriteElementString("Льготы",
                                                Convert.ToString(wb.Worksheet(1).Row(t).Cell(16).Value).Trim().Replace('.', ','));
                                            myXml.WriteStartElement("ИтогоКОплате");
                                            myXml.WriteElementString("Всего",
                                                Convert.ToString(wb.Worksheet(1).Row(t).Cell(17).Value).Trim().Replace('.', ','));
                                            myXml.WriteStartElement("ЗаКоммунальныеУслуги");
                                            myXml.WriteElementString("ИндивидуальноеПотребление",
                                                Convert.ToString(wb.Worksheet(1).Row(t).Cell(20).Value).Trim().Replace('.', ','));
                                            myXml.WriteElementString("ОбщедомовыеНужды",
                                                Convert.ToString(wb.Worksheet(1).Row(t).Cell(22).Value).Trim().Replace('.', ','));
                                            myXml.WriteEndElement();
                                            myXml.WriteEndElement();
                                            myXml.WriteEndElement();
                                        }
                                    }
                                    myXml.WriteStartElement("ИтогоКОплате");
                                    myXml.WriteElementString("РазмерПлатыИндивид",
                                        Convert.ToString(wb.Worksheet(1).Row(i + 26).Cell(11).Value).Trim().Replace('.', ','));
                                    myXml.WriteElementString("РазмерПлатыДом",
                                        Convert.ToString(wb.Worksheet(1).Row(i + 26).Cell(12).Value).Trim().Replace('.', ','));
                                    myXml.WriteElementString("ВсегоНачислено",
                                        Convert.ToString(wb.Worksheet(1).Row(i + 26).Cell(13).Value).Trim().Replace('.', ','));
                                    myXml.WriteElementString("Перерасчеты",
                                        Convert.ToString(wb.Worksheet(1).Row(i + 26).Cell(14).Value).Trim().Replace('.', ','));
                                    myXml.WriteElementString("Всего",
                                        Convert.ToString(wb.Worksheet(1).Row(i + 26).Cell(17).Value).Trim().Replace('.', ','));
                                    myXml.WriteElementString("ИтогоИндивид",
                                        Convert.ToString(wb.Worksheet(1).Row(i + 26).Cell(20).Value).Replace('.', ','));
                                    myXml.WriteElementString("ИтогоДом",
                                        Convert.ToString(wb.Worksheet(1).Row(i + 26).Cell(22).Value).Replace('.', ','));
                                    myXml.WriteElementString("Долг",
                                        Convert.ToString(wb.Worksheet(1).Row(i + 29).Cell(27).Value).Replace('.', ','));
                                    myXml.WriteElementString("Оплачено",
                                        Convert.ToString(wb.Worksheet(1).Row(i + 30).Cell(27).Value).Replace('.', ','));
                                    myXml.WriteEndElement();
                                    myXml.WriteEndElement();

                                    #endregion Рвздел3

                                    #region Раздел4

                                    myXml.WriteStartElement("Раздел4");
                                    for (int r = i + 13; r <= i + 25; r++)
                                    {
                                        if (wb.Worksheet(1).Row(r).Cell(1).Value != null &&
                                            Convert.ToString(wb.Worksheet(1).Row(r).Cell(1).Value).Trim() != "")
                                        {
                                            myXml.WriteStartElement("СправочнаяИнформация");
                                            string serv = Convert.ToString(wb.Worksheet(1).Row(r).Cell(1).Value).Trim();
                                            myXml.WriteElementString("ВидУслуги", serv);
                                            myXml.WriteStartElement("НормативПотребления");
                                            myXml.WriteElementString("Индивидуальное",
                                                Convert.ToString(wb.Worksheet(1).Row(r).Cell(23).Value).Trim().Replace('.', ','));
                                            myXml.WriteElementString("Общедомовое",
                                                Convert.ToString(wb.Worksheet(1).Row(r).Cell(25).Value).Trim().Replace('.', ','));
                                            myXml.WriteEndElement();
                                            myXml.WriteStartElement("Показания");
                                            myXml.WriteElementString("Индивидуальные",
                                                Convert.ToString(wb.Worksheet(1).Row(r).Cell(28).Value).Trim().Replace('.', ','));
                                            myXml.WriteElementString("Общедомовые",
                                                Convert.ToString(wb.Worksheet(1).Row(r).Cell(31).Value).Trim().Replace('.', ','));
                                            myXml.WriteEndElement();
                                            myXml.WriteStartElement("ОбъемКоммунальныхУслуг");
                                            myXml.WriteElementString("ПомещенияДома",
                                                Convert.ToString(wb.Worksheet(1).Row(r).Cell(34).Value).Trim().Replace('.', ','));
                                            myXml.WriteElementString("ОбщедомовыеНуждыДома",
                                                Convert.ToString(wb.Worksheet(1).Row(r).Cell(37).Value).Trim().Replace('.', ','));
                                            myXml.WriteEndElement();
                                            myXml.WriteEndElement();
                                        }
                                    }
                                    myXml.WriteEndElement();

                                    #endregion Рвздел4

                                    #region Раздел5

                                    myXml.WriteStartElement("Раздел5");
                                    //for (int j = z; j <= z + 3; j++)
                                    //{
                                    //    colPart5 = 1;
                                    //    for (int r = 2; r <= 18; r++)
                                    //    {
                                    //        if (Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim() != "" && Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim() != "Раздел 6")
                                    //        {
                                    //            if (Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim() == "Виды услуг" || Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim() == "Вид услуг"
                                    //                || Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim().Contains("Рассрочка платежа")
                                    //                || Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim().Contains("учетом рассрочки"))
                                    //            {
                                    //                if (j != z + 3)
                                    //                    j++;
                                    //                else
                                    //                    break;
                                    //            }
                                    //            else
                                    //            {
                                    //                if (r >= 14)
                                    //                    colPart5 = 3;
                                    //                switch (colPart5)
                                    //                {
                                    //                    case 1:
                                    //                        {
                                    //                            myXml.WriteStartElement("Услуга");
                                    //                            myXml.WriteElementString("ВидУслуги", Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim());
                                    //                            break;
                                    //                        }
                                    //                    case 2:
                                    //                        {
                                    //                            myXml.WriteElementString("ОснованиеПерерасчета", Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim());
                                    //                            break;
                                    //                        }
                                    //                    case 3:
                                    //                        {
                                    //                            myXml.WriteElementString("Сумма", Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim());
                                    //                            myXml.WriteEndElement();
                                    //                            break;
                                    //                        }
                                    //                }
                                    //                colPart5++;
                                    //            }
                                    //        }
                                    //    }
                                    //}
                                    myXml.WriteEndElement();

                                    #endregion

                                    #region Примечание

                                    //string t1 = Convert.ToString(wb.Worksheet(1).Row(i + 36).Cell(12).Value).Trim();
                                    myXml.WriteElementString("Примечание",
                                        Convert.ToString(wb.Worksheet(1).Row(i + 37).Cell(12).Value).Trim());

                                    #endregion Прмечание
                                }
                            }
                        }
                        else
                        {
                            for (int i = 1; i <= 40000; i = i + 71)
                            {
                                if (Convert.ToString(wb.Worksheet(1).Row(i).Cell(8).Value).Contains("ПЛАТЕЖНЫЙ ДОКУМЕНТ (СЧЕТ)"))
                                {
                                    if (i != 1)
                                        myXml.WriteEndElement();
                                    myXml.WriteStartElement("ЛицевойСчет");

                                    #region Раздел1
                                    myXml.WriteStartElement("Раздел1");
                                    myXml.WriteElementString("Период", Convert.ToString(wb.Worksheet(1).Row(i + 7).Cell(1).Value).Trim().Split(separator11, StringSplitOptions.None)[1].Trim().Split(' ')[0] + " " +
                                        Convert.ToString(wb.Worksheet(1).Row(i + 7).Cell(1).Value).Trim().Split(separator11, StringSplitOptions.None)[1].Trim().Split(' ')[1]);

                                    if (Convert.ToString(wb.Worksheet(1).Row(i + 7).Cell(1).Value).ToUpper().Contains("НЕ ПРИВАТИЗИРОВАНА"))
                                        myXml.WriteElementString("Приватизирована", "Нет");
                                    else
                                        myXml.WriteElementString("Приватизирована", "Да");
                                    myXml.WriteElementString("НомерЛС", Convert.ToString(wb.Worksheet(1).Row(i + 7).Cell(1).Value).Split(separator, StringSplitOptions.None)[1].Trim()
                                        .Split(separator1, StringSplitOptions.None)[0].Trim());
                                    myXml.WriteElementString("ФИО", Convert.ToString(wb.Worksheet(1).Row(i + 7).Cell(1).Value).Trim().Split(separator1, StringSplitOptions.None)[1]
                                        .Split(separator2, StringSplitOptions.None)[0].Trim());
                                    myXml.WriteElementString("АдресПомещения", Convert.ToString(wb.Worksheet(1).Row(i + 7).Cell(1).Value).Trim().Split(separator2, StringSplitOptions.None)[1].Trim().
                                        Split(separator3, StringSplitOptions.None)[0]);

                                    myXml.WriteElementString("ПлощадьПомещения", Convert.ToString(wb.Worksheet(1).Row(i + 7).Cell(1).Value).Split(separator3, StringSplitOptions.None)[1]
                                        .Split(separatorM, StringSplitOptions.None)[0].Trim());
                                    myXml.WriteElementString("ПрописаноПроживает", Convert.ToString(wb.Worksheet(1).Row(i + 7).Cell(1).Value).Split(separator41, StringSplitOptions.None)[1]
                                        .Split(separatorP, StringSplitOptions.None)[0].Trim());
                                    myXml.WriteElementString("ПлощадьДома", Convert.ToString(wb.Worksheet(1).Row(i + 7).Cell(1).Value).Split(separator5, StringSplitOptions.None)[1]
                                        .Split(separatorM, StringSplitOptions.None)[0].Trim());
                                    myXml.WriteElementString("ПлощадьМОП", Convert.ToString(wb.Worksheet(1).Row(i + 7).Cell(1).Value).Split(separator6, StringSplitOptions.None)[1]
                                        .Split(separatorM, StringSplitOptions.None)[0].Trim());
                                    myXml.WriteElementString("Проживает", Convert.ToString(wb.Worksheet(1).Row(i + 7).Cell(1).Value).Split(separator71, StringSplitOptions.None)[1]
                                        .Split(separatorP, StringSplitOptions.None)[0].Trim());
                                    myXml.WriteElementString("Организация", Convert.ToString(wb.Worksheet(1).Row(i + 9).Cell(1).Value).Split(separator8, StringSplitOptions.None)[1].Trim());
                                    myXml.WriteEndElement();
                                    #endregion Рвздел1

                                    #region Раздел2
                                    myXml.WriteStartElement("Раздел2");
                                    string paymentTo = Convert.ToString(wb.Worksheet(1).Row(i + 8).Cell(20).Value).Trim();
                                    myXml.WriteElementString("ПолучательПлатежа", paymentTo);
                                    string bankInfo = Convert.ToString(wb.Worksheet(1).Row(i + 8).Cell(26).Value).Trim();
                                    myXml.WriteElementString("БанковскийСчет", bankInfo.Trim());

                                    myXml.WriteElementString("ПлатежныйКод", Convert.ToString(wb.Worksheet(1).Row(i + 8).Cell(37).Value).Trim().Replace('*', ' ').Trim());

                                    myXml.WriteElementString("ВидПлаты", Convert.ToString(wb.Worksheet(1).Row(i + 8).Cell(45).Value).Trim().Replace("(ООО «УК «Ассоциация Управляющих Компаний»)", " ").Trim());

                                    myXml.WriteElementString("СуммаКОплате", Convert.ToString(wb.Worksheet(1).Row(i + 8).Cell(51).Value).Trim());

                                    myXml.WriteEndElement();
                                    #endregion Рвздел2

                                    #region Раздел3
                                    myXml.WriteStartElement("Раздел3");
                                    for (int t = i + 22; t <= i + 49; t++)
                                    {
                                        if (wb.Worksheet(1).Row(t).Cell(1).Value != null && Convert.ToString(wb.Worksheet(1).Row(t).Cell(1).Value).Trim() != "")
                                        {
                                            myXml.WriteStartElement("Услуга");
                                            myXml.WriteElementString("ВидУслуги", Convert.ToString(wb.Worksheet(1).Row(t).Cell(1).Value).Trim());
                                            myXml.WriteElementString("ЕдиницаИзмерения", Convert.ToString(wb.Worksheet(1).Row(t).Cell(5).Value).Trim());
                                            myXml.WriteStartElement("ОбъемКоммунальныхУслуг");
                                            myXml.WriteElementString("ИндивидуальноеПотребление", Convert.ToString(wb.Worksheet(1).Row(t).Cell(7).Value).Trim());
                                            myXml.WriteElementString("ОбщедомовыеНужды", Convert.ToString(wb.Worksheet(1).Row(t).Cell(9).Value).Trim());
                                            myXml.WriteEndElement();
                                            myXml.WriteElementString("Тариф", Convert.ToString(wb.Worksheet(1).Row(t).Cell(11).Value).Trim());
                                            myXml.WriteStartElement("РазмерПлатыЗаКоммунальныеУслуги");
                                            myXml.WriteElementString("ИндивидуальноеПотребление", Convert.ToString(wb.Worksheet(1).Row(t).Cell(13).Value).Trim());
                                            myXml.WriteElementString("ОбщедомовыеНужды", Convert.ToString(wb.Worksheet(1).Row(t).Cell(17).Value).Trim());
                                            myXml.WriteEndElement();
                                            myXml.WriteElementString("ВсегоНачислено", Convert.ToString(wb.Worksheet(1).Row(t).Cell(18).Value).Trim());
                                            myXml.WriteElementString("Перерасчеты", Convert.ToString(wb.Worksheet(1).Row(t).Cell(22).Value).Trim());
                                            //myXml.WriteElementString("Льготы", Convert.ToString(wb.Worksheet(1).Row(t).Cell(18).Value).Trim());
                                            myXml.WriteStartElement("ИтогоКОплате");
                                            myXml.WriteElementString("Всего", Convert.ToString(wb.Worksheet(1).Row(t).Cell(25).Value).Trim());
                                            myXml.WriteStartElement("ЗаКоммунальныеУслуги");
                                            myXml.WriteElementString("ИндивидуальноеПотребление", Convert.ToString(wb.Worksheet(1).Row(t).Cell(29).Value).Trim());
                                            myXml.WriteElementString("ОбщедомовыеНужды", Convert.ToString(wb.Worksheet(1).Row(t).Cell(32).Value).Trim());
                                            myXml.WriteEndElement();
                                            myXml.WriteEndElement();
                                            myXml.WriteEndElement();
                                        }
                                        if (t <= i + 46)
                                            t++;
                                    }
                                    myXml.WriteStartElement("ИтогоКОплате");
                                    myXml.WriteElementString("РазмерПлатыИндивид", Convert.ToString(wb.Worksheet(1).Row(i + 51).Cell(13).Value).Trim());
                                    myXml.WriteElementString("РазмерПлатыДом", Convert.ToString(wb.Worksheet(1).Row(i + 51).Cell(17).Value).Trim());
                                    myXml.WriteElementString("ВсегоНачислено", Convert.ToString(wb.Worksheet(1).Row(i + 51).Cell(18).Value).Trim());
                                    myXml.WriteElementString("Перерасчеты", Convert.ToString(wb.Worksheet(1).Row(i + 51).Cell(22).Value).Trim());
                                    myXml.WriteElementString("Всего", Convert.ToString(wb.Worksheet(1).Row(i + 51).Cell(25).Value).Trim());
                                    myXml.WriteElementString("ИтогоИндивид", Convert.ToString(wb.Worksheet(1).Row(i + 51).Cell(29).Value));
                                    myXml.WriteElementString("ИтогоДом", Convert.ToString(wb.Worksheet(1).Row(i + 51).Cell(32).Value));
                                    myXml.WriteElementString("Долг", Convert.ToString(wb.Worksheet(1).Row(i + 56).Cell(41).Value));
                                    myXml.WriteElementString("Оплачено", Convert.ToString(wb.Worksheet(1).Row(i + 57).Cell(41).Value));
                                    myXml.WriteEndElement();
                                    myXml.WriteEndElement();

                                    

                                    #endregion Рвздел3

                                    #region Раздел4
                                    myXml.WriteStartElement("Раздел4");
                                    for (int r = i + 22; r <= i + 49; r++)
                                    {
                                        if (wb.Worksheet(1).Row(r).Cell(1).Value != null && Convert.ToString(wb.Worksheet(1).Row(r).Cell(1).Value).Trim() != "")
                                        {
                                            myXml.WriteStartElement("СправочнаяИнформация");
                                            string serv = Convert.ToString(wb.Worksheet(1).Row(r).Cell(1).Value).Trim();
                                            myXml.WriteElementString("ВидУслуги", serv);
                                            myXml.WriteStartElement("НормативПотребления");
                                            myXml.WriteElementString("Индивидуальное", Convert.ToString(wb.Worksheet(1).Row(r).Cell(35).Value).Trim());
                                            myXml.WriteElementString("Общедомовое", Convert.ToString(wb.Worksheet(1).Row(r).Cell(38).Value).Trim());
                                            myXml.WriteEndElement();
                                            myXml.WriteStartElement("Показания");
                                            myXml.WriteElementString("Индивидуальные", Convert.ToString(wb.Worksheet(1).Row(r).Cell(40).Value).Trim());
                                            myXml.WriteElementString("Общедомовые", Convert.ToString(wb.Worksheet(1).Row(r).Cell(43).Value).Trim());
                                            myXml.WriteEndElement();
                                            myXml.WriteStartElement("ОбъемКоммунальныхУслуг");
                                            myXml.WriteElementString("ПомещенияДома", Convert.ToString(wb.Worksheet(1).Row(r).Cell(48).Value).Trim());
                                            myXml.WriteElementString("ОбщедомовыеНуждыДома", Convert.ToString(wb.Worksheet(1).Row(r).Cell(52).Value).Trim());
                                            myXml.WriteEndElement();
                                            myXml.WriteEndElement();
                                        }
                                        if (r <= i + 46)
                                            r++;
                                    }
                                    myXml.WriteEndElement();
                                    #endregion Рвздел4

                                    #region Раздел5
                                    myXml.WriteStartElement("Раздел5");
                                    //for (int j = z; j <= z + 3; j++)
                                    //{
                                    //    colPart5 = 1;
                                    //    for (int r = 2; r <= 18; r++)
                                    //    {
                                    //        if (Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim() != "" && Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim() != "Раздел 6")
                                    //        {
                                    //            if (Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim() == "Виды услуг" || Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim() == "Вид услуг"
                                    //                || Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim().Contains("Рассрочка платежа")
                                    //                || Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim().Contains("учетом рассрочки"))
                                    //            {
                                    //                if (j != z + 3)
                                    //                    j++;
                                    //                else
                                    //                    break;
                                    //            }
                                    //            else
                                    //            {
                                    //                if (r >= 14)
                                    //                    colPart5 = 3;
                                    //                switch (colPart5)
                                    //                {
                                    //                    case 1:
                                    //                        {
                                    //                            myXml.WriteStartElement("Услуга");
                                    //                            myXml.WriteElementString("ВидУслуги", Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim());
                                    //                            break;
                                    //                        }
                                    //                    case 2:
                                    //                        {
                                    //                            myXml.WriteElementString("ОснованиеПерерасчета", Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim());
                                    //                            break;
                                    //                        }
                                    //                    case 3:
                                    //                        {
                                    //                            myXml.WriteElementString("Сумма", Convert.ToString(wb.Worksheet(1).Row(j).Cell(r).Value).Trim());
                                    //                            myXml.WriteEndElement();
                                    //                            break;
                                    //                        }
                                    //                }
                                    //                colPart5++;
                                    //            }
                                    //        }
                                    //    }
                                    //}
                                    myXml.WriteEndElement();
                                    #endregion

                                    #region Примечание

                                    //string t1 = Convert.ToString(wb.Worksheet(1).Row(i + 62).Cell(15).Value).Trim();
                                    myXml.WriteElementString("Примечание", Convert.ToString(wb.Worksheet(1).Row(i + 62).Cell(16).Value).Trim());
                                    #endregion Прмечание
                                }
                            }
                        }
                    }
                    
                    myXml.WriteEndElement();
                    myXml.WriteEndElement();
                    myXml.Flush();
                    myXml.Close();
                }
                catch
                {
                    myXml.Flush();
                    myXml.Close();
                }
                
                //errorRow.WriteLine("Не найден ЛС либо найдено больше 1-го ЛС. Строка = " + (i + 2).ToString() + "|ФИО = " + dt1.Rows[i]["fio"].ToString() + "|Адресс = " + dt1.Rows[i]["address"].ToString());
                
                //wb.Save();
                errorRow.Close();
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = @"C:\Temp\7-Zip\7z.exe";
                string targetCompressName = @"C:\Temp\" + fileNameShort + ".zip";
                string filetozip = "\"" + savePath + fileNameShort + ".xml" + " " + "\"" + savePath + @"error.txt" + " ";
                startInfo.Arguments = "a -tzip \"" + targetCompressName + "\" \"" + filetozip + "\" -mx=9";
                startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                Process x = Process.Start(startInfo);
                x.WaitForExit();

                string path = targetCompressName;
                byte[] bts = File.ReadAllBytes(path);
                File.Delete(targetCompressName);
                Response.Clear();
                Response.ClearHeaders();
                Response.AddHeader("Content-Type", "Application/octet-stream");
                Response.AddHeader("Content-Length", bts.Length.ToString());
                Response.AddHeader("Content-Disposition", "attachment; filename=" + fileNameShort + ".zip");
                Response.BinaryWrite(bts);
                File.Delete(targetCompressName);
                File.Delete(savePath + fileName);
                File.Delete(savePath + fileNameShort + ".xml");
                Response.Flush();
                Response.End();

                

                Label1.Text = "Файл создан";
            }
            else
            {
                // Notify the user that a file was not uploaded.
                Label1.Text = "Необходимо загрузить фаил в формате .xlsx";
            }
        }

        protected void Button4_Click(object sender, EventArgs e)
        {
            Response.Redirect("http://85.140.61.250/GkhService/Service1.svc/GetDataCSV?token=f8f84d10cc6727b20becb7c5e85de047");
        }
    }
}