using LIMSClasses.Objects;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using WatersInterop.Controls;
using WIForms.Controls;

namespace WatersInterop
{
    public partial class EnterEqpForm : EnterDataFormClass
    {
        public EnterEqpForm(ExcelInteractClass ei, Microsoft.Office.Interop.Excel.Worksheet ws): base(ei)
        {
            InitializeComponent();
            WatersInterop wi = WatersInterop.GetWi();
            labCBox.Items.AddRange(Labs.GetAllLabsForCurUserAsArray());
            classCBox.Items.Clear();
            classCBox.Items.AddRange( wi.GetListItems("Классы оборудования"));            
            chkIntNBox.Value = 12;
        }
        private new class DatePicker : Control
        {
            private DateTimePicker vsdt;
            private string key;
            public DatePicker(DateTimePicker dt, string key)
            {
                vsdt = dt;
                this.key = key;
            }

            public override string Text
            {
                get
                {
                    if ((key.ToLower().Contains("время") && !vsdt.CustomFormat.Contains("HH"))
                        ||
                        (key.ToLower().Contains("дата") && !vsdt.CustomFormat.Contains("dd.MM")))
                    {
                        return "";
                    }
                    return vsdt.Value.ToString(vsdt.CustomFormat);
                }
                set
                {
                    vsdt.Value = SamplePlaceSelection.GetDefaultDatevalue(value, true);
                }
            }
        }
        public static string[] GetFieldNames()
        {
            return new string[] { 
            "Лаборатория",  
            "Дата прихода",
            "Дата выпуска",
            "Заводской номер",
            "Инвентарный номер",
            "Класс",  
            "Цел. упак., вн. вид, комплектность",
            "Нал. пасп., инстр. по эксп.",
            "Завод изготовитель",
            "Наименование",
            "Тип оборудования",  
            "Дата поверки",
            "Межповерочный интервал",
            "Дата окончания аттестации",
            "Дата отправки на поверку",
            "Дата возвращёния с поверки",  
            "Дата техобслуживания",
            "Состояние",
            "Дата списания с бухгалтерского учёта",
            "Ф.И.О. принявшего",
            "Номер документа",
            "Дата ввода в эксплуатацию",
            "Наименование характеристики",
            "Наименование объекта",
            "Диапазон измерений",
            "Класс точности",
            "Право собственности",
            "Место установки",
            "Номер протокола",
            "Дата по протоколу",
            "Стоимость",
            "Страна изготовитель"
            };
        }
        protected override string GetCurrentID()
        {
            return ei.GetWatersInterop().GetCurrentExperimentId() ;
        }

        protected override Dictionary<string, string> GetDefValues( Microsoft.Office.Interop.Excel.Worksheet ws)
        {
            if (ws != null)
            {
                Dictionary<string, string> defvalues = GetDefValuesByNames(ws, EnterEqpForm.GetFieldNames());
                return defvalues;
            }
            else
            {
                return new Dictionary<string, string>();
            }
        }
        public override Dictionary<string, Control> GetFieldNamesControls()
        {
            string[] fieldnames = GetFieldNames();
            Dictionary<string, Control> vs = new Dictionary<string, Control>(fieldnames.Length);
            vs.Add(fieldnames[0], labCBox);
            vs.Add(fieldnames[1], new DatePicker(dateInDP, fieldnames[1]));
            vs.Add(fieldnames[2], new DatePicker(dateMadeDP, fieldnames[2]));
            vs.Add(fieldnames[3], factNumBox);
            vs.Add(fieldnames[4], invNumBox);
            vs.Add(fieldnames[5], classCBox);
            vs.Add(fieldnames[6], packCBox);
            vs.Add(fieldnames[7], manCBox);
            vs.Add(fieldnames[8], factBox);
            vs.Add(fieldnames[9], nameBox);
            vs.Add(fieldnames[10], typeBox);
            vs.Add(fieldnames[11], new DatePicker(dateCheckDP, fieldnames[11]));
            vs.Add(fieldnames[12], chkIntNBox);
            vs.Add(fieldnames[13], new DatePicker(dateEndChkDP, fieldnames[13]));
            vs.Add(fieldnames[14], new DatePicker(dateSandToCkDP, fieldnames[14]));
            vs.Add(fieldnames[15], new DatePicker(dateBackFrCkDP, fieldnames[15]));
            vs.Add(fieldnames[16], new DatePicker(dateTechDP, fieldnames[16]));
            vs.Add(fieldnames[17], condBox);
            vs.Add(fieldnames[18], new DatePicker(dateWOffDP, fieldnames[18]));      
            vs.Add(fieldnames[19], fioBox);
            vs.Add(fieldnames[20], docNumBox);
            vs.Add(fieldnames[21], new DatePicker(DateInUsePicker, fieldnames[21]));
            vs.Add(fieldnames[22], NameTestBox);
            vs.Add(fieldnames[23], ProductBox);
            vs.Add(fieldnames[24], DiapBox);
            vs.Add(fieldnames[25], ClassTochnBox);
            vs.Add(fieldnames[26], PravSobBox);
            vs.Add(fieldnames[27], LocationBox);
            vs.Add(fieldnames[28], ProtocolBox);
            vs.Add(fieldnames[29], new DatePicker(dateProtocol, fieldnames[29]));
            vs.Add(fieldnames[30], CostBox);
            vs.Add(fieldnames[31], CountryBox);
            return vs;
        }

        private void classCBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (classCBox.Text)
            {
                case "ИО":
                    label10.Text = "Дата аттестации";
                    label11.Text = "Дата окончания действия аттестации";
                    label14.Text = "Межаттестационный интервал";
                    label21.Visible = true;
                    docNumBox.Visible = true;
                    label21.Text = "Номер аттестата";
                    NameTestLabel.Visible = true;
                    NameTestLabel.Text = "Наименование характеристики/испытания";
                    label22.Visible = true;
                    label22.Text = "Наименование объекта";
                    ProductBox.Visible = true;
                    ClassTochnLabel.Visible = false;
                    DiapBox.Visible = true;
                    ClassTochnBox.Visible = false;
                    DiapLabel.Visible = true;
                    DiapLabel.Text = "Основные технические характеристики";
                    label23.Visible = true;
                    label23.Text = "Номер протокола";
                    ProtocolBox.Visible = true;
                    label24.Visible = true;
                    dateProtocol.Visible = true;
                    break;

                case "СИ":
                    label10.Text = "Дата поверки";
                    label11.Text = "Дата окончания действия поверки";
                    label14.Text = "Межповерочный интервал";
                    label21.Visible = true;
                    NameTestLabel.Visible = true;
                    NameTestLabel.Text = "Наименование характеристики";
                    label22.Visible = false;
                    ProductBox.Visible = false;
                    ClassTochnLabel.Visible = true;
                    DiapBox.Visible = true;
                    ClassTochnLabel.Text = "Класс точности";
                    ClassTochnBox.Visible = true;
                    DiapLabel.Visible = true;
                    DiapLabel.Text = "Диапазон измерений";
                    docNumBox.Visible = true;
                    label21.Text = "Номер свидетельства";
                    label23.Visible = false;
                    ProtocolBox.Visible = false;
                    label24.Visible = false;
                    dateProtocol.Visible = false;
                    break;
                default:
                    label10.Text = "Дата проверки";
                    label11.Text = "Дата окончания действия проверки";
                    label14.Text = "Межпроверочный интервал";
                    label21.Visible = false;
                    docNumBox.Visible = false;
                    label22.Visible = false;
                    NameTestLabel.Visible = true;
                    NameTestLabel.Text = "Назначение";
                    ProductBox.Visible = false;
                    ClassTochnLabel.Visible = false;
                    ClassTochnBox.Visible = false;
                    DiapBox.Visible = false;
                    DiapLabel.Visible = false;
                    label23.Visible = false;
                    ProtocolBox.Visible = false;
                    label24.Visible = false;
                    dateProtocol.Visible = false;
                    break;
            }
        }



        private void dateCheckDP_ValueChanged(object sender, EventArgs e)
        {
            dateEndChkDP.MinDate = dateCheckDP.Value;
        }

        private void chkIntNBox_ValueChanged(object sender, EventArgs e)
        {
            DateTime newDate = dateCheckDP.Value.AddMonths((int)chkIntNBox.Value);
            if (dateEndChkDP.Value != newDate)
            {
                dateEndChkDP.Value = newDate;
            }

        }

        private void dateEndChkDP_ValueChanged(object sender, EventArgs e)
        {
            //DateTime dateB = dateCheckDP.Value;
            //DateTime dateE = dateEndChkDP.Value;
            TimeSpan times = dateEndChkDP.Value.Subtract(dateCheckDP.Value);
                int i = times.Days / 30;
                if (chkIntNBox.Value != i)
                {
                    chkIntNBox.Value = i;
                }
            ;
        }

        protected override void SetIdIntoName(ExcelInteractClass ei, Microsoft.Office.Interop.Excel.Worksheet ws, Dictionary<string, object> evals)
        {
            
        }

        protected override void ClearIdData(ExcelInteractClass ei, Microsoft.Office.Interop.Excel.Worksheet newws)
        {
            
        }

        public static void EnterEqpDataToSheet(ExcelInteractClass ei, Microsoft.Office.Interop.Excel.Worksheet ws)
        {

            EnterEqpForm eqf = new EnterEqpForm(ei, ws);
            eqf.EnterDataToSheet( ws);
        }
    }
}