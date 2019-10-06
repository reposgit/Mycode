using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using WatersInterop.Forms;
using System.Globalization;
using Utils;
using Utils.VLK;
using LIMSClasses.Interfaces;
using LIMSClasses.Objects;
using DAO.Interfaces;
using System.Linq;

namespace WatersInterop
{
    [ComVisible(true)]
    public class Calculations
    {
        private ExcelInteractClass ei;
        private WatersInterop wi;

        private static Calculations calc = null;

        public static Calculations GetCalculations(ExcelInteractClass ei)
        {
            if (calc == null)
            {
                calc = new Calculations();
            }
            calc.ei = ei;
            calc.wi = ei.GetWatersInterop();
            return calc;
        }

        /// <summary>
        /// Функция возвращает элемент массива который был создан путем сплита по заданному разделютелю
        /// </summary>
        /// <param name="str">Входная строка</param>
        /// <param name="index">Индекс элемента массива который необходимо вернуть</param>
        /// <param name="delimeter">Разделитель</param>
        /// <returns></returns>
        public static string GetWord(string str, int index, string delimeter)
        {
            return CalcFunctions.GetWord(str, index, delimeter);
        }

        [ComVisible(true)]
        public object CalcMinConcentration(string docid, double x)
        {
            return CalcFunctions.CalcMinConcentration(docid, x);
        }

        [ComVisible(true)]
        public void AddNameToActiveCell()
        {
            AddNameForm.ShowDialog(ei);
        }

        [ComVisible(true)]
        public void ShowLastRow()
        {
            DebugModule.ShowLastRow(ei.GetActiveSheet());
        }
        [ComVisible(true)]
        public void ShowLastCol()
        {
            DebugModule.ShowLastCol(ei.GetActiveSheet());
        }
        [ComVisible(true)]
        public void ShowAllSheets()
        {
            DebugModule.ShowAllSheets(ei.GetApplication());
        }
        [ComVisible(true)]
        public void HideAllShsExceptCur()
        {
            DebugModule.HideAllShsExceptCur(ei.GetApplication());
        }

        [ComVisible(true)]
        public string MyRound(string n)
        {
            return Utils.OkrFunctions.MyRound(n);
        }
        [ComVisible(true)]
        public string GetDateOfAnalysis(string batchno)
        {
            return BatchCountReport.GetDateOfAnalisys(wi, wi.GetCurrentUserId(), batchno);
        }
        [ComVisible(true)]
        public string RoundNeopr(string n)
        {
            return Utils.OkrFunctions.RoundNeopr(n);
        }
        [ComVisible(true)]
        public string RoundNeoprGOSTGas(string n)
        {
            return Utils.OkrFunctions.RoundNeoprGOSTGas(n);
        }
        [ComVisible(true)]
        public string OkrToZnach(string n, int nznach)
        {
            return Utils.OkrFunctions.OkrToZnach(n, nznach);
        }
        [ComVisible(true)]
        public string BatchSamplingDate(string docid, string batchno)
        {
            return ReportClass.BatchSamplingDate(wi, docid, batchno);
        }
        [ComVisible(true)]
        public string BatchSamplingDateTime(string docid, string batchno, bool includeTime)
        {
            if (GetBatchMetaData(docid,batchno,"Наименование пробы") != "Точечная") { includeTime = false; }
            return ReportClass.BatchSamplingDateTime(wi, docid, batchno, includeTime);
        }
        [ComVisible(true)]
        public string BatchSamplingDateTime(string startdate, string enddate, string timestring, bool includeTime)
        {
            return ReportClass.BatchSamplingDateTime(startdate, enddate, timestring, includeTime);
        }
        [ComVisible(true)]
        public string BatchSamplingDateTimeByTable(string docid, string batchno, bool includeTime, string table)
        {
            return ReportClass.BatchSamplingDateTimeByTable(wi, docid, batchno, includeTime, table);
        }
        [ComVisible(true)]
        public string[] GetStringsFromRange(Microsoft.Office.Interop.Excel.Range strings)
        {
            return ReportClass.GetStringsFromRange(strings);
        }
        [ComVisible(true)]
        public string MaxLenStr(Microsoft.Office.Interop.Excel.Range strings)
        {
            return ReportClass.MaxLenStr(strings);
        }
        [ComVisible(true)]
        public string MaxLenStr(string[] strs)
        {
            return ReportClass.MaxLenStr(strs);
        }
        [ComVisible(true)]
        public string SumDistinctStringsRange(Microsoft.Office.Interop.Excel.Range strings, string delimeter)
        {
            return ReportClass.SumDistinctStringsRange(ei, strings, delimeter);
        }
        [ComVisible(true)]
        public string GetDateInterval(Microsoft.Office.Interop.Excel.Range r)
        {
            return ReportClass.GetDateInterval(r);
        }
        [ComVisible(true)]
        public string GetDateIntervalMinMinutes(Microsoft.Office.Interop.Excel.Range r, double minutes)
        {
            return ReportClass.GetDateIntervalMinMinutes(r, minutes);
        }
        [ComVisible(true)]
        public string GetDateInterval(string[] strs)
        {
            return ReportClass.GetDateInterval(strs, TimeSpan.FromSeconds(0));
        }
        [ComVisible(true)]
        public object FirstNonEmptyUpperRow(Microsoft.Office.Interop.Excel.Range cell)
        {
            return ReportClass.FirstNonEmptyUpperRow(cell);
        }
        [ComVisible(true)]
        public string Nervo(double r, double rk, string strformat)
        {
            return ReportClass.Nervo(r, rk, strformat);
        }
        [ComVisible(true)]
        public string DateToWord(string date)
        {
            return ReportClass.DateToWord(date);
        }
        [ComVisible(true)]
        public string AnalyzedObject(string batchno)
        {
            return ReportClass.AnalyzedObject(batchno);
        }
        [ComVisible(true)]
        public string ConvertDate(string src)
        {
            return ReportClass.ConvertDate(src);
        }
        [ComVisible(true)]
        public string GetResultString(string resvalue, string pogrvalue, string precstring)
        {
            return ReportClass.GetResultString(resvalue, pogrvalue, precstring);
        }
        [ComVisible(true)]
        public object GetResultDigit(string resvalue, string pogrvalue, string precstring)
        {
            return ReportClass.GetResultDigit(resvalue, pogrvalue, precstring);
        }
        [ComVisible(true)]
        public string GetResultWithPogr(string resvalue, string pogrvalue, string precstring)
        {
            return ReportClass.GetResultWithPogr(resvalue, pogrvalue, precstring);
        }
        [ComVisible(true)]
        public string GetMethodprecStr(Microsoft.Office.Interop.Excel.Range r, string methodname, string testid, string lab, double c)
        {
            return VLKModule.GetMethodprecStr(r, "", methodname, testid, lab, "", c);
            //return VLKModule.GetMethodprecLabStr(wi, r, methodname, testid, lab, c);
        }
        [ComVisible(true)]
        public string GetMethodprec(Microsoft.Office.Interop.Excel.Range r, string methodname, string testid, string lab, double c)
        {
            return VLKModule.GetMethodprec(r, "", methodname, testid, lab, "", c).ToString();
            //return VLKModule.GetMethodprecLabStr(wi, r, methodname, testid, lab, c);
        }
        [ComVisible(true)]
        public string GetBatchMetaData(string docid, string batchno, string key)
        {
            return CalcFunctions.GetBatchMetaData(docid, batchno, key);
        }
        /// <summary>
        /// Метод получения инструментов и их параметров (Аттестован до, заводской номер)
        /// </summary>
        /// <param name="docid"></param>
        /// <param name="batchno"></param>
        /// <param name="idSI"></param>
        /// <returns></returns>
        [ComVisible(true)]
        public string GetBatchMetaDataSI(string docid, string batchno, string idSI)
        {
            if (batchno == string.Empty || idSI == string.Empty)
            {
                return string.Empty;
            }
            else
            {
                return wi.getBatchMetaDataByKey(wi.GetCurrentUserId(), docid, batchno, idSI);
            }
        }
        [ComVisible(true)]
        public string GetBatchMetaDataByTable(string docid, string batchno, string key, string table)
        {
            if (batchno == string.Empty || key == string.Empty)
            {
                return string.Empty;
            }
            else
            {
                return wi.getBatchMetaDataByKeyByTableForCurUser(docid, batchno, key, table);
            }
        }
        [ComVisible(true)]
        public string GetBatchMetaDataForPlan(string docid, string batchno, string key)
        {
            return wi.getBatchMetaDataByKeyForCurUserForPlan(batchno, key);
        }
        [ComVisible(true)]
        public string GetBatchMetaDataForCountPlan(string docid, string batchno, string key)
        {
            return wi.getBatchMetaDataByKeyByTable(wi.GetCurrentUserId(), docid, batchno, key, "IPM_COUNTPLANS_FS_LIKE_VIEW");
        }

        [ComVisible(true)]
        public string GetBatcOptPtoMultiMethod(string docid, string batchno, string key, string methodname, string testname, string opr, string kuv)
        {
            string result = GetBatchMDtoMultiMethod(docid, batchno, key, methodname, testname, opr);
            if (result == "")
            {
                kuv = kuv.Replace(" ", "");
                key = "ОптПлотнК" + kuv;
                result = GetBatchMDtoMultiMethod(docid, batchno, key, methodname, testname, opr);
                if (result == "")
                {
                    key = key.Replace("мм", "");
                    result = GetBatchMDtoMultiMethod(docid, batchno, key, methodname, testname, opr);
                }
            }
            return result;
        }
        /// <summary>
        /// Возвращает все показатели, привязанные к методу с типом "результат"
        /// </summary>
        /// <param name="methodsids">ячейка с ид методов</param>
        /// <returns>показатели</returns>
        [ComVisible(true)]
        public string GetMethodsTests(string methodsids)
        {
            return Utils.TypesUtils.ColToString(
                    MethodTests.GetByMethodIdsAndTestType(methodsids, "Результат")
                        .ConvertAll(test => test.ToFormatString("{Testdescr}")).ToArray(), ",");
        }
        [ComVisible(true)]
        public string GetMethodsTestsShortDesc(string methodsids)
        {
            return Utils.TypesUtils.ColToString(
                    MethodTests.GetByMethodIdsAndTestType(methodsids, "Результат")
                        .ConvertAll(test => test.ToFormatString("{Testshortdescr}")).ToArray(), ",");
        }
        [ComVisible(true)]
        public object GetBatcMasBSMultiMethod(string docid, string batchno, string key, string methodname, string testname, string opr, string n)
        {
            string[] arM = new string[7];
            int ind = -1;
            for (int i = 7; i > 0; i--)
            {
                arM[i - 1] = GetBatchMDtoMultiMethod(docid, batchno, key + i.ToString(), methodname, testname, opr);
                if (arM[i - 1] != "" && ind == -1) ind = i;
            }
            if (ind != -1) return WatersInterop.tryToDigit(arM[ind - int.Parse(n)], "");
            else return "";
        }
        [ComVisible(true)]
        public string GetBatchMDtoMultiMethod(string docid, string batchno, string key, string methodname, string testname, string opr)
        {
            string result = "";
            bool getSr = true;
            if (key.Contains("<or>"))
            {
                string[] keys = key.Split(new string[] { "<or>" }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string k in keys)
                {
                    string res = GetBatchMDtoMultiMethod(docid, batchno, k, methodname, testname, opr);
                    if (res != string.Empty)
                    {
                        return res;
                    }
                }
                return "";
            }
            else
            {
                if (opr.Contains("#"))
                {
                    opr = opr.Replace("#", "");
                    getSr = false;
                }
                if (key != "" && batchno != "")
                {
                    if (key == testname) key = "";
                    if (opr != "")
                    {
                        result = wi.getBatchMetaDataByKey(wi.GetCurrentUserId(), docid, batchno, methodname + "_" + key + opr);
                        if (result == "") result = wi.getBatchMetaDataByKey(wi.GetCurrentUserId(), docid, batchno, methodname + "_" + testname + key + opr);
                    }

                    if (result == "" && getSr)
                    {
                        result = wi.getBatchMetaDataByKey(wi.GetCurrentUserId(), docid, batchno, methodname + "_" + key);
                        if (result == "") result = wi.getBatchMetaDataByKey(wi.GetCurrentUserId(), docid, batchno, methodname + "_" + testname + key);
                    }
                    if (key.ToLower().Contains("формула"))
                    {
                        string rest = result;
                        try
                        {
                            if (!string.IsNullOrEmpty(rest) &&
                                rest.Split('=').Length > 1
                                && rest.Split('=')[0].Replace(" ", "") == rest.Split('=')[1].Replace(" ", ""))
                            {
                                rest = result.Split('=')[0].Replace(" ", "");
                            }
                        }
                        catch
                        {
                            rest = result;
                        }
                        result = rest;
                    }
                    if (key.ToLower().Contains("предповт") && result == "")
                    {
                        result = GetBatchMDtoMultiMethod(docid, batchno, "ПределПовторяемости", methodname, testname, opr);
                    }
                    string[] stringSep = new string[] { "<;>" };
                    if (result.Contains(stringSep[0]))
                    {
                        string[] arrRes = result.Split(stringSep, System.StringSplitOptions.RemoveEmptyEntries);
                        if (arrRes == null)
                        {
                            result = "";
                        }
                        else
                        {
                            result = arrRes[0];
                        }
                    }
                    if (result == "—") return "";
                    double vs;
                    if (double.TryParse(result, out vs))
                    {
                        //result = vs.ToString("0.0000000000000000000000").TrimEnd('0').TrimEnd(',');
                        result = vs.ToString("N15").TrimEnd('0').TrimEnd(',');
                    }
                    return result;
                }
                else return "";
            }
        }

        
        [ComVisible(true)]
        public string GetDocBatchMetaData(string docbatchid, string key)
        {
            return wi.getBatchMetaDataByKey(wi.GetCurrentUserId(), WatersInterop.GetDocFromDocBatchId(docbatchid), WatersInterop.GetBatchFromDocBatchId(docbatchid), key);
        }

        [ComVisible(true)]
        public DateTime TryToDate(string s, DateTime defval)
        {
            return WatersInterop.tryToDate(s, defval);
        }
        [ComVisible(true)]
        public string TryToDateStr(string s, string format)
        {
            return TypesUtils.tryToDateStr(s, format);
        }


        [ComVisible(true)]
        public double TryToNumber(string s, double ifnotnumber)
        {
            return WatersInterop.tryToNumber(s, ifnotnumber);
        }
        [ComVisible(true)]
        public object TryToDigit(string s, string ifnotnumber)
        {
            return WatersInterop.tryToDigit(s, ifnotnumber);
        }
        [ComVisible(true)]
        public string GetSomeDocid(Microsoft.Office.Interop.Excel.Range r, string lab, string docdescription, string key1, string Value1, string sortmetadataname, string DateTo)
        {
            return ei.GetSomeDocidBeforeDate(wi.GetCurrentUserId(), r, lab, docdescription, key1, Value1, sortmetadataname, DateTo);
        }

        [ComVisible(true)]
        public string GetSomeDocidBy3Keys(Microsoft.Office.Interop.Excel.Range r, string lab, string docdescription, string key1, string Value1, string key2, string Value2, string key3, string Value3,
        string sortmetadataname, string DateTo)
        {
            return ei.GetSomeDocidBeforeDate(wi.GetCurrentUserId(), r, lab, docdescription, key1, Value1, key2, Value2, key3, Value3, sortmetadataname, DateTo);
        }

        [ComVisible(true)]
        public object getMethodpogrforlab(Microsoft.Office.Interop.Excel.Range r, string methodname, string testid, string lab, double c)
        {
            return getMethodpogrforlabbyproduct(r, "", methodname, testid, lab, c);
        }
        [ComVisible(true)]
        public object getMethodpogrforlabbyproduct(Microsoft.Office.Interop.Excel.Range r, string product, string methodname, string testid, string lab, double c)
        {
            return VLKModule.getMethodpogr(r, product, methodname, testid, lab, "", c);
        }
        [ComVisible(true)]
        public object getMethodpogrUpLabStr(Microsoft.Office.Interop.Excel.Range r, string methodname, string testid, string lab, double c)
        {
            return getMethodpogrUpLabStrByPRoduct(r, "", methodname, testid, lab, c);
        }
        [ComVisible(true)]
        public object getMethodpogrUpLabStrByPRoduct(Microsoft.Office.Interop.Excel.Range r, string product, string methodname, string testid, string lab, double c)
        {
            return VLKModule.GetMethodpogrUpLabStr(r, product, methodname, testid, lab, "", c);
        }
        [ComVisible(true)]
        public object getMethodpogrUpLab(Microsoft.Office.Interop.Excel.Range r, string methodname, string testid, string lab, double c)
        {
            return getMethodpogrUpLabbyproduct(r, "", methodname, testid, lab, c);
        }
        [ComVisible(true)]
        public object getMethodpogrUpLabbyproduct(Microsoft.Office.Interop.Excel.Range r, string product, string methodname, string testid, string lab, double c)
        {
            return VLKModule.getMethodpogrUpLab(r, product, methodname, testid, lab, "", c);
        }



        [ComVisible(true)]
        public object getMethodpogr(Microsoft.Office.Interop.Excel.Range r, string methodname, string testid, double c)
        {
            return getMethodpogrbyproduct(r, "", methodname, testid, c);
        }
        [ComVisible(true)]
        public object getMethodpogrbyproduct(Microsoft.Office.Interop.Excel.Range r, string product, string methodname, string testid, double c)
        {
            return VLKModule.getMethodpogr(r, product, methodname, testid, "", "", c);
        }
        [ComVisible(true)]
        public object getMethodpogrWithDescription(Microsoft.Office.Interop.Excel.Range r, string methodname, string testid, string lab, string description, double c)
        {
            return VLKModule.getMethodpogr(r, "", methodname, testid, lab, description, c);
        }
        [ComVisible(true)]
        public object GetMethodSistprecUp(Microsoft.Office.Interop.Excel.Range r, string methodname, string testid, string lab, double c)
        {
            return GetMethodSistprecUpbyproduct(r, "", methodname, testid, lab, c);
        }
        [ComVisible(true)]
        public object GetMethodSistprecUpbyproduct(Microsoft.Office.Interop.Excel.Range r, string product, string methodname, string testid, string lab, double c)
        {
            return VLKModule.GetMethodSistprecUp(r, product, methodname, testid, lab, "", c);
        }
        [ComVisible(true)]
        public object GetMethodSistprecDown(Microsoft.Office.Interop.Excel.Range r, string methodname, string testid, string lab, double c)
        {
            return GetMethodSistprecDownbyproduct(r, "", methodname, testid, lab, c);
        }
        [ComVisible(true)]
        public object GetMethodSistprecDownbyproduct(Microsoft.Office.Interop.Excel.Range r, string product, string methodname, string testid, string lab, double c)
        {
            return VLKModule.GetMethodSistprecDown(r, product, methodname, testid, lab, "", c);
        }
        [ComVisible(true)]
        public object getMethodpogrKoeff(Microsoft.Office.Interop.Excel.Range r, string methodname, string testid, double c, double k)
        {
            return getMethodpogrKoeffForLab(r, methodname, testid, "", c, k);
        }
        [ComVisible(true)]
        public object getMethodOpredCount(Microsoft.Office.Interop.Excel.Range r, string methodname, string testid, string lab, double c)
        {
            return VLKModule.getMethodOpredCount(r, methodname, testid, lab, "", c);
        }
        [ComVisible(true)]
        public object getMethodpogrKoeffForLab(Microsoft.Office.Interop.Excel.Range r, string methodname, string testid, string lab, double c, double k)
        {
            return getMethodpogrKoeffForLabbyproduct(r, "", methodname, testid, lab, "", c, k);
        }
        [ComVisible(true)]
        public object getMethodpogrKoeffForLabbyproduct(Microsoft.Office.Interop.Excel.Range r, string product, string methodname, string testid, string lab, string description, double c, double k)
        {
            object result = VLKModule.getMethodpogr(r, product, methodname, testid, lab, description, c);
            if (result.ToString().Contains("<") || result.ToString().Contains(">"))
            {
                result = VLKModule.getMethodpogr(r, product, methodname, testid, lab, description, c / k);
                if (Math.Abs(WatersInterop.tryToNumber(result.ToString(), -999999) - -999999) >= 0.00001)
                {
                    result = WatersInterop.tryToNumber(result.ToString(), -999999) * k;
                }
            }
            return result;
        }

        [ComVisible(true)]
        public object getMethodpovt(Microsoft.Office.Interop.Excel.Range r, string methodname, string testid, string lab, double c)
        {
            return getMethodpovtbyproduct(r, "", methodname, testid, lab, "", c);
        }
        [ComVisible(true)]
        public object getMethodpovtbyproduct(Microsoft.Office.Interop.Excel.Range r, string product, string methodname, string testid, string lab, string description, double c)
        {
            return VLKModule.getMethodpovt(r, product, methodname, testid, lab, description, c);
        }
        [ComVisible(true)]
        public object getMethodpovtWithDescription(Microsoft.Office.Interop.Excel.Range r, string methodname, string testid, string lab, string description, double c)
        {
            return VLKModule.getMethodpovt(r, "", methodname, testid, lab, description, c);
        }

        [ComVisible(true)]
        public object getMethodpovtManyArg(Microsoft.Office.Interop.Excel.Range r, string methodname, string testid, string lab, string c)
        {
            return getMethodpovtManyArgbyproduct(r, "", methodname, testid, lab, c);
        }
        [ComVisible(true)]
        public object getMethodpovtManyArgbyproduct(Microsoft.Office.Interop.Excel.Range r, string product, string methodname, string testid, string lab, string c)
        {
            double retV = 0.0;
            string retVal = VLKModule.getMethodpovtManyArg(r, product, methodname, testid, lab, "", c);
            if (double.TryParse(retVal, out retV))
                return Math.Round(retV, 4);
            else return retVal;
        }

        [ComVisible(true)]
        public object getMethodpovtForLab(Microsoft.Office.Interop.Excel.Range r, string methodname, string testid, string lab, double c)
        {
            return getMethodpovtForLabbyproduct(r, "", methodname, testid, lab, c);
        }
        [ComVisible(true)]
        public object getMethodpovtForLabbyproduct(Microsoft.Office.Interop.Excel.Range r, string product, string methodname, string testid, string lab, double c)
        {
            return VLKModule.getMethodpovt(r, product, methodname, testid, lab, "", c);
        }

        [ComVisible(true)]
        public object getMethodpovtKoeffForLab(Microsoft.Office.Interop.Excel.Range r, string methodname, string testid, string lab, double c, double k)
        {
            return getMethodpovtKoeffForLabbyproduct(r, "", methodname, testid, lab, "", c, k);
        }
        [ComVisible(true)]
        public object getMethodpovtKoeffForLabbyproduct(Microsoft.Office.Interop.Excel.Range r, string product, string methodname, string testid, string lab, string description, double c, double k)
        {
            object result;
            /*result = VLKModule.getMethodpovt(wi, r, methodname, testid, lab, c);
            if (result.ToString().Contains("<") || result.ToString().Contains(">"))
            {*/
            result = VLKModule.getMethodpovt(r, product, methodname, testid, lab, description, c / k);
            if (Math.Abs(WatersInterop.tryToNumber(result.ToString(), -999999) - -999999) >= 0.000001)
            {
                result = WatersInterop.tryToNumber(result.ToString(), -999999) /** k*/; //Закомментировал по письму Дмитриевой Т.В. от 7.11.2013
            }
            //}
            return result;
        }

        [ComVisible(true)]
        public object getMethodpovtKoeff(Microsoft.Office.Interop.Excel.Range r, string methodname, string testid, double c, double k)
        {

            return getMethodpovtKoeffForLab(r, methodname, testid, "", c, k);
        }
        [ComVisible(true)]
        public string GetMethodUnit(Microsoft.Office.Interop.Excel.Range r, string methodname, string testid, string lab, string c)
        {
            return GetMethodUnitForLab(r, methodname, testid, "%", c);
        }

        [ComVisible(true)]
        public string GetMethodUnitForLab(Microsoft.Office.Interop.Excel.Range r, string methodname, string testid, string lab, string c)
        {
            string result = VLKModule.GetMethodUnit(r, methodname, testid, lab, "", c);
            if (result == "Ошибка. Единицы измерения не заданы") return "";
            return result;
        }


        [ComVisible(true)]
        public string GetMethodDiap(Microsoft.Office.Interop.Excel.Range r, string methodname, string testid, double c)
        {
            return GetMethodDiapbyproduct(r, "", methodname, testid, c);
        }
        [ComVisible(true)]
        public string GetMethodDiapbyproduct(Microsoft.Office.Interop.Excel.Range r, string product, string methodname, string testid, double c)
        {
            return VLKModule.GetMethodDiap(r, product, methodname, testid, "", "", c);
        }
        [ComVisible(true)]
        public object GetMethodStabreq(Microsoft.Office.Interop.Excel.Range r, string product, string methodname, string testid, string lab, string description, double c)
        {
            return VLKModule.GetMethodStabreq(r, product, methodname, testid, lab, description, c);
        }

        [ComVisible(true)]
        public string GetMethodDiapForLabbyproduct(Microsoft.Office.Interop.Excel.Range r, string product, string methodname, string testid, string lab, double c)
        {
            return VLKModule.GetMethodDiap(r, product, methodname, testid, lab, "", c);
        }
        [ComVisible(true)]
        public string GetMethodDiapForLab(Microsoft.Office.Interop.Excel.Range r, string methodname, string testid, string lab, double c)
        {
            return GetMethodDiapForLabbyproduct(r, "", methodname, testid, lab, c);
        }
        [ComVisible(true)]
        public string GetMethodPokpovtStr(Microsoft.Office.Interop.Excel.Range r, string methodname, string testid, double c)
        {
            return GetMethodPokpovtStrbyproduct(r, "", methodname, testid, c);
        }
        [ComVisible(true)]
        public string GetMethodPokpovtStrbyproduct(Microsoft.Office.Interop.Excel.Range r, string product, string methodname, string testid, double c)
        {
            return VLKModule.GetMethodPokpovtStr(r, product, methodname, testid, "", "", c);
        }
        [ComVisible(true)]
        public string GetMethodPokLabprecStr(Microsoft.Office.Interop.Excel.Range r, string methodname, string testid, string lab, double c)
        {
            return GetMethodPokLabprecStrbyproduct(r, "", methodname, testid, lab, c);
        }
        [ComVisible(true)]
        public string GetMethodPokLabprecStrbyproduct(Microsoft.Office.Interop.Excel.Range r, string product, string methodname, string testid, string lab, double c)
        {
            return VLKModule.GetMethodPokLabprecStr(r, product, methodname, testid, lab, "", c);
        }
        
        [ComVisible(true)]
        public string GetMethodprecLabStr(Microsoft.Office.Interop.Excel.Range r, string product, string methodname, string testid, string lab, string description, double c)
        {
            return VLKModule.GetMethodprecLabStr(r, product, methodname, testid, lab, "", c);
        }
        
        [ComVisible(true)]
        public object GetMethodPokLabprec(Microsoft.Office.Interop.Excel.Range r, string product, string methodname, string testid, string lab, string description, double c)
        {
            return VLKModule.GetMethodPokLabprec(r, product, methodname, testid, lab, "", c);
        }

        [ComVisible(true)]
        public object GetMethodPokPovt(Microsoft.Office.Interop.Excel.Range r, string product, string methodname, string testid, string lab, string description, double c)
        {
            return VLKModule.GetMethodPokpovt(r, product, methodname, testid, lab, "", c);
        }

        [ComVisible(true)]
        public string GetMethodPokpovtStrForLab(Microsoft.Office.Interop.Excel.Range r, string methodname, string testid, string lab, double c)
        {
            return GetMethodPokpovtStrForLabbyproduct(r, "", methodname, testid, lab, c);
        }
        [ComVisible(true)]
        public string GetMethodPokpovtStrForLabbyproduct(Microsoft.Office.Interop.Excel.Range r, string product, string methodname, string testid, string lab, double c)
        {
            return VLKModule.GetMethodPokpovtStr(r, product, methodname, testid, lab, "", c);
        }
        [ComVisible(true)]
        public string GetMethodPovtStr(Microsoft.Office.Interop.Excel.Range r, string methodname, string testid, string lab, double c)
        {
            return GetMethodPovtStrbyproduct(r, "", methodname, testid, lab, c);
        }
        [ComVisible(true)]
        public string GetMethodPovtStrbyproduct(Microsoft.Office.Interop.Excel.Range r, string product, string methodname, string testid, string lab, double c)
        {
            return VLKModule.GetMethodpovtStr(r, product, methodname, testid, lab, "", c);
        }
        [ComVisible(true)]
        public string GetMethodpogrStr(Microsoft.Office.Interop.Excel.Range r, string methodname, string testid, double c)
        {
            return GetMethodpogrStrForLab(r, methodname, testid, "", c);
        }
        [ComVisible(true)]
        public string GetMethodpogrStrForLab(Microsoft.Office.Interop.Excel.Range r, string methodname, string testid, string lab, double c)
        {
            return GetMethodpogrStrForLabbyproduct(r, "", methodname, testid, lab, c);
        }
        [ComVisible(true)]
        public string GetMethodpogrStrForLabbyproduct(Microsoft.Office.Interop.Excel.Range r, string product, string methodname, string testid, string lab, double c)
        {
            return VLKModule.GetMethodpogrStr(r, product, methodname, testid, lab, "", c);
        }
        [ComVisible(true)]
        public string GetMetaNameFromMethodTestId(string methodid, string testid)
        {
            return CalcModule.GetMetaNameFromMethodTestId(methodid, testid);
        }
        public string GetTestResult(string docid, string batchno, string methodid, string testid, string numberprec)
        {
            DebugModule.WriteDebugMessage("GetTestResult: docid [{0}], batchno [{1}], methodid [{2}], testid [{3}], numberprec [{4}]", docid, batchno, methodid, testid, numberprec);
            return GetResultString(GetBatchMetaData(docid, batchno, GetMetaNameFromMethodTestId(methodid, testid)), "", numberprec);
        }
        public string GetMethodNumberPrec(Microsoft.Office.Interop.Excel.Range r, string methodname, string testid, double c)
        {
            return GetMethodNumberPrecForLab(r, methodname, testid, "", c);
        }

        public string GetMethodNumberPrecForLab(Microsoft.Office.Interop.Excel.Range r, string methodname, string testid, string lab, double c)
        {
            return GetMethodNumberPrecForLabbyproduct(r, "", methodname, testid, lab, c);
        }
        public string GetMethodNumberPrecForLabbyproduct(Microsoft.Office.Interop.Excel.Range r, string product, string methodname, string testid, string lab, double c)
        {
            return VLKModule.GetMethodNumberPrec(r, product, methodname, testid, lab, "", c);
        }
        [ComVisible(true)]
        public void UpdateMethodsDataOnWS(Microsoft.Office.Interop.Excel.Worksheet ws
                    , int startrow, string productid, string[] methodids)
        {
            VLKModule.UpdateMethodsDataOnWS(ei, startrow, productid, methodids);
        }
        [ComVisible(true)]
        public string GetFormula(Microsoft.Office.Interop.Excel.Range r)
        {
            return ei.GetFormula(r);
        }
        [ComVisible(true)]
        public string GetFormuls(Microsoft.Office.Interop.Excel.Range r)
        {
            return ei.GetFormula(r);
        }

        [ComVisible(true)]
        public void AllFinished()
        {
            DebugModule.MarkAll(ei, true);
        }
        [ComVisible(true)]
        public void AllUnFinished()
        {
            DebugModule.MarkAll(ei, false);
        }

        [ComVisible(true)]
        public bool CompareStringsByElements(string s1, string s2, string delimeter)
        {
            string[] vsa1 = s1.Split(new string[] { delimeter }, StringSplitOptions.None);
            string[] vsa2 = s2.Split(new string[] { delimeter }, StringSplitOptions.None);
            foreach (string vs1 in vsa1)
            {
                bool flag = false;
                foreach (string vs2 in vsa2)
                {
                    if (vs1 == vs2)
                    {
                        flag = true;
                        break;
                    }
                }
                if (!flag)
                {
                    return false;
                }
            }
            return true;

        }

        [ComVisible(true)]
        public string RoundTemp(double d)
        {
            int id = (int)d;
            double rd = Math.Abs(d - (double)id);
            if (rd >= 0.25)
            {
                if (rd < 0.75)
                {
                    if (id < 0)
                    {
                        d = id - 0.5;
                    }
                    else
                    {
                        d = id + 0.5;
                    }
                    return d.ToString();
                }
                else
                {
                    if (id < 0)
                    {
                        d = id - 1;
                    }
                    else
                    {
                        d = id + 1;
                    }
                    return d.ToString();
                }
            }
            else
            {
                return id.ToString();
            }
        }

        public bool IsInInterval(string val, string interval, bool ignorelower)
        {
            double value;
            double l;
            double u;
            string[] vsa;
            value = TryToNumber(val, -99999999);
            if (Math.Abs(value - -99999999) <= 0.00000000001)
            {
                return false;
            }
            if (interval.Contains(" до "))
            {
                vsa = interval.Split(new string[] { " до " }, StringSplitOptions.RemoveEmptyEntries);
                //если не чисолвые значения, то вернуть false
                l = TryToNumber(vsa[0], 99999999);
                u = TryToNumber(vsa[1], -99999999);

                return (l <= value || ignorelower) && value <= u;
            }
            return false;
        }
        [ComVisible(true)]
        public string GetCalibrGrafDocid(Microsoft.Office.Interop.Excel.Range r, string lab, string kuveta1, string optplotn1
                , string kuveta2, string optplotn2, string prodnumber, string samplingplace, string test)
        {
            return KalibrGrafForm.GetCalibrGrafDocid(ei, r, lab, kuveta1, optplotn1, kuveta2, optplotn2, prodnumber, samplingplace, test);
        }
        [ComVisible(true)]
        public string GetWorkNumber(Microsoft.Office.Interop.Excel.Range r, string what, int n /*= 1*/, int searchcolindex/*= 1*/, int colindex/*= 2*/)
        {
            return GetSomethingInRange(r, what, true, n, searchcolindex, colindex);
        }
        [ComVisible(true)]
        public string GetSomethingInRange(Microsoft.Office.Interop.Excel.Range r, string what, bool isInstr, int n, int searchcolindex, int colindex)
        {
            try
            {
                int j = 0;
                object[,] values = (object[,])r.Value2;
                for (int i = 1; i <= r.Rows.Count; i++)
                {
                    if (ExcelInteractClass.GetObjectStringValue(values[i, searchcolindex]).ToLower() == what.ToLower()
                        || (ExcelInteractClass.GetObjectStringValue(values[i, searchcolindex]).ToLower().Contains(what.ToLower()) && isInstr))
                    {
                        j++;
                        if (j == n)
                        {
                            return ExcelInteractClass.GetObjectStringValue(values[i, colindex]);
                        }
                    }
                    if (ExcelInteractClass.GetCell(r, i, searchcolindex).Row > r.Worksheet.UsedRange.Rows.Count + r.Worksheet.UsedRange.Row)
                    {
                        return "";
                    }
                }
                return "";
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }
        [ComVisible(true)]
        public object ReturnChar(Microsoft.Office.Interop.Excel.Range r, string charstr, double c)
        {
            return VLKModule.ReturnChar(r, charstr, c);
        }

        [ComVisible(true)]
        public string GetUslByIntervals(Microsoft.Office.Interop.Excel.Range r)
        {
            return GetUslByIntervals((object[,])r.Value2);
        }
        public string GetUslByIntervals(object[,] data)
        {
            return ReportClass.GetStringIntervals(data, new string[,] { { "Направление ветра ", "" }, { "P= ", " мм рт. ст." }, { "вл= ", "%" }, { "t= ", "°C" } }, "; ", 0);
        }


        #region Блок переноса функций из "С прим. СО ККШ"
        [ComVisible(true)]
        public double getAn(int i, int n) { return VLKModule.getAn(i, n); }
        [ComVisible(true)]
        public int getAllAnCount(double gamma, int pn, double val, string function) { return VLKModule.getAllAnCount(gamma, pn, val, function); }
        [ComVisible(true)]
        public double getprecancountf(double gamma, int pn, double val) { return VLKModule.getprecancountf(gamma, pn, val); }
        [ComVisible(true)]
        public double getpovtancountf(double gamma, int pn, double val) { return VLKModule.getpovtancountf(gamma, pn, val); }
        [ComVisible(true)]
        public double getpograncountf(double gamma, int pn, double val) { return VLKModule.getpograncountf(gamma, pn, val); }
        [ComVisible(true)]
        public int CalcCount(Microsoft.Office.Interop.Excel.Range datar, Microsoft.Office.Interop.Excel.Range vyv)
        {
            return VLKModule.calcCount(datar, vyv);
        }
        [ComVisible(true)]
        public double CalcSum(Microsoft.Office.Interop.Excel.Range datar, Microsoft.Office.Interop.Excel.Range vyv)
        {
            return VLKModule.CalcSum(datar, vyv);
        }
        [ComVisible(true)]
        public double CalcSigmaRlb(Microsoft.Office.Interop.Excel.Range datar, Microsoft.Office.Interop.Excel.Range vyv)
        {
            return VLKModule.CalcSigmaRlb(datar, vyv);
        }
        [ComVisible(true)]
        public double CalcSigmaRlOneVal(Microsoft.Office.Interop.Excel.Range datar, Microsoft.Office.Interop.Excel.Range vyv)
        {
            return VLKModule.CalcSigmaRlbOneVal(datar, vyv);
        }
        [ComVisible(true)]
        public double CalcSigmaSl(Microsoft.Office.Interop.Excel.Range datar, Microsoft.Office.Interop.Excel.Range vyv, double tetta)
        {
            return VLKModule.CalcSigmaSl(datar, vyv, tetta);
        }
        [ComVisible(true)]
        public string CalcIsk(Microsoft.Office.Interop.Excel.Range datar, Microsoft.Office.Interop.Excel.Range vyv, Microsoft.Office.Interop.Excel.Range nprange)
        {
            return VLKModule.CalcIsk(datar, vyv, nprange);
        }
        [ComVisible(true)]
        public string getVyv(string povt, string pogr, string prec)
        {
            return VLKModule.getVyv(povt, pogr, prec);
        }
        [ComVisible(true)]
        public string GetDorPrPrev(double d, double pr, double sr, int rowindex, Microsoft.Office.Interop.Excel.Range cell, Microsoft.Office.Interop.Excel.Range povtcell)
        {
            return VLKModule.GetDorPrPrev(d, pr, sr, rowindex, cell, povtcell);
        }
        [ComVisible(true)]
        public string GetDorPrPogrPrev(double du, double dl, double pru, double prl, double sr, int rowindex, Microsoft.Office.Interop.Excel.Range Range, Microsoft.Office.Interop.Excel.Range povtcell)
        {
            return VLKModule.GetDorPrPogrPrev(du, dl, pru, prl, sr, rowindex, Range, povtcell);
        }


        [ComVisible(true)]
        public string GetPovtPrecConclusion(double d, double pr, double sr, int rowindex, Microsoft.Office.Interop.Excel.Range cell)
        {
            return VLKModule.GetPovtPrecConclusion(d, pr, sr, rowindex, cell);
        }
        [ComVisible(true)]
        public string GetPogrConclusion(double du, double dl, double pru, double prl, double sr, int rowindex, Microsoft.Office.Interop.Excel.Range cell)
        {
            return VLKModule.GetPogrConclusion(du, dl, pru, prl, sr, rowindex, cell);
        }
        [ComVisible(true)]

        public double getFunc(Microsoft.Office.Interop.Excel.Worksheet ws, double f)
        {
            int i, rc, ri, ri1;
            double udelta;
            double ldelta;
            double uval;
            double lval;


            rc = ws.UsedRange.Rows.Count;

            ri = -1;
            ri1 = -1;
            udelta = 99999;
            ldelta = 99999;
            for (i = 3; i <= rc; i++)
            {
                if ((double)ExcelInteractClass.GetCellValue(ws, i, 1) >= f)
                {
                    if (udelta > (double)ExcelInteractClass.GetCellValue(ws, i, 1) - f)
                    {
                        udelta = (double)ExcelInteractClass.GetCellValue(ws, i, 1) - f;
                        ri = i;
                    }
                }
                if ((double)ExcelInteractClass.GetCellValue(ws, i, 1) < f)
                {
                    if (ldelta > f - (double)ExcelInteractClass.GetCellValue(ws, i, 1))
                    {
                        ldelta = f - (double)ExcelInteractClass.GetCellValue(ws, i, 1);
                        ri1 = i;
                    }
                }
            }

            if (ri != -1 && ri1 != -1)
            {
                udelta = (double)ExcelInteractClass.GetCellValue(ws, ri, 1);
                ldelta = (double)ExcelInteractClass.GetCellValue(ws, ri1, 1);
                uval = (double)ExcelInteractClass.GetCellValue(ws, ri, 2);
                lval = (double)ExcelInteractClass.GetCellValue(ws, ri1, 2);
                return lval + (uval - lval) * (f - ldelta) / (udelta - ldelta);
            }
            throw (new Exception("Не найдено диаппазона в который попало бы f"));

        }






        #endregion

        private double getKoh(int f)
        {
            Microsoft.Office.Interop.Excel.Worksheet ws = ei.GetWorksheetFromActiveWorkbook("Кохрен", false);

            for (int i = 3; i <= ws.UsedRange.Rows.Count; i++)
            {
                Microsoft.Office.Interop.Excel.Range r = ws.Range[ws.Cells[i, 1], ws.Cells[i, 1]];
                if (Math.Abs(double.Parse(ExcelInteractClass.GetObjectStringValue(ws.Range[ws.Cells[i, 1], ws.Cells[i, 1]].Value2)) - f) <= 0.0000000001
                    && Math.Abs(double.Parse(ExcelInteractClass.GetObjectStringValue(ws.Range[ws.Cells[i, 2], ws.Cells[i, 2]].Value2)) - 1) <= 0.0000000001)
                    return double.Parse(ExcelInteractClass.GetObjectStringValue(ws.Range[ws.Cells[i, 3], ws.Cells[i, 3]].Value2));
            }
            return 0;
        }

        private double getGrabs(int l)
        {
            Microsoft.Office.Interop.Excel.Worksheet _ws = ei.GetWorksheetFromActiveWorkbook("Грабс", false);
            for (int i = 2; i <= _ws.UsedRange.Rows.Count; i++)
                if (Math.Abs(double.Parse(ExcelInteractClass.GetObjectStringValue(_ws.Range[_ws.Cells[i, 1], _ws.Cells[i, 1]].Value2)) - l) <= 0.0000000001)
                    return double.Parse(ExcelInteractClass.GetObjectStringValue(_ws.Range[_ws.Cells[i, 2], _ws.Cells[i, 2]].Value2));
            return 99999;
        }

        [ComVisible(true)]
        public void CalcCharsForKontPeriod(double c, double delta)
        {
            Microsoft.Office.Interop.Excel.Worksheet _ws;
            Microsoft.Office.Interop.Excel.Worksheet _logws;
            Microsoft.Office.Interop.Excel.Worksheet _protws;
            Microsoft.Office.Interop.Excel.Worksheet _dataws;

            int i = 0, l = 0, sr = 0, logr = 0;
            double t = 0.0, ttabl = 0.0, Gmax = 0.0, Gtabl = 0.0, sigmarl = 0.0,
                /*sigmbrl = 0.0,*/ obavg = 0.0, tetta = 0.0, deltalstu = 0.0, deltalstl = 0.0, deltalu = 0.0, deltall = 0.0,
                sigmabrl = 0.0;
            bool isk = false;

            _ws = ei.GetWorksheetFromActiveWorkbook("Данные для расчёта", false);
            _logws = ei.GetWorksheetFromActiveWorkbook("Журнал расчёта", false);
            _protws = ei.GetWorksheetFromActiveWorkbook("Протокол", false);
            _dataws = ei.GetWorksheetFromActiveWorkbook("Данные", false);

            l = _ws.UsedRange.Rows.Count;
            logr = 1;

            Microsoft.Office.Interop.Excel.Range r_ws = _ws.Range[_ws.Cells[1, 1], _ws.Cells[l, 5]];
            r_ws.Copy(Type.Missing);
            Microsoft.Office.Interop.Excel.Range log_ws = _logws.Range[_logws.Cells[5, 1], _logws.Cells[5, 1]];
            _logws.Paste(Type.Missing);

            logr = 6 + l;

            /*_ws.Range[_ws.Cells(1, 1), _ws.Cells(l, 6)).Sort key1:=Range("E1"), Order1:=xlAscending, Header:= _
                xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
                DataOption1:=xlSortNormal*/
            Gmax = 1;
            while (Gmax > Gtabl)
            {
                Gmax = 0;
                for (i = 1; i <= l; i++)
                    Gmax = Gmax + double.Parse(ExcelInteractClass.GetObjectStringValue(_ws.Range[_ws.Cells[i, 5], _ws.Cells[i, 5]].Value2));
                Gmax = double.Parse(ExcelInteractClass.GetObjectStringValue(_ws.Range[_ws.Cells[l, 5], _ws.Cells[l, 5]].Value2)) / Gmax;
                Gtabl = getKoh(l);
                l--;
                if (Gmax > Gtabl)
                {
                    _logws.Range[_logws.Cells[logr, 1], _logws.Cells[logr, 1]].Value2 = "Расчетное значение кр.Кохрена (" + Math.Round(Gmax, 3) + ") больше табличного(" + Math.Round(Gtabl, 3) + ")";
                    logr++;
                    _logws.Range[_logws.Cells[logr, 1], _logws.Cells[logr, 1]].Value2 = "Дисперсии не однородны";
                    logr++;
                    _logws.Range[_logws.Cells[logr, 1], _logws.Cells[logr, 1]].Value2 = "Исключаем " + _logws.Range[_logws.Cells[l + 1, 1], _logws.Cells[l + 1, 1]].Value2 + " контрольную процедуру";
                    logr++;
                    _logws.Range[_logws.Cells[logr, 1], _logws.Cells[logr, 1]].Value2 = "Осталось " + l + " значений";
                    logr++;
                }
            }
            l = l + 1;
            Gmax = 0;
            for (i = 1; i <= l; i++)
                if (_ws.Range[_ws.Cells[i, 5], _ws.Cells[i, 5]].Value2 != null)
                    Gmax = Gmax + double.Parse(ExcelInteractClass.GetObjectStringValue(_ws.Range[_ws.Cells[i, 5], _ws.Cells[i, 5]].Value2));
            sigmarl = Math.Sqrt(Gmax / l);
            _logws.Range[_logws.Cells[logr, 1], _logws.Cells[logr, 1]].Value2 = "Дисперсии однородны";
            logr++;
            _logws.Range[_logws.Cells[logr, 1], _logws.Cells[logr, 1]].Value2 = "Показатель повторяемости равен:" + Math.Round(sigmarl, 3);
            _dataws.Range[_dataws.Cells[21, 2], _dataws.Cells[21, 2]].Value2 = Math.Round(sigmarl, 3);
            logr++;
            isk = true;
            sr = 1;
            while (isk)
            {
                /*((Microsoft.Office.Interop.Excel.Range)_ws.Columns.get_Item(sr, 1)).Sort(_ws.Columns[l, 6],
                Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending, Type.Missing, Type.Missing,
                Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending, Type.Missing,
                Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending, Microsoft.Office.Interop.Excel.XlYesNoGuess.xlGuess,
                Type.Missing, Type.Missing,
                Microsoft.Office.Interop.Excel.XlSortOrientation.xlSortColumns, Microsoft.Office.Interop.Excel.XlSortMethod.xlPinYin,
                Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal, Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal);
                */
                Gtabl = getGrabs(l);
                _logws.Range[_logws.Cells[logr, 1], _logws.Cells[logr, 1]].Value2 = "Табличное значение критерия Грабса: " + Math.Round(Gtabl, 3);
                logr++;
                Gmax = 0;
                for (i = sr; i <= l; i++)
                    if (_ws.Range[_ws.Cells[i, 4], _ws.Cells[i, 4]].Value2 != null)
                        Gmax = Gmax + double.Parse(ExcelInteractClass.GetObjectStringValue(_ws.Range[_ws.Cells[i, 4], _ws.Cells[i, 4]].Value2));
                obavg = Gmax / (l - sr + 1);
                Gmax = 0;
                for (i = sr; i <= l; i++)
                    if (_ws.Range[_ws.Cells[i, 4], _ws.Cells[i, 4]].Value2 != null)
                        Gmax = Gmax + (double.Parse(_ws.Range[_ws.Cells[i, 4], _ws.Cells[i, 4]].Value2.ToString()) - obavg) *
                            (double.Parse(_ws.Range[_ws.Cells[i, 4], _ws.Cells[i, 4]].Value2.ToString()) - obavg);
                sigmabrl = Math.Sqrt(Gmax / (l - sr));
                isk = false;
                if (_ws.Range[_ws.Cells[sr, 4], _ws.Cells[sr, 4]].Value2 != null)
                    Gmax = (obavg - double.Parse(_ws.Range[_ws.Cells[sr, 4], _ws.Cells[sr, 4]].Value2.ToString())) / sigmabrl;

                if (Gmax > Gtabl)
                {
                    _logws.Range[_logws.Cells[logr, 4], _logws.Cells[logr, 4]].Value2 = "Статистика Грабса для мин. (" + Math.Round(Gmax, 3) + ") больше критической(" + Math.Round(Gtabl, 3) + ")";
                    logr++;
                    _logws.Range[_logws.Cells[logr, 4], _logws.Cells[logr, 4]].Value2 = "Исключаем " + _ws.Range[_ws.Cells[sr, 1], _ws.Cells[sr, 1]].Value2 + " контрольную процедуру";
                    logr++;
                    _logws.Range[_logws.Cells[logr, 4], _logws.Cells[logr, 4]].Value2 = "Осталось " + (l - sr) + " значений";
                    logr++;
                    sr++;
                    isk = true;
                }
                else
                {
                    _logws.Range[_logws.Cells[logr, 4], _logws.Cells[logr, 4]].Value2 = "Статистика Грабса для мин. значения (" + Math.Round(Gmax, 3) + ") меньше критической(" + Math.Round(Gtabl, 3) + ")";
                    logr++;
                    isk = false;
                }
                Gmax = (double.Parse(ExcelInteractClass.GetObjectStringValue(_ws.Range[_ws.Cells[l, 4], _ws.Cells[l, 4]].Value2)) - obavg) / sigmabrl;

                if (Gmax > Gtabl)
                {
                    _logws.Range[_logws.Cells[logr, 1], _logws.Cells[logr, 1]].Value2 = "Статистика Грабса для макс. (" + Math.Round(Gmax, 3) + ") больше критической(" + Math.Round(Gtabl, 3) + ")";
                    logr++;
                    _logws.Range[_logws.Cells[logr, 1], _logws.Cells[logr, 1]].Value2 = "Исключаем "
                        + ExcelInteractClass.GetObjectStringValue(_ws.Range[_ws.Cells[l, 1], _ws.Cells[l, 1]].Value2) + " контрольную процедуру";
                    logr++;
                    l--;
                    _logws.Range[_logws.Cells[logr, 1], _logws.Cells[logr, 1]].Value2 = "Осталось " + (l - sr) + " значений";
                    logr++;
                    isk = true;
                }
                else
                {
                    _logws.Range[_logws.Cells[logr, 1], _logws.Cells[logr, 1]].Value2 = "Статистика Грабса для макс. значения (" + Math.Round(Gmax, 3) + ") меньше критической(" + Math.Round(Gtabl, 3) + ")";
                    logr++;
                    isk = false;
                }
            }
            l = l - sr + 1;
            _logws.Range[_logws.Cells[logr, 1], _logws.Cells[logr, 1]].Value2 = "Осталось:" + l + " контрольных проб";
            logr++;
            _logws.Range[_logws.Cells[logr, 1], _logws.Cells[logr, 1]].Value2 = "Показатель внутрилабораторной прецизионности равен:" + Math.Round(sigmabrl, 3);
            _dataws.Range[_dataws.Cells[24, 2], _dataws.Cells[24, 2]].Value2 = "'" + Math.Round(sigmabrl, 3);
            logr++;
            _protws.Range[_protws.Cells[18, 2], _protws.Cells[18, 2]].Value2 = Math.Round(sigmabrl, 3);
            tetta = obavg - c;
            t = Math.Abs(tetta) / Math.Sqrt(Math.Pow(sigmabrl, 2) / l + Math.Pow(delta, 2) / 3);
            ttabl = CalcModule.GetStud(l - 1);

            if (t < ttabl)
            {
                _logws.Range[_logws.Cells[logr, 1], _logws.Cells[logr, 1]].Value2 = "t(" + Math.Round(t, 3) + ") меньше табличного значения (" + Math.Round(ttabl, 3) + "). Оценка систематической погрешности незначима";
                logr++;
                tetta = 0;
            }
            else
            {
                _logws.Range[_logws.Cells[logr, 1], _logws.Cells[logr, 1]].Value2 = "t(" + Math.Round(t, 3) + ") больше табличного значения (" + Math.Round(ttabl, 3) + "). Оценка систематической погрешности равна:" + Math.Round(tetta, 3);
                logr++;
            }
            deltalstu = tetta + 2 * Math.Sqrt(Math.Pow(sigmabrl, 2) / l + Math.Pow(delta, 2) / 3);
            deltalstl = tetta - 2 * Math.Sqrt(Math.Pow(sigmabrl, 2) / l + Math.Pow(delta, 2) / 3);
            if (Math.Abs(deltalstu) - Math.Abs(deltalstl) < 0.0000000001)
            {
                _protws.Range[_protws.Cells[18, 3], _protws.Cells[18, 3]].Value2 = "±" + Math.Round(deltalstu, 3);
            }
            else
            {
                _protws.Range[_protws.Cells[18, 3], _protws.Cells[18, 3]].Value2 = "(" + Math.Round(deltalstl, 3) + ";" + Math.Round(deltalstu, 3) + ")";
            }
            _dataws.Range[_dataws.Cells[25, 2], _dataws.Cells[25, 2]].Value2 = "'" + Math.Round(deltalstu, 3);
            _dataws.Range[_dataws.Cells[26, 2], _dataws.Cells[26, 2]].Value2 = "'" + Math.Round(deltalstl, 3);
            _logws.Range[_logws.Cells[logr, 1], _logws.Cells[logr, 1]].Value2 = "Верхняя " + Math.Round(deltalstu, 3) + " и нижняя " + Math.Round(deltalstl, 3) + " границы показателя правильности";
            logr++;

            if (deltalstu / sigmabrl < 0.8)
            {
                _logws.Range[_logws.Cells[logr, 1], _logws.Cells[logr, 1]].Value2 = "Выполняется условие дельта / показатель прецизионности < 0.8 (см. В.3.2.6 примечание 1 РМГ 76). Показатель точности берем как 2*показатель прецизионности";
                logr++;
                deltalu = deltall = 2 * sigmabrl;
                //deltall = deltau;
            }
            else
            {
                deltalu = tetta + 2 * Math.Sqrt(Math.Pow(sigmabrl, 2) + Math.Pow((deltalstu / 2), 2));
                deltall = tetta - 2 * Math.Sqrt(Math.Pow(sigmabrl, 2) + Math.Pow((deltalstu / 2), 2));
            }
            if (Math.Abs(deltalu) - Math.Abs(deltall) < 0.0000000001)
            {
                _protws.Range[_protws.Cells[18, 4], _protws.Cells[18, 4]].Value2 = "±" + Math.Round(deltalu, 3);
            }
            else
            {
                _protws.Range[_protws.Cells[18, 4], _protws.Cells[18, 4]].Value2 = "(" + Math.Round(deltall, 3) + ";" + Math.Round(deltalu, 3) + ")";
            }
            _dataws.Range[_dataws.Cells[29, 2], _dataws.Cells[29, 2]].Value2 = "'" + Math.Round(deltalu, 3);
            _dataws.Range[_dataws.Cells[30, 2], _dataws.Cells[30, 2]].Value2 = "'" + Math.Round(deltall, 3);
            _logws.Range[_logws.Cells[logr, 1], _logws.Cells[logr, 1]].Value2 = "Верхняя " + Math.Round(deltalu, 3) + " и нижняя " + Math.Round(deltall, 3) + " границы показателя точности";
            logr++;
            _dataws.Range[_dataws.Cells[8, 2], _dataws.Cells[8, 2]].Value2 = SystemTime.Current;
        }

     
        public string GetSpecNorm(string batchno, string orderattribute)
        {
            return Specs.GetNormForSample(batchno, orderattribute);
        }
        public string GetSpecNormWithNaimspec(string batchno, string orderattribute, string naimspec)
        {
            return Specs.GetNormForSample(batchno, orderattribute, naimspec);
        }
        public string GetSpecRaschNorm(string batchno, string orderattribute)
        {
            return Specs.GetRaschNormForSample(batchno, orderattribute);
        }
        public string GetSpecNumPrec(string batchno, string orderattribute)
        {
            return Specs.GetNumPrecForSample(batchno, orderattribute);
        }
        public string GetSpecNo(string batchno, string orderattribute)
        {
            return Specs.GetSpecnoForSample(batchno, orderattribute);
        }

        public bool CheckSpecNorm(string batchno, string orderattribute, string val, string pogr)
        {
            return Specs.CheckRaschNormForSample(batchno, orderattribute, "fororderattributeselect", val, pogr, false);
        }
        public bool CheckSpecNorm(string diap, string val, string pogr)
        {
            double pg = TryToNumber(pogr, 99999999);
            double vs = TryToNumber(val.Replace(">", "").Replace("<", ""), 99999999);
            if (Math.Abs(vs - 99999999) <= 0.0000000001)
                return false;
            if (Math.Abs(pg - 99999999) <= 0.0000000001)
                pg = 0;
            if ((CalcModule.getLowerLimit(diap) - pg) < vs && (CalcModule.getUpperLimit(diap) + pg) >= vs)
                return false;
            else return true;
            //CREATE OR REPLACE FUNCTION SARNEWPROD."ISNOTINDIAPPOGR" (diap varchar2,c varchar2, pog varchar2) return number DETERMINISTIC as
            //vs number;
            //pg number;
            //begin
            //    pg:=trytonumber(pog, 99999999);
            //    vs:=trytonumber(replace(replace(c,'>',''),'<',''), 99999999);

            //    if vs=99999999 then
            //      return 0;
            //    end if;
            //    if pg=99999999 then
            //      pg:=0;
            //    end if;
            //    if ((GetLowerLimit(diap) - pg) < vs and (GetUpperLimit(diap) + pg) >= vs) then
            //        return 0;
            //    else
            //        return 1;
            //    end if;
            //end;
        }
        public string ReplaceDigitsTo(string str, string repl)
        {
            string s = str;
            for (int i = 1; i <= 9; i++)
            {
                s = s.Replace(i.ToString(), repl);
            }
            return s;
        }

        public string GetNumFormatByNumber(string str)
        {
            string vs = ReplaceDigitsTo(str, "0").Replace(',', '.').TrimStart('0').Replace("-", "");
            if (vs.StartsWith("."))
            {
                return "0" + vs;
            }
            else
            {
                return vs;
            }
        }
        public object GetSomething(Microsoft.Office.Interop.Excel.Range r, string lab, string docdescription, string key1, string Value1, string whattoget, string sortmetadataname, string DateTo)
        {
            string docid;
            docid = GetSomeDocid(r, lab, docdescription, key1, Value1, sortmetadataname, DateTo);
            if (docid != string.Empty)
            {
                string res = ei.GetWatersInterop().getBatchMetaDataByKeyForCurUser(docid, whattoget);
                if (Math.Abs(WatersInterop.tryToNumber(res, 99999999) - 99999999) >= 0.0000000001)
                {
                    return WatersInterop.tryToNumber(res, 99999999);
                }
            }
            return "";
        }
        public double GetGiriPopr(string docid, Microsoft.Office.Interop.Excel.Range nomrange)
        {
            double popr;
            if (docid != string.Empty)
            {
                double getGiriPopr = 0;
                for (int i = 1; i <= nomrange.Rows.Count; i++)
                {
                    string nom = ExcelInteractClass.GetCellValue(nomrange, i, 1).ToString();
                    if (nom != string.Empty)
                    {
                        popr = WatersInterop.tryToNumber(ei.GetWatersInterop().getBatchMetaDataByKeyForCurUser(docid, nom), -999);
                        if (Math.Abs(popr - -999) >= 0.0000000001)
                        {
                            getGiriPopr += popr;
                        }
                    }
                }
                return getGiriPopr;
            }
            return 0;
        }


        public object CalcConcentration(string docid, double x, out double y1, out double y2)
        {
            return CalcConcentrationByNaim(docid, x, out y1, out y2, "СилаТока", "Концентрация");
        }

        public object CalcVolume(string docid, double x)
        {
            return CalcFunctions.CalcVolume(docid, x);
        }

        public object CalcConcentrationByNaim(string docid, double x, out double y1, out double y2, string xnaim, string ynaim)
        {
            double x1 = 0, x2 = 0, arrIx = 0, arrC = 0;

            y1 = 0;
            y2 = 0;
            string strC = ei.GetWatersInterop().getBatchMetaDataByKeyForCurUser(docid, ynaim);
            string strIx = ei.GetWatersInterop().getBatchMetaDataByKeyForCurUser(docid, xnaim);

            string[] strArrIx = strIx.Split('|');
            string[] strArrC = strC.Split('|');

            for (int i = 0; i < strArrIx.Length; i++)
            {
                arrIx = WatersInterop.tryToNumber(strArrIx[i], 99999999);
                arrC = WatersInterop.tryToNumber(strArrC[i], 99999999);
                if (x < arrIx)
                {
                    x2 = arrIx; x1 = WatersInterop.tryToNumber(strArrIx[i - 1], 9999999);
                    y2 = arrC; y1 = WatersInterop.tryToNumber(strArrC[i - 1], 9999999);

                    return ei.interpol(x, x1, y1, x2, y2);
                }
                if (Math.Abs(x - arrIx) <= 0.0000000001)
                {
                    x1 = arrIx; x2 = WatersInterop.tryToNumber(strArrIx[i + 1], 9999999);
                    y1 = arrC; y2 = WatersInterop.tryToNumber(strArrC[i + 1], 9999999);

                    return ei.interpol(x, x1, y1, x2, y2);
                }
            }
            return null;
        }


        public double CalcClassChist2(double index)
        {
            return CalcFunctions.CalcClassChist2(index);
        }
        public double ReturnPosition(string All, string key)
        {
            string[] Arr = All.Split('|');
            for (int i = 0; i < Arr.Length; i++)
            {
                if (Arr[i].ToLower().Contains(key))
                {
                    return i;
                }
            }
            return -1;
        }

        public string SplitData(string str, int index, string delimeter)
        {
            return CalcFunctions.SplitData(str, index, delimeter);
        }

        public object CalcMaxConcentration(string docid, double x)
        {
            double y1 = 0;
            double y2 = 0;
            CalcConcentration(docid, x, out y1, out y2);
            return y2;
        }

        public double TRosCalc(double x, Microsoft.Office.Interop.Excel.Range xr, Microsoft.Office.Interop.Excel.Range yr)
        {
            double x1 = 0, y1 = 0, x2 = 0, y2 = 0;
            for (int r = 1; r <= xr.Rows.Count - 1; r++)
            {
                if (ExcelInteractClass.GetCellValue(xr, r, 1) is double && ExcelInteractClass.GetCellValue(yr, r, 1) is double &&
                    x >= (double)ExcelInteractClass.GetCellValue(xr, r, 1))
                {
                    x1 = (double)ExcelInteractClass.GetCellValue(xr, r, 1);
                    y1 = (double)ExcelInteractClass.GetCellValue(yr, r, 1);
                }
                if (ExcelInteractClass.GetCellValue(xr, r + 1, 1) is double && ExcelInteractClass.GetCellValue(yr, r + 1, 1) is double
                    && x <= (double)ExcelInteractClass.GetCellValue(xr, r + 1, 1))
                {
                    x2 = (double)ExcelInteractClass.GetCellValue(xr, r + 1, 1);
                    y2 = (double)ExcelInteractClass.GetCellValue(yr, r + 1, 1);
                    break;
                }

            }
            return ei.interpol(x, x1, y1, x2, y2);
        }
        public string GetProductIdByBatchNo(string batchno)
        {
            return wi.getProductId(wi.getBatchMetaDataByKeyForCurUser(batchno, "Анализируемый объект"));
        }

        public string GetBatchInstrumentsByOrderattr(string batchno, string orderattribute, string instrtype, string instrdelimeter)
        {
            return GetBatchInstruments(batchno, CalcModule.getMethodIdFromKey(orderattribute)
                    , CalcModule.getTestIdFromKey(orderattribute), instrtype, instrdelimeter);
        }
        public string GetBatchInstrumentsInfoByOrderattr(string batchno, string orderattribute, string instrtype, string instrdelimeter, string formatstr)
        {
            return GetBatchInstrumentsInfo(batchno, CalcModule.getMethodIdFromKey(orderattribute)
                    , CalcModule.getTestIdFromKey(orderattribute), instrtype, instrdelimeter, formatstr);
        }
        public string GetBatchInstruments(string batchno, string methoddescr, string testmeta, string instrtype, string instrdelimeter)
        {
            return GetBatchInstrumentsInfo(batchno, methoddescr, testmeta, instrtype, instrdelimeter, "{Description} ({ProdNumber})");
        }
        /// <summary>
        /// Возвращает информацию по приборам привязанным к пробе
        /// </summary>
        /// <param name="batchno">Проб</param>
        /// <param name="methoddescr">Метод</param>
        /// <param name="testmeta">Показатель</param>
        /// <param name="instrtype">Тип оборудования</param>
        /// <param name="instrdelimeter">Разделитель оборудования в возвращаемой строке</param>
        /// <param name="formatstr">Задаёт формат возвращаемой строки, можно указывать любые свойства класса Instrument, например,
        /// {DateNextPov} - вернёт дату следующей поверки</param>
        /// <returns></returns>
        public string GetBatchInstrumentsInfo(string batchno, string methoddescr, string testmeta, string instrtype, string instrdelimeter, string formatstr)
        {
            List<IBatchOrderAttribute> data = BatchOrderAttributes.Instance.getBatchMetaData(batchno);

            return Utils.TypesUtils.ColToString(data.Where(b => b.Orderattribute.StartsWith(CalcModule.GetMetaNameFromMethodTestId(methoddescr, "_СИ_ИД_"))
                                                                || b.Orderattribute.StartsWith(CalcModule.GetMetaNameFromMethodTestId(methoddescr, testmeta) + "_СИ_ИД_"))
                                                   .Select(b =>
                                                    !string.IsNullOrEmpty(GetInstrumentInformation(b.Wert1, formatstr)) ?
                                                    GetInstrumentInformation(b.Wert1, formatstr) :
                                                    formatstr == "{Description}" ? data.FirstOrDefault(p => p.Orderattribute.Contains(String.Format("_СИ_Описание_{0}", b.Orderattribute.Split('_').Last()))).Wert1
                                                    : formatstr == "{ProdNumber}" ? data.FirstOrDefault(p => p.Orderattribute.Contains(String.Format("_СИ_ЗавН_{0}", b.Orderattribute.Split('_').Last()))).Wert1 : "Данные из справочника удалены"
                                                    ).ToList()
                                , instrdelimeter);
        }

        public string GetInstrumentInformation(string instrumentid, string formatstr)
        {
            LIMSClasses.Interfaces.IInstrument instrument = Instruments.GetInstrumentById(instrumentid);

            return instrument != null ? instrument.ToFormatString(formatstr) : "";
        }
        public string GetInstrumentInformationByPN(string worknumber, string formatstr)
        {
            return Instruments.GetInstrumentByWorknumber(worknumber).ToFormatString(formatstr);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="batchnos">Номера проб разделённые через запятую</param>
        /// <param name="methoddescr">ИД метода</param>
        /// <param name="testmeta">Имя метаданных показателя</param>
        /// <param name="instrtype">Наименование типа оборудования</param>
        /// <param name="instrdelimeter">Разделитель оборудования в выходной строке</param>
        /// <param name="formatstr">Формат выходной строки</param>
        /// <returns>Строка содержащая всё оборудование разделённое через указанный разделитель и в заданном формате</returns>
        public string GetBatchInstrumentsInfo2(string batchnos, string methoddescr, string testmeta, string instrtype, string instrdelimeter, string formatstr)
        {
            /*string[] data = ei.GetWatersInterop().getBatchMetaData(ei.GetWatersInterop().GetCurrentUserId()
                                , "", batchno, true, "fororderattributeselect");
             */

            string[] arrayOfBathcs = batchnos.Split(',');

            if (arrayOfBathcs == null || arrayOfBathcs.Length <= 0) return "Ошибка входных параметров!";

            List<string> instruments = new List<string>();

            foreach (string batchno in arrayOfBathcs)
            {
                List<IBatchOrderAttribute> data = BatchOrderAttributes.Instance.getBatchMetaData(batchno);
                foreach (IBatchOrderAttribute b in data)
                {
                    if (b.Orderattribute.StartsWith(CalcModule.GetMetaNameFromMethodTestId(b.Orderattribute.Split('_')[0], "_СИ_ИД_"))
                          || b.Orderattribute.StartsWith(CalcModule.GetMetaNameFromMethodTestId(b.Orderattribute.Split('_')[0], testmeta) + "_СИ_ИД_"))
                    {
                        LIMSClasses.Interfaces.IInstrument instr = Instruments.GetInstrumentById(b.Wert1);

                        if (instr != null && (instrtype == string.Empty || instrtype == "%" || instr.Type == instrtype))
                        {
                            if (!instruments.Contains(instr.ToFormatString(formatstr)))
                                instruments.Add(
                                instr.ToFormatString(formatstr)
                                );
                        }
                    }
                }
            }
            return TypesUtils.ColToString(instruments, instrdelimeter);
        }

        /// <summary>
        /// Возвращает информацию по приборам привязанным к пробе
        /// </summary>
        /// <param name="batchno">Проб</param>
        /// <param name="methoddescr">Метод</param>
        /// <param name="testmeta">Показатель</param>
        /// <param name="instrNameDescriprion">Наименование оборудования</param>
        /// <param name="instrdelimeter">Разделитель оборудования в возвращаемой строке</param>
        /// <param name="formatstr">Задаёт формат возвращаемой строки, можно указывать любые свойства класса Instrument, например,
        /// {DateNextPov} - вернёт дату следующей поверки</param>
        /// <returns></returns>
        public string GetBatchInstrumentsInfoByNameDescription(string batchno, string methoddescr, string testmeta, string instrNameDescriprion, string instrdelimeter, string formatstr)
        {
            /*string[] data = ei.GetWatersInterop().getBatchMetaData(ei.GetWatersInterop().GetCurrentUserId()
                                , "", batchno, true, "fororderattributeselect");
             */
            List<IBatchOrderAttribute> data = BatchOrderAttributes.Instance.getBatchMetaData(batchno);
            foreach (IBatchOrderAttribute b in data)
            {
                if (b.Orderattribute.StartsWith(CalcModule.GetMetaNameFromMethodTestId(b.Orderattribute.Split('_')[0], "_СИ_ИД_"))
                      || b.Orderattribute.StartsWith(CalcModule.GetMetaNameFromMethodTestId(b.Orderattribute.Split('_')[0], testmeta) + "_СИ_ИД_"))
                {
                    try
                    {
                        LIMSClasses.Interfaces.IInstrument instr = Instruments.GetInstrumentById(b.Wert1);
                        if (instr != null && instr.Description.Contains(instrNameDescriprion))
                        {
                            return instr.ToFormatString(formatstr);
                        }
                    }
                    catch (Exception)
                    {
                        return "";
                    }

                }
            }
            return "";
        }

        public string GetEnteredOpredCount(string batchno, string orderattribute)
        {
            for (int i = 4; i >= 0; i--)
            {
                if (ei.GetWatersInterop().getBatchMetaDataByKeyForCurUser(batchno, orderattribute + i.ToString() + "опр") != string.Empty)
                {
                    return i.ToString();
                }
            }
            return "0";
        }
        public string ConvertNumberToTextGenitive(long number)
        {
            return RuDateAndMoneyConverter.NumeralsToTxt((long)number, TextCase.Genitive, false, false);

        }
        public string ConvertNumberToTextNominative(long number)
        {
            return RuDateAndMoneyConverter.NumeralsToTxt((long)number, TextCase.Nominative, false, false);

        }

        public int CalcPageNumber(Microsoft.Office.Interop.Excel.Worksheet ws)
        {

            Microsoft.Office.Interop.Excel.Worksheet aws = (Microsoft.Office.Interop.Excel.Worksheet)ws.Application.ActiveSheet;
            ((Microsoft.Office.Interop.Excel._Worksheet)ws).Activate();
            ws.Application.ActiveWindow.View = Microsoft.Office.Interop.Excel.XlWindowView.xlNormalView;
            ws.ResetAllPageBreaks();
            ws.Application.ActiveWindow.View = Microsoft.Office.Interop.Excel.XlWindowView.xlPageBreakPreview;
            int count = 0;
            int VPcount = ws.VPageBreaks.Count;
            int HPcount = ws.HPageBreaks.Count;
            if (VPcount == 0)
                count = HPcount + 1;
            else
                count = (VPcount + 1) * (HPcount + 1);
            ws.Application.ActiveWindow.View = Microsoft.Office.Interop.Excel.XlWindowView.xlNormalView;
            ((Microsoft.Office.Interop.Excel._Worksheet)aws).Activate();
            return count;
        }
        public string SetFooterText(Microsoft.Office.Interop.Excel.Worksheet ws, string text, string data)
        {
            return ws.PageSetup.CenterHeader = text + " " + data;

        }

        public void AddDataToDataForMethods(Microsoft.Office.Interop.Excel.Range r, Microsoft.Office.Interop.Excel.Range argsX1,
            Microsoft.Office.Interop.Excel.Range argsX2, string MethodDescription)
        {
            object[,] temp_arrArgsX1 = (object[,])argsX1.Value; // Temperature
            object[,] temp_arrArgsX2 = (object[,])argsX2.Value; // Precision
            object[] arrArgsX1 = new object[temp_arrArgsX1.GetLength(0) > temp_arrArgsX1.GetLength(1) ? temp_arrArgsX1.GetLength(0) : temp_arrArgsX1.GetLength(1)];
            object[] arrArgsX2 = new object[temp_arrArgsX2.GetLength(0) > temp_arrArgsX2.GetLength(1) ? temp_arrArgsX2.GetLength(0) : temp_arrArgsX2.GetLength(1)];
            if (temp_arrArgsX1.GetLength(0) > temp_arrArgsX1.GetLength(1))
            {
                for (int i = 0; i < temp_arrArgsX1.GetLength(0); i++) { arrArgsX1[i] = temp_arrArgsX1[i + 1, 1]; }
            }
            else
            {
                for (int i = 0; i < temp_arrArgsX1.GetLength(1); i++) { arrArgsX1[i] = temp_arrArgsX1[1, i + 1]; }
            }
            if (temp_arrArgsX2.GetLength(0) > temp_arrArgsX2.GetLength(1))
            {
                for (int i = 0; i < temp_arrArgsX2.GetLength(0); i++) { arrArgsX2[i] = temp_arrArgsX2[i + 1, 1]; }
            }
            else
            {
                for (int i = 0; i < temp_arrArgsX2.GetLength(1); i++) { arrArgsX2[i] = temp_arrArgsX2[1, i + 1]; }
            }

            object[,] arrValues = (object[,])r.Value;
            string[,] strarrValues = new string[arrValues.GetLength(0), arrValues.GetLength(1)];

            for (int i = 0; i < arrValues.GetLength(0); i++)
            {
                for (int j = 0; j < arrValues.GetLength(1); j++)
                {
                    strarrValues[i, j] = ExcelInteractClass.GetObjectStringValue(arrValues[i + 1, j + 1]);
                }
            }
            TableDataForMethods.InsertDataToTableDataForMethods(MethodDescription, arrArgsX1.ToList().ConvertAll(item => ExcelInteractClass.GetObjectStringValue(item)).ToArray()
, arrArgsX2.ToList().ConvertAll(item => ExcelInteractClass.GetObjectStringValue(item)).ToArray()
, strarrValues);
        }



        public double Getchcount(Microsoft.Office.Interop.Excel.Range xrange)
        {
            int c;
            double avg;
            avg = 0;
            c = 0;
            for (int i = 1; i <= xrange.Rows.Count; i++)
            {
                for (int j = 1; j <= xrange.Columns.Count; j++)
                {
                    if (ExcelInteractClass.GetCellValue(xrange, i, j).ToString() != string.Empty
                        && Math.Abs(WatersInterop.tryToNumber(ExcelInteractClass.GetCellValue(xrange, i, j).ToString(), 99999999) - 99999999) >= 0.0000000001)
                    {
                        avg += WatersInterop.tryToNumber(ExcelInteractClass.GetCellValue(xrange, i, j).ToString(), 99999999);
                        c++;
                    }
                }
            }
            if (c > 0)
            {
                avg = avg / c;
                if (Math.Abs(Math.Truncate(avg) - avg) < 0.00000000001)
                {
                    return Math.Truncate(avg);
                }
                else
                {
                    return Math.Truncate(avg) + 1;
                }
            }
            else
            {
                return 0;
            }
        }

        [ComVisible(true)]
        public string GetMethodSistprecU(Microsoft.Office.Interop.Excel.Range r, string product, string methodname, string testid, string lab, string description, double c)
        {
            return ReturnChar(r, VLKModule.GetMethodSistprecUpStr(r, product, methodname, testid, lab, description, c), c).ToString();
        }

        [ComVisible(true)]
        public string GetMethodSistprecD(Microsoft.Office.Interop.Excel.Range r, string product, string methodname, string testid, string lab, string description, double c)
        {
            return ReturnChar(r, VLKModule.GetMethodSistprecDownStr(r, product, methodname, testid, lab, description, c), c).ToString();
        }

        public string getPovtVLKBatchIdBySourceBatchNo(string SourceBatchNo)
        {
            string[] vs = wi.getPovtVLKBatchIdBySourceBatchNo(SourceBatchNo);
            return vs != null && vs.Length >= 1 ? vs[0] : "";
        }

        public double getCellWithResultFromSheet(string batchno, string testname, string methodname)
        {
            double x = 0;
            Microsoft.Office.Interop.Excel.Application app = ei.GetApplication();
            foreach (Microsoft.Office.Interop.Excel.Worksheet ws in app.ActiveWorkbook.Sheets)
            {

                if (ws.Name.Contains(methodname) && ei.GetBatchnoFromSheet(ws) == batchno)
                    x = Convert.ToDouble(ei.GetWorksheetNamedMetaRangeValue(ws, testname, 99999999999));
            }
            return x;
        }

        #region Работа с Арбитражными пробами
        /// <summary>
        /// Редактирование параметров арбитражной пробы
        /// </summary>
        /// <param name="BatchNo">Номер пробы</param>
        /// <param name="r"></param>
        public void arbitriesSampleForm()
        {
            Microsoft.Office.Interop.Excel.Worksheet ws = ei.GetSelection().Worksheet;
            Microsoft.Office.Interop.Excel.Range _selectedRange = ei.GetSelection();
            Microsoft.Office.Interop.Excel.Range _probCol;
            _probCol = ei.GetWorksheetNamedRange(ws, "Для номера пробы");
            Microsoft.Office.Interop.Excel.Range _probNum = ExcelInteractClass.GetRangeByRC(ws, _selectedRange.Row, _probCol.Column, _selectedRange.Row, _probCol.Column);

            if (_probNum.Value2 == null) return;

            using (var form = new ArbitriesSampleForm(ExcelInteractClass.GetObjectStringValue(_probNum.Value2)))
            {
                form.ShowDialog();
                object temp = _probNum.Value2;
                _probNum.Value2 = "-999999";
                _probNum.Value2 = temp;
            }
        }

        public string ReturnArbitriesValue(string batchNo, string column)
        {
            IArbitriesSamplesDAO dao = LIMSClasses.Configuration.ArbitriesSamplesDAO;
            return dao.ReturnArbitriesValue(batchNo, column);
        }

        public void MakeArbitriesIsFinished()
        {
            Microsoft.Office.Interop.Excel.Worksheet ws = ei.GetSelection().Worksheet;
            string res = EnterTextDataForm.GetTextValue("ФИО снявшего с хранения", -1, false);
            if (string.IsNullOrEmpty(res)) return;

            Microsoft.Office.Interop.Excel.Range _selectedRange = ei.GetSelection();
            ArbitriesSampleForm form;
            for (int i = _selectedRange.Row; i < _selectedRange.Row + _selectedRange.Rows.Count; i++)
            {
                Microsoft.Office.Interop.Excel.Range _probCol;
                Microsoft.Office.Interop.Excel.Range _dateAnalyse;
                Microsoft.Office.Interop.Excel.Range _selectedRow = _selectedRange.Range[ws.Cells[_selectedRange.Row, 1], ws.Cells[_selectedRange.Row, ws.UsedRange.Columns.Count]];
                _probCol = ei.GetWorksheetNamedRange(ws, "Для номера пробы");
                Microsoft.Office.Interop.Excel.Range _probNum = ExcelInteractClass.GetRangeByRC(ws, i, _probCol.Column, i, _probCol.Column);

                List<string> _listOfNamedCells = ei.GetNamesInRange(_selectedRow);
                if (_probNum.Value2 == null)
                    return;

                form = new ArbitriesSampleForm(ExcelInteractClass.GetObjectStringValue(_probNum.Value2));
                form.loadInformation();
                form.DataSnatiyaSHraneniya = SystemTime.Current;
                form.FIOSnyavshegoSHraneniya = res;
                form.OKBut_Click("", new EventArgs());
                object temp = _probNum.Value2;
                _probNum.Value2 = "-999999";
                _probNum.Value2 = temp;
            }
        }
        #endregion

        public object CalcVlagnost(Microsoft.Office.Interop.Excel.Range range, double tsuh, double tvlag)
        {
            object[,] data = ExcelInteractClass.GetRangeFormulas(range.Worksheet.UsedRange);
            int r1 = -1;
            int c1 = -1;
            double tdelta = tsuh - tvlag;
            double delta = 99999999;

            for (int r = data.GetLowerBound(0); r <= data.GetUpperBound(0); r++)
            {
                double vsd = WatersInterop.tryToNumber(ExcelInteractClass.GetObjectStringValue(data[r, data.GetLowerBound(0)]), -99999999);
                if (Math.Abs(vsd - -99999999) >= 0.0000000001 && vsd < tsuh && delta > (tsuh - vsd))
                {
                    delta = (tsuh - vsd);
                    r1 = r;
                }
            }
            if (r1 != -1)
            {
                delta = 99999999;
                for (int c = data.GetLowerBound(1); c <= data.GetUpperBound(1); c++)
                {
                    double vsd = WatersInterop.tryToNumber(ExcelInteractClass.GetObjectStringValue(data[data.GetLowerBound(1), c]), -99999999);
                    if (Math.Abs(vsd - -99999999) >= 0.0000000001 && vsd < tdelta && delta > (tdelta - vsd))
                    {
                        delta = (tdelta - vsd);
                        c1 = c;
                    }
                }
                if (c1 != -1)
                {
                    double tsuh1, tsuh2, vl1, vl2, vl3, vl4, tdelta1, tdelta2;
                    tsuh1 = WatersInterop.tryToNumber(ExcelInteractClass.GetObjectStringValue(data[r1, data.GetLowerBound(0)]), -99999999);
                    tsuh2 = WatersInterop.tryToNumber(ExcelInteractClass.GetObjectStringValue(data[r1 + 1, data.GetLowerBound(0)]), -99999999);
                    tdelta1 = WatersInterop.tryToNumber(ExcelInteractClass.GetObjectStringValue(data[data.GetLowerBound(1), c1]), -99999999);
                    tdelta2 = WatersInterop.tryToNumber(ExcelInteractClass.GetObjectStringValue(data[data.GetLowerBound(1), c1 + 1]), -99999999);
                    vl1 = WatersInterop.tryToNumber(ExcelInteractClass.GetObjectStringValue(data[r1, c1]), -99999999);
                    vl2 = WatersInterop.tryToNumber(ExcelInteractClass.GetObjectStringValue(data[r1 + 1, c1]), -99999999);
                    vl1 = CalcVl(tsuh, tsuh1, tsuh2, vl1, vl2);
                    vl3 = WatersInterop.tryToNumber(ExcelInteractClass.GetObjectStringValue(data[r1, c1 + 1]), -99999999);
                    vl4 = WatersInterop.tryToNumber(ExcelInteractClass.GetObjectStringValue(data[r1 + 1, c1 + 1]), -99999999);
                    vl3 = CalcVl(tsuh, tsuh1, tsuh2, vl3, vl4);
                    return ei.interpol(tdelta, tdelta1, Math.Round(vl1, 0), tdelta2, Math.Round(vl3, 0));
                }
                else
                {
                    return string.Format("В таблице не найдено значение разницы температур {0}", tdelta.ToString());
                }
            }
            else
            {
                return string.Format("В таблице не найдено значение темп.сух. терм. {0}", tsuh.ToString());
            }
        }

        private double CalcVl(double tsuh, double tsuh1, double tsuh2, double vl1, double vl2)
        {
            //добавил условие в соответствии с п. 5.6 инструкции по ВИТ, что если влажность меняется больше 1% при изменении темп. на 1 градус
            //, то интерполируем иначе округляем.            
            if (Math.Abs(vl2 - vl1) / Math.Abs(tsuh2 - tsuh1) > 1)
            {
                vl1 = ei.interpol(tsuh, tsuh1, vl1, tsuh2, vl2);
            }
            else
            {
                if (Math.Abs(tsuh - tsuh1) >= Math.Abs(tsuh - tsuh2))
                {
                    vl1 = vl2;
                }
            }
            return vl1;
        }

        public double Substring(string Val, int Count, int CountOfDecimalNum)
        {
            int delSeparator = Val.IndexOf(CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator);
            Val = Val.Replace(CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator, string.Empty) + "0000000";
            int findFirstDigit = -1;
            for (int i = 0; i < Val.Length; i++)
            {
                int tryInt = -1;
                if (int.TryParse(Val.Substring(i, 1), out tryInt))
                {
                    if (tryInt > 0)
                    {
                        findFirstDigit = i;
                        break;
                    }
                }
            }
            Val = Val.Substring(0, findFirstDigit + Count);
            Val = Val.Insert(delSeparator, CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator);
            return Math.Round(double.Parse(Val), CountOfDecimalNum);
        }

        public string GetPurposeDescriptionById(string purposeid)
        {
            return Purposes.GetPurpose(purposeid).Description;
        }

        public bool CheckRaschNormForSample(string batchno, string orderattribute, string table, string res, string pogr, bool woreqfalse)
        {
            return Specs.CheckRaschNormForSample(batchno, orderattribute, table, res, pogr, woreqfalse);
        }

        public string GetPovtBatchID(string batchno)
        {
            return WatersInterop.GetWi().GetPovtBatchID(batchno);
        }

        /// <summary>
        /// Метод 
        /// </summary>
        /// <param name="batchno"></param>
        /// <param name="groupCount"></param>
        /// <param name="arrayOfValues"></param>
        /// <param name="arrayOfBatches"></param>
        /// <returns></returns>
        public double CalcAndPutScope(string batchno, string groupCount, Microsoft.Office.Interop.Excel.Range arrayOfValues, Microsoft.Office.Interop.Excel.Range arrayOfBatches)
        {
            double returnedValue = 0;
            Dictionary<string, double> keyVal = new Dictionary<string, double>();

            object[,] batches = (object[,])arrayOfBatches.Cells.Value;
            object[,] values = (object[,])arrayOfValues.Cells.Value;

            List<string> tempList = new List<string>();
            double minR = double.MaxValue, maxR = double.MinValue;
            double R = 0;
            for (int i = 1, j = 1; i <= batches.Length; i++, ++j)
            {
                if (values[i, 1] == null || batches[i, 1] == null || batches[i, 1].ToString() == "0" || values[i, 1].ToString() == "0") continue;
                if (Convert.ToDouble(values[i, 1]) < minR) minR = Convert.ToDouble(values[i, 1]);
                if (Convert.ToDouble(values[i, 1]) > maxR) maxR = Convert.ToDouble(values[i, 1]);

                tempList.Add(batches[i, 1].ToString());

                if (tempList.Count == Convert.ToInt32(groupCount))
                {
                    R = maxR - minR;
                    foreach (string key in tempList)
                    {
                        if (!keyVal.ContainsKey(key))
                            keyVal.Add(key, R);
                        else keyVal[key] = R;
                    }
                    j = 1;
                    tempList = new List<string>();
                    minR = double.MaxValue;
                    maxR = double.MinValue;
                }
                else
                {
                    if (!keyVal.ContainsKey(batches[i, 1].ToString()))
                        keyVal.Add(batches[i, 1].ToString(), -999999);
                }
            }

            if (keyVal.ContainsKey(batchno))
                returnedValue = keyVal[batchno];
            else returnedValue = -999999;

            return returnedValue;
        }
        /// <summary>
        /// Получение данных из объекта Products по заданному id и pattern
        /// </summary>
        /// <param name="productid">ИД продукта</param>
        /// <param name="pattern">Поле таблицы</param>
        /// <returns>Значение из поля pattern</returns>
        public string GetProductDataById(string productid, string pattern)
        {
            return Products.GetProduct(productid).ToFormatString(pattern);
        }
        /// <summary>
        /// Получение данных из объекта MethodTests по заданному id и pattern
        /// </summary>
        /// <param name="testid">ИД связи MethodTests</param>
        /// <param name="pattern">Поля таблиц MetodTests</param>
        /// <returns>Значение из поля pattern</returns>
        public string GetTestDataByTestId(string testid, string pattern)
        {
            return MethodTests.GetById(testid).ToFormatString(pattern);
        }
        /// <summary>
        /// 
        /// </summary>Получение данных из объекта Methods по заданному id и pattern
        /// <param name="methodid">ИД метода</param>
        /// <param name="pattern">Поле таблицы</param>
        /// <returns>Значение из поля pattern</returns>
        public string GetMethodDataById(string methodid, string pattern)
        {
            return Methods.GetMethodById(methodid).ToFormatString(pattern);
        }
        public string GetMethodCharStrForBatchno(string batchno,string metaname,string description, string c, double k, string pattern)
        {
            return MethodChars.Instance.GetMethodCharStr(
                    MethodChars.Instance.GetMethodChars(Samples.GetSample(batchno, Users.GetCurrentUser()), MethodTests.GetByMetaName(metaname), description)
                    ,new List<string> { c },k,pattern);
        }
        public double GetMethodCharForBatchno(string batchno, string metaname, string description, string c, double k, string pattern)
        {
            return MethodChars.Instance.GetMethodChar(
                    Samples.GetSample(batchno, Users.GetCurrentUser()), MethodTests.GetByMetaName(metaname), description,c , k, pattern);
        }

    } 
}
