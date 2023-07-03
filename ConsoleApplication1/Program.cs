using System.IO;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Xml;
using System;
using System.Linq;
using System.Xml.Linq;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace ConsoleApplication1
{
    class Model
    {
        public string FilePath;
        public Model(string filePath) {
            this.FilePath = filePath;
        }
        public virtual void XMItoCSharp(XmlElement root, bool isPackage) 
        {
            return;
        }
        public virtual void XMItoCSharp()
        {
            return;
        }
    }
    /*class TypeVerificator
    {
        private static bool FindActivePackageEl(XmlNodeList xPackagedList, string type)
        {
            foreach (XmlNode node in xPackagedList)
            {
                var attr = node.Attributes["xsi:type"];
                if (attr == null) continue;
                if (attr.Value.Equals(type))
                    return true;
            }
            return false;
        }

        public static bool VerificateDiagramType(UCDModel model)
        {
            var root = new XmlDocument();
            root.Load(model.FilePath);

            XmlNodeList xPackagedList = root.GetElementsByTagName("packagedElement");

            if (FindActivePackageEl(xPackagedList, "uml:UseCase") || root.GetElementsByTagName("ownedUseCase").Count != 0)
                return true;

            return false;
        }
    }*/
    internal static class TypeDefiner
    {
        private static bool FindActivePackageEl(XmlNodeList xPackagedList, string type)
        {
            foreach (XmlNode node in xPackagedList)
            {
                var attr = node.Attributes["xsi:type"];
                if (attr == null) continue;
                if (attr.Value.Equals(type))
                    return true;
            }
            return false;
        }
        private static bool FindActiveElement(XmlNodeList xNodes, string name)
        {
            foreach (XmlNode node in xNodes)
            {
                if (node.Name == name || FindActiveElement(node.ChildNodes, "ownedUseCase"))
                    return true;
            }
            return false;
        }
        public static List<string> DefineDiagramType(string filePath)
        {
            var root = new XmlDocument();
            root.Load(filePath);

            List<string> types = new List<string>();

            XmlNodeList xPackagedList;
            XmlNodeList xNodes;
            try
            {
                xPackagedList = root.GetElementsByTagName("packagedElement");
                xNodes = root.ChildNodes;
            }
            catch (NullReferenceException)
            {
                types.Add("Неопределено");
                return types;
            }

            if (FindActivePackageEl(xPackagedList, "uml:Activity"))
                types.Add("AD");

            if (FindActivePackageEl(xPackagedList, "uml:UseCase") || FindActiveElement(xNodes, "ownedUseCase"))
                types.Add("UCD");

            if (types.Count == 0)
                types.Add("Неопределено");

            return types;
        }
    }
    class Program
    {
        static void ConsoleOutput(UCDMetricCalculator mc, string filePath)
        {
            Console.WriteLine("Выбранный файл: " + filePath);
            Console.WriteLine("\nВычисленные метрики:\r\nNOUC: " + mc.nouc + "\r\nNOA: " + mc.noa + "\r\nNOUCA: " + Math.Round(mc.nouca, 2)
                + "\r\nUC2: " + mc.ucSecond + "\r\nUC3: " + Math.Round((decimal)mc.ucThird, 2) + "\r\nUC4: " + Math.Round((decimal)mc.ucFourth, 2) + "\r\nCTE: " + mc.CTE
                + "\r\nCIE: " + mc.CIE + "\r\nCucd: " + mc.cUcd);
        }
        static void FileUCDOutput(StreamWriter sw, UCDMetricCalculator mc, string filePath)
        {
            sw.WriteLine("Выбранный файл: " + filePath);
            sw.WriteLine("Вычисленные метрики:\r\nNOUC: " + mc.nouc + "\r\nNOA: " + mc.noa + "\r\nNOUCA: " + Math.Round(mc.nouca, 2)
                + "\r\nUC2: " + mc.ucSecond + "\r\nUC3: " + Math.Round((decimal)mc.ucThird, 2) + "\r\nUC4: " + Math.Round((decimal)mc.ucFourth, 2) + "\r\nCTE: " + mc.CTE
                + "\r\nCIE: " + mc.CIE + "\r\nCucd: " + mc.cUcd);
            sw.WriteLine();
        }
        static void FileADOutput(StreamWriter sw, ADMetricCalculator mc, string filePath)
        {
            sw.WriteLine("Выбранный файл: " + filePath);
            sw.WriteLine("Вычисленные метрики:\r\nNoSw: " + mc.nosw + "\r\nNoAct: " + mc.noact + "\r\nNoDn: " + mc.nodn
                + "\r\nNoF: " + mc.nof + "\r\nNoJ: " + mc.noj + "\r\nNoE: " + mc.noe + "\r\nC: " + mc.comp);
            sw.WriteLine();
        }
        static void ExcelMetricsUCDOutput(Excel.Worksheet ws, UCDMetricCalculator mc, int curIndex, string filePath)
        {
            ws.Cells[curIndex, 1] = filePath.Split('\\')[filePath.Split('\\').Count() - 1];
            ws.Cells[curIndex, 2] = mc.nouc;
            ws.Cells[curIndex, 3] = mc.noa;
            ws.Cells[curIndex, 4] = Math.Round(mc.nouca, 2);
            ws.Cells[curIndex, 5] = mc.ucSecond;
            ws.Cells[curIndex, 6] = Math.Round(mc.ucThird, 2);
            ws.Cells[curIndex, 7] = Math.Round(mc.ucFourth, 2);
            ws.Cells[curIndex, 8] = mc.CTE;
            ws.Cells[curIndex, 9] = mc.CIE;
            ws.Cells[curIndex, 10] = mc.cUcd;
        }
        static void ExcelMetricsADOutput(Excel.Worksheet ws, ADMetricCalculator mc, int curIndex, string filePath)
        {
            ws.Cells[curIndex, 1] = filePath.Split('\\')[filePath.Split('\\').Count() - 1];
            ws.Cells[curIndex, 2] = mc.nosw;
            ws.Cells[curIndex, 3] = mc.noact;
            ws.Cells[curIndex, 4] = mc.nodn;
            ws.Cells[curIndex, 5] = mc.nof;
            ws.Cells[curIndex, 6] = mc.noj;
            ws.Cells[curIndex, 7] = mc.noe;
            ws.Cells[curIndex, 8] = mc.comp;
        }
        static void excelElOut(Excel.Worksheet wsDet, List<Element> delElems, ref int detIndex, string filePath, string methodName)
        {
            foreach (var el in delElems)
            {
                try
                {
                    wsDet.Cells[detIndex, 1] = filePath.Split('\\')[filePath.Split('\\').Count() - 1];
                    wsDet.Cells[detIndex, 2] = methodName;
                    wsDet.Cells[detIndex, 3] = el.Type == "uml:Actor" ? "Актор" : "Прецедент";
                    wsDet.Cells[detIndex, 4] = el.Name;
                    wsDet.Cells[detIndex, 5] = "-";
                    wsDet.Cells[detIndex, 6] = "-";
                    detIndex++;
                }
                catch
                {
                }
            }
        }
        static void excelElOut(Excel.Worksheet wsDet, List<BaseNode> delElems, ref int detIndex, string filePath, string methodName)
        {
            foreach (var el in delElems)
            {
                wsDet.Cells[detIndex, 1] = filePath.Split('\\')[filePath.Split('\\').Count() - 1];
                wsDet.Cells[detIndex, 2] = methodName;
                wsDet.Cells[detIndex, 3] = el.getType();
                wsDet.Cells[detIndex, 4] = el.getName() == "" ? el.getId() : el.getName();
                if (el.getType() == ElementType.FLOW)
                {
                    wsDet.Cells[detIndex, 5] = el.getSrc();
                    wsDet.Cells[detIndex, 6] = el.getTarget();
                }
                detIndex++;
            }
        }
        static void excelConnOut(Excel.Worksheet wsDet, List<Connection> delCons, List<Element> elems, ref int detIndex, string filePath, string methodName)
        {
            foreach (var conn in delCons)
            {
                try
                {
                    wsDet.Cells[detIndex, 1] = filePath.Split('\\')[filePath.Split('\\').Count() - 1];
                    wsDet.Cells[detIndex, 2] = methodName;
                    string conType = "";
                    switch (conn.Type)
                    {
                        case "uml:Association":
                            conType = "Ассоциация";
                            break;
                        case "include":
                            conType = "Включение";
                            break;
                        case "extend":
                            conType = "Расширение";
                            break;
                        default:
                            break;
                    }
                    wsDet.Cells[detIndex, 3] = conType;
                    wsDet.Cells[detIndex, 4] = "-";

                    Element elFrom = elems.Find(el => el.Id == conn.IdFrom);
                    Element elTo = elems.Find(el => el.Id == conn.IdTo);

                    wsDet.Cells[detIndex, 5] = elFrom == null ? "?" : elFrom.Name;
                    wsDet.Cells[detIndex, 6] = elTo == null ? "?" : elTo.Name;
                    detIndex++;
                }
                catch { }
            }
        }
        static int ExcelUCDFixesOutput(List<string> files, string path, string workDir)
        {
            //C:\\Users\\datsunnn\\Documents\\Visual Studio 2010\\Projects\\UMLMetrics\\UCD\\UCDFixes.xlsx
            if (File.Exists(workDir + "\\UCD\\UCDFixes.xlsx"))
            {
                try
                {
                    //C:\\Users\\datsunnn\\Documents\\Visual Studio 2010\\Projects\\UMLMetrics\\UCD\\UCDFixes.xlsx
                    File.Delete(workDir + "\\UCD\\UCDFixes.xlsx");
                }
                catch
                {
                    Console.WriteLine("\nНе удается создать Excel файл. Пожалуйста закройте открытый Excel файл с исправлениями в AD и попробуйте снова.");
                    Console.ReadLine();
                    return -1;
                }
            }

            var excelApp = new Excel.Application();
            var wb = excelApp.Workbooks.Add();

            while (wb.Sheets.Count > 1)
            {
                Excel.Worksheet ws = (Excel.Worksheet)wb.Sheets[1];
                ws.Delete();
            }

            Excel.Worksheet wsTotal = excelApp.Worksheets.Add();
            wsTotal.Name = "Общие сведения исправлений";

            Excel.Worksheet wsDet = excelApp.Worksheets.Add();
            wsDet.Name = "Подробности исправлений";

            Excel.Worksheet wst = (Excel.Worksheet)wb.Sheets[3];
            wst.Delete();

            wsTotal.Cells[1, 1] = "Название файла";
            wsTotal.Cells[1, 2] = "Количество удаленных дубликатов";
            wsTotal.Cells[1, 3] = "Количество удаленных неверно использованных ассоциаций";
            wsTotal.Cells[1, 4] = "Количество удаленных циклических связей";
            wsTotal.Cells[1, 5] = "Количество удаленных изолированных акторов";
            wsTotal.Cells[1, 6] = "Количество удаленных изолированных прецедентов";
            wsTotal.Cells[1, 7] = "Количество удаленных незавершенных связей";
            wsTotal.Cells[1, 8] = "Итог";

            wsDet.Cells[1, 1] = "Название файла";
            wsDet.Cells[1, 2] = "Вид исправления";
            wsDet.Cells[1, 3] = "Тип елемента";
            wsDet.Cells[1, 4] = "Название елемента";
            wsDet.Cells[1, 5] = "Первый конец связи";
            wsDet.Cells[1, 6] = "Второй конец связи";

            int totalFixes = 0;
            int totalIndex = 2;
            int detIndex = 2;
            foreach (var file in files)
            {
                if (Path.GetExtension(path + '\\' + file) == ".xmi" || Path.GetExtension(path + '\\' + file) == ".XMI")
                {
                    try
                    {
                        var forCheck = new XmlDocument();
                        forCheck.Load(file);
                    }
                    catch
                    {
                        continue;
                    }
                    UCDModel curModel = new UCDModel(file);
                    if (TypeDefiner.DefineDiagramType(file).Contains("UCD"))
                    {
                        List<Connection> conns = new List<Connection>();
                        List<Element> elems = new List<Element>();

                        foreach (var con in curModel.Conns)
                            conns.Add(con);
                        foreach (var elem in curModel.Elems)
                            elems.Add(elem);

                        List<Connection> delCons;
                        List<Element> delElems;

                        UCDFileFixer ff = new UCDFileFixer(conns, elems);

                        wsTotal.Cells[totalIndex, 1] = file.Split('\\')[file.Split('\\').Count() - 1];

                        int fileFixes = 0;

                        delElems = ff.RemoveDuplicatesElems(null);
                        delCons = ff.RemoveDuplicatesConns(null);

                        fileFixes += delElems.Count + delCons.Count;
                        wsTotal.Cells[totalIndex, 2] = fileFixes;

                        excelElOut(wsDet, delElems, ref detIndex, curModel.FilePath, "Удаление дубликатов");
                        excelConnOut(wsDet, delCons, elems, ref detIndex, curModel.FilePath, "Удаление дубликатов");

                        delCons = ff.RemoveMissusedAssociations(null);

                        fileFixes += delCons.Count;
                        wsTotal.Cells[totalIndex, 3] = delCons.Count;
                        excelConnOut(wsDet, delCons, elems, ref detIndex, curModel.FilePath, "Удаление неверно использованных ассоциаций");

                        delCons = ff.RemoveLoopedConns(null);

                        fileFixes += delCons.Count;
                        wsTotal.Cells[totalIndex, 4] = delCons.Count;
                        excelConnOut(wsDet, delCons, elems, ref detIndex, curModel.FilePath, "Удаление циклических связей");

                        delElems = ff.RemoveIsolatedActors(null);

                        fileFixes += delElems.Count;
                        wsTotal.Cells[totalIndex, 5] = delElems.Count;
                        excelElOut(wsDet, delElems, ref detIndex, curModel.FilePath, "Удаление изолированных акторов");

                        var tempTuple = ff.RemoveIsolatedUcs(null);
                        delElems = tempTuple.Item1;
                        delCons = tempTuple.Item2;
                        int tmpCount = delElems.Count + delCons.Count;
                        fileFixes += tmpCount;
                        wsTotal.Cells[totalIndex, 6] = tmpCount;
                        excelElOut(wsDet, delElems, ref detIndex, curModel.FilePath, "Удаление изолированных прецедентов");
                        excelConnOut(wsDet, delCons, delElems, ref detIndex, curModel.FilePath, "Удаление изолированных прецедентов");

                        delCons = ff.RemoveIncompleteConns(null);

                        fileFixes += delCons.Count;
                        wsTotal.Cells[totalIndex, 7] = delCons.Count;
                        excelConnOut(wsDet, delCons, elems, ref detIndex, curModel.FilePath, "Удаление незавершенных связей");

                        wsTotal.Cells[totalIndex, 8] = fileFixes;
                        totalFixes += fileFixes;
                        totalIndex++;
                    }
                }
            }
            wsTotal.Cells[(totalIndex + 1), 1] = "Общее кол-во";
            wsTotal.Cells[(totalIndex + 1), 2] = totalFixes;
            try
            {
                //C:\\Users\\datsunnn\\Documents\\Visual Studio 2010\\Projects\\UMLMetrics\\UCD\\UCDFixes.xlsx
                wb.SaveAs(workDir + "\\UCD\\UCDFixes.xlsx");
                excelApp.Quit();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine("Неудалось сохранить excel файл, скорее всего он уже открыт.\r\n Закройте его и попробуйте снова.");
            }
            return 0;
        }
        static int ExcelADFixesOutput(List<string> files, string path, string workDir)
        {
            //C:\\Users\\datsunnn\\Documents\\Visual Studio 2010\\Projects\\UMLMetrics\\AD\\ADFixes.xlsx
            if (File.Exists(workDir + "\\AD\\ADFixes.xlsx"))
            {
                try
                {
                    //C:\\Users\\datsunnn\\Documents\\Visual Studio 2010\\Projects\\UMLMetrics\\AD\\ADFixes.xlsx
                    File.Delete(workDir + "\\AD\\ADFixes.xlsx");
                }
                catch
                {
                    Console.WriteLine("\nНе удается создать Excel файл. Пожалуйста закройте открытый Excel файл с исправлениями в AD и попробуйте снова.");
                    Console.ReadLine();
                    return -1;
                }
            }
            var excelApp = new Excel.Application();
            var wb = excelApp.Workbooks.Add();

            while (wb.Sheets.Count > 1)
            {
                Excel.Worksheet ws = (Excel.Worksheet)wb.Sheets[1];
                ws.Delete();
            }

            Excel.Worksheet wsTotal = excelApp.Worksheets.Add();
            wsTotal.Name = "Общие сведения исправлений";

            Excel.Worksheet wsDet = excelApp.Worksheets.Add();
            wsDet.Name = "Подробности исправлений";

            Excel.Worksheet wst = (Excel.Worksheet)wb.Sheets[3];
            wst.Delete();

            wsTotal.Cells[1, 1] = "Название файла";
            wsTotal.Cells[1, 2] = "Количество удаленных дубликатов участников (id)";
            wsTotal.Cells[1, 3] = "Количество удаленных дубликатов участников (names)";
            wsTotal.Cells[1, 4] = "Количество удаленных дубликатов активностей (id)";
            wsTotal.Cells[1, 5] = "Количество удаленных дубликатов активностей (names)";
            wsTotal.Cells[1, 6] = "Количество удаленных дубликатов условных переходов";
            wsTotal.Cells[1, 7] = "Количество удаленных дубликатов разветвителей";
            wsTotal.Cells[1, 8] = "Количество удаленных дубликатов синхронизаторов";
            wsTotal.Cells[1, 9] = "Количество удаленных дубликатов переходов";
            wsTotal.Cells[1, 10] = "Количество удаленных Незавершенных переходов";

            wsTotal.Cells[1, 11] = "Итог";

            wsDet.Cells[1, 1] = "Название файла";
            wsDet.Cells[1, 2] = "Вид исправления";
            wsDet.Cells[1, 3] = "Тип елемента";
            wsDet.Cells[1, 4] = "Название/id елемента";
            wsDet.Cells[1, 5] = "id источника";
            wsDet.Cells[1, 6] = "id цели";

            int totalFixes = 0;
            int totalIndex = 2;
            int detIndex = 2;
            foreach (var file in files)
            {
                if (Path.GetExtension(path + '\\' + file) == ".xmi" || Path.GetExtension(path + '\\' + file) == ".XMI")
                {
                    try
                    {
                        var forCheck = new XmlDocument();
                        forCheck.Load(file);
                    }
                    catch
                    {
                        continue;
                    }
                    ADModel curModel = new ADModel(file);
                    if (TypeDefiner.DefineDiagramType(file).Contains("AD"))
                    {
                        var adNodesList = new ADNodesList();
                        XmiParser parser = new XmiParser(adNodesList);
                        bool hasJoinOrFork = false;
                        parser.Parse(curModel, ref hasJoinOrFork);

                        ADFileFixer aff = new ADFileFixer(adNodesList);

                        wsTotal.Cells[totalIndex, 1] = file.Split('\\')[file.Split('\\').Count() - 1];

                        int fileFixes = 0;

                        var delSw = aff.RemoveDublicSwimlanes(null);
                        List<BaseNode> delElems = new List<BaseNode>();
                        foreach (var el in delSw)
                            delElems.Add(el);
                        fileFixes += delSw.Count;
                        wsTotal.Cells[totalIndex, 2] = fileFixes;
                        excelElOut(wsDet, delElems, ref detIndex, curModel.FilePath, "Удаление дубликатов участников (id)");

                        var delNaSwim = aff.RemoveDublicNamesSwimlanes(null);
                        delElems = new List<BaseNode>();
                        foreach (var el in delNaSwim)
                            delElems.Add(el);
                        fileFixes += delNaSwim.Count;
                        wsTotal.Cells[totalIndex, 3] = delNaSwim.Count;
                        excelElOut(wsDet, delElems, ref detIndex, curModel.FilePath, "Удаление дубликатов участников (names)");

                        var delAct = aff.RemoveDublicActivities(null);
                        delElems = new List<BaseNode>();
                        foreach (var el in delAct)
                            delElems.Add(el);
                        fileFixes += delAct.Count;
                        wsTotal.Cells[totalIndex, 4] = delAct.Count;
                        excelElOut(wsDet, delElems, ref detIndex, curModel.FilePath, "Удаление дубликатов активностей (id)");

                        var delNaActiv = aff.RemoveDublicNamesSwimlanes(null);
                        delElems = new List<BaseNode>();
                        foreach (var el in delNaActiv)
                            delElems.Add(el);
                        fileFixes += delNaActiv.Count;
                        wsTotal.Cells[totalIndex, 5] = delNaActiv.Count;
                        excelElOut(wsDet, delElems, ref detIndex, curModel.FilePath, "Удаление дубликатов участников (names)");

                        var delDN = aff.RemoveDublicDesNodes(null);
                        delElems = new List<BaseNode>();
                        foreach (var el in delDN)
                            delElems.Add(el);
                        fileFixes += delDN.Count;
                        wsTotal.Cells[totalIndex, 6] = delDN.Count;
                        excelElOut(wsDet, delElems, ref detIndex, curModel.FilePath, "Удаление дубликатов условных переходов");

                        var delForks = aff.RemoveDublicForks(null);
                        delElems = new List<BaseNode>();
                        foreach (var el in delForks)
                            delElems.Add(el);
                        fileFixes += delForks.Count;
                        wsTotal.Cells[totalIndex, 7] = delForks.Count;
                        excelElOut(wsDet, delElems, ref detIndex, curModel.FilePath, "Удаление дубликатов разветвителей");

                        var delJoins = aff.RemoveDublicJoins(null);
                        delElems = new List<BaseNode>();
                        foreach (var el in delJoins)
                            delElems.Add(el);
                        fileFixes += delJoins.Count;
                        wsTotal.Cells[totalIndex, 8] = delJoins.Count;
                        excelElOut(wsDet, delElems, ref detIndex, curModel.FilePath, "Удаление дубликатов синхронизаторов");

                        var delFlows = aff.RemoveDublicFlows(null);
                        delElems = new List<BaseNode>();
                        foreach (var el in delFlows)
                            delElems.Add(el);
                        fileFixes += delFlows.Count;
                        wsTotal.Cells[totalIndex, 9] = delFlows.Count;
                        excelElOut(wsDet, delElems, ref detIndex, curModel.FilePath, "Удаление дубликатов переходов");

                        var delInFlows = aff.RemoveIncompleteFlows(null);
                        delElems = new List<BaseNode>();
                        foreach (var el in delInFlows)
                            delElems.Add(el);
                        fileFixes += delInFlows.Count;
                        wsTotal.Cells[totalIndex, 10] = delInFlows.Count;
                        excelElOut(wsDet, delElems, ref detIndex, curModel.FilePath, "Удаление незавершенных переходов");

                        wsTotal.Cells[totalIndex, 11] = fileFixes;
                        totalFixes += fileFixes;
                        totalIndex++;
                    }
                }
            }
            wsTotal.Cells[(totalIndex + 1), 1] = "Общее кол-во";
            wsTotal.Cells[(totalIndex + 1), 2] = totalFixes;
            try
            {
                //C:\\Users\\datsunnn\\Documents\\Visual Studio 2010\\Projects\\UMLMetrics\\AD\\ADFixes.xlsx
                wb.SaveAs(workDir + "\\AD\\ADFixes.xlsx");
                excelApp.Quit();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine("Неудалось сохранить excel файл, скорее всего он уже открыт.\r\n Закройте его и попробуйте снова.");
            }
            return 0;
        }
        static void createExcelMetrics(Excel.Application excelApp, Excel.Workbook wb, ref Excel.Worksheet ws, bool type)
        {
            while (wb.Sheets.Count > 1)
            {
                Excel.Worksheet wst = (Excel.Worksheet)wb.Sheets[1];
                wst.Delete();
            }

            ws = excelApp.Worksheets.Add();
            ws.Name = "Значения метрик";
            wb.Worksheets.Add(ws);

            Excel.Worksheet wste = (Excel.Worksheet)wb.Sheets[1];
            wste.Delete();
            wste = (Excel.Worksheet)wb.Sheets[2];
            wste.Delete();

            for (int i = 0; i < (type ? UCDMetricCalculator.metricNames.Count : ADMetricCalculator.metricsNames.Count); i++)
                ws.Cells[1, i + 2] = (type ? UCDMetricCalculator.metricNames[i] : ADMetricCalculator.metricsNames[i]);
            if (type)
            {
                ws.Cells[1, 2].Font.Bold = true;
                ws.Cells[1, 3].Font.Bold = true;
                ws.Cells[1, 4].Font.Bold = true;
                ws.Cells[1, 5].Font.Bold = true;
                ws.Cells[1, 8].Font.Bold = true;
                ws.Cells[1, 9].Font.Bold = true;
                ws.Cells[1, 10].Font.Bold = true;
            }

        }
        static void Main(string[] args)
        {
            string exePath = Assembly.GetExecutingAssembly().Location;
            string configPath = exePath + "\\..\\config.txt";
            if (!File.Exists(configPath))
            {
                Console.WriteLine(configPath + "\n");
                Console.WriteLine("Файл config.txt не был найдет. Добавьте его в папку с exe-файлом и попробуйте снова.\r\n");
                Console.ReadLine();
                return;
            }

            StreamReader swdir = new StreamReader(configPath);
            string workDir = swdir.ReadLine();
            swdir.Close();

            Console.WriteLine("Введите путь до папки с xmi файлами. Если эта папка находится в рабочей директории, просто введите название папки с xmi-файлами.\r\n");

            string path = Console.ReadLine();
            if (path.Split('\\').Count() == 1)
                path = workDir + "\\" + path;

            if (!Directory.Exists(path))
            {
                Console.WriteLine(path);
                Console.WriteLine("Указанная директория не существует!");
                Console.ReadLine();
                return;
            }

            UCDModel curUCDModel = null;
            ADModel curADModel = null;

            //UCD\\UCDMetrics.txt
            StreamWriter UCDsw = new StreamWriter(workDir + "\\UCD\\UCDMetrics.txt");
            //UCD\\UCDlogs.txt
            StreamWriter UCDlw = new StreamWriter(workDir + "\\UCD\\UCDlogs.txt");
            //AD\\ADMetrics.txt
            StreamWriter ADsw = new StreamWriter(workDir + "\\AD\\ADMetrics.txt");
            //AD\\ADlogs.txt
            StreamWriter ADlw = new StreamWriter(workDir + "\\AD\\ADlogs.txt");

            List<string> files = Directory.GetFiles(path).ToList();
            int totalUCDCount = 0;
            int totalADCount = 0;

            if (ExcelUCDFixesOutput(files, path, workDir) == -1 || ExcelADFixesOutput(files, path, workDir) == -1)
                return;

            //UCD\\UCDMetrics.xlsx
            if (File.Exists(workDir + "\\UCD\\UCDMetrics.xlsx"))
            {
                try
                {
                    //UCD\\UCDMetrics.xlsx
                    File.Delete(workDir + "\\UCD\\UCDMetrics.xlsx");
                }
                catch
                {
                    Console.WriteLine("\nНе удается создать Excel файл. Пожалуйста закройте открытый Excel файл с UCD метрика и попробуйте снова.");
                    Console.ReadLine();
                    return;
                }
            }
            //AD\\ADMetrics.xlsx
            if (File.Exists(workDir + "\\AD\\ADMetrics.xlsx"))
            {
                try
                {
                    //AD\\ADMetrics.xlsx
                    File.Delete(workDir + "\\AD\\ADMetrics.xlsx");
                }
                catch
                {
                    Console.WriteLine("\nНе удается создать Excel файл. Пожалуйста закройте открытый Excel файл с AD метрика и попробуйте снова.");
                    Console.ReadLine();
                    return;
                }
            }

            var UCDExcelApp = new Excel.Application();
            var UCDwb = UCDExcelApp.Workbooks.Add();
            Excel.Worksheet UCDWsMetrics = null;

            createExcelMetrics(UCDExcelApp, UCDwb, ref UCDWsMetrics, true);


            var ADExcelApp = new Excel.Application();
            var ADwb = ADExcelApp.Workbooks.Add();
            Excel.Worksheet ADWsMetrics = null;

            createExcelMetrics(ADExcelApp, ADwb, ref ADWsMetrics, false);

            int curExcelUCDIndex = 2;
            int curExcelADIndex = 2;
            foreach (var file in files)
            {
                if (Path.GetExtension(path + '\\' + file) == ".xmi" || Path.GetExtension(path + '\\' + file) == ".XMI")
                {
                    try
                    {
                        var forCheck = new XmlDocument();
                        forCheck.Load(file);
                    }
                    catch(Exception e)
                    {
                        UCDsw.WriteLine("Выбранный файл: " + curUCDModel.FilePath);
                        UCDsw.WriteLine("Программа не может вычислить метрики.\nСкорее всего ошибка в xmi файле.\nПроверьте xmi-файл после чего попробуйте снова.");
                        UCDsw.WriteLine(e.Message + '\n');
                        continue;
                    }
                    List<string> modelTypes = TypeDefiner.DefineDiagramType(file);
                    foreach (var modelType in modelTypes)
                    {

                        if (modelType == "UCD")
                            curUCDModel = new UCDModel(file);
                        else
                            curADModel = new ADModel(file);

                        if (modelType == "Неопределено")
                        {
                            continue;
                        }

                        switch (modelType)
                        {
                            case "UCD":
                                UCDlw.WriteLine("ВЫБРАННЫЙ ФАЙЛ " + curUCDModel.FilePath + ":");
                                UCDFileFixer UCDfileFixer = new UCDFileFixer(curUCDModel.Conns, curUCDModel.Elems);
                                totalUCDCount = UCDfileFixer.Fix(UCDlw);
                                break;
                            case "AD":
                                ADlw.WriteLine("ВЫБРАННЫЙ ФАЙЛ " + curADModel.FilePath + ":");
                                break;
                            default:
                                break;
                        }


                        switch (modelType)
                        {
                            case "UCD":
                                UCDMetricCalculator mc = new UCDMetricCalculator(curUCDModel);
                                try
                                {
                                    mc.Calculate();
                                }
                                catch (Exception e)
                                {
                                    UCDsw.WriteLine("Выбранный файл: " + curUCDModel.FilePath);
                                    UCDsw.WriteLine("Программа не может вычислить метрики.\nСкорее всего ошибка в xmi файле.\nПроверьте xmi-файл после чего попробуйте снова.");
                                    UCDsw.WriteLine(e.Message + '\n');
                                    break;
                                }
                                FileUCDOutput(UCDsw, mc, file);
                                ExcelMetricsUCDOutput(UCDWsMetrics, mc, curExcelUCDIndex, file.Split('\\')[file.Split('\\').Count() - 1]);
                                curExcelUCDIndex++;
                                break;
                            case "AD":
                                ADMetricCalculator am = new ADMetricCalculator(curADModel, ADlw);
                                try
                                {
                                    am.Calculate();
                                }
                                catch (Exception e)
                                {
                                    ADsw.WriteLine("Выбранный файл: " + curADModel.FilePath);
                                    ADsw.WriteLine("Программа не может вычислить метрики.\nСкорее всего ошибка в xmi файле.\nПроверьте xmi-файл после чего попробуйте снова.");
                                    ADsw.WriteLine(e.Message + '\n');
                                    break;
                                }
                                FileADOutput(ADsw, am, file);
                                ExcelMetricsADOutput(ADWsMetrics, am, curExcelADIndex, file.Split('\\')[file.Split('\\').Count() - 1]);
                                curExcelADIndex++;
                                break;
                            default:
                                break;
                        }
                    }
                }
            }

            UCDlw.WriteLine("ИТОГ ПО ВСЕМ ФАЙЛАМ: " + totalUCDCount);
            ADlw.WriteLine("ИТОГ ПО ВСЕМ ФАЙЛАМ: " + totalADCount);

            //UCD\\UCDMetrics.xlsx
            UCDwb.SaveAs(workDir + "\\UCD\\UCDMetrics.xlsx");
            UCDExcelApp.Quit();
            //AD\\ADMetrics.xlsx
            ADwb.SaveAs(workDir + "\\AD\\ADMetrics.xlsx");
            ADExcelApp.Quit();

            UCDsw.Close();
            UCDlw.Close();
            ADsw.Close();
            ADlw.Close();
            Console.WriteLine("Программа вычислила метрики!\r\n");
            Console.WriteLine("Метрики и исправления в xmi файлах находятся по пути: \r\nС:" + workDir);
            Console.ReadLine();
        }
    }
}