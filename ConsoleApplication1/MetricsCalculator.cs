using System.IO;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Xml;
using System;
using System.Linq;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApplication1
{
    interface IMetricsCalculator
    {
        void Calculate();
    }
    class UCDMetricCalculator : IMetricsCalculator
    {
        public static List<string> metricNames = new List<string> { "nouc=UC1", "noa", "nouca", "UC2", "UC3", "UC4", "CTE", "CIE", "Cucd" };
        private List<Element> actors;
        private List<Element> useCases;
        private List<Connection> conns;
        public int nouc;
        public int noa;
        public double nouca;
        public int ucSecond;
        public double ucThird;
        public double ucFourth;
        public int cUcd;
        public int CTE;
        public int CIE;
        public UCDMetricCalculator(UCDModel model)
        {
            conns = model.Conns;
            actors = model.Elems.Where(e => e.Type == "uml:Actor").ToList();
            useCases = model.Elems.Where(e => e.Type == "uml:UseCase").ToList();
        }
        public void Calculate()
        {
            nouc = CalcNouc();
            noa = CalcNoa();
            nouca = CalcNouca();
            CalcUcNth();
            CalcCUcd();
        }
        public int CalcNouc()
        {
            return useCases.Count;
        }
        public int CalcNoa()
        {
            return actors.Count;
        }
        public double CalcNouca()
        {
            return (double)noa / nouc;
        }
        public void CalcUcNth()
        {
            int n = useCases.Count;
            int m = actors.Count;
            int[,] c = new int[n, m];

            for (int i = 0; i < n; i++)
                for (int j = 0; j < m; j++)
                    c[i, j] = 0;

            List<Connection> assocConn = new List<Connection>();
            foreach (var conn in conns)
                if (conn.Type == "Association")
                    assocConn.Add(conn);

            foreach (var conn in assocConn)
            {
                int ind1 = useCases.FindIndex(uc => uc.Id == conn.IdTo || uc.Id == conn.IdFrom);
                int ind2 = actors.FindIndex(a => a.Id == conn.IdFrom || a.Id == conn.IdTo);
                if (ind1 != -1 && ind2 != -1)
                    c[ind1, ind2] = 1;
            }

            ucSecond = 0;
            for (int i = 0; i < n; i++)
                for (int j = 0; j < m; j++)
                    ucSecond += c[i, j];

            int[,] d = new int[n, m];
            for (int i = 0; i < n; i++)
                for (int j = 0; j < m; j++)
                {
                    if (c[i, j] == 0)
                    {
                        d[i, j] = 0;
                        continue;
                    }

                    if(i==1 && j == 3)
                        i = 1;

                    int possIndex = getExtededUcIndex(i);
                    if (possIndex == -1)
                        d[i, j] = c[i, j];
                    else
                    {
                        int temp = c[i, j] - c[possIndex, j];
                        d[i, j] = temp;
                    }
                }

            int[,] e = new int[n, m];
            for (int i = 0; i < n; i++)
                for (int j = 0; j < m; j++)
                {
                    if (d[i, j] == 0)
                    {
                        e[i, j] = 0;
                        continue;
                    }

                    List<int> tmp = getListOfIncluded(i);
                    bool redundancy = false;
                    foreach (var index in tmp)
                    {
                        if (d[index, j] == 1)
                        {
                            redundancy = true;
                            break;
                        }
                    }

                    if (redundancy)
                        e[i, j] = 0;
                    else
                        e[i, j] = 1;

                }

            ucThird = 0;
            for (int i = 0; i < n; i++)
            {
                int tempSumm = 0;
                for (int j = 0; j < m; j++)
                    tempSumm += e[i, j];
                ucThird += Math.Pow(tempSumm, 1.4);
            }

            int summE = 0;
            int summC = 0;
            for (int i = 0; i < n; i++)
                for (int j = 0; j < m; j++)
                {
                    summE += e[i, j];
                    summC += c[i, j];
                }

            ucFourth = 0.1 * nouc * nouc + ucThird + 0.1 * (summC - summE);
        }
        public void CalcCUcd()
        {
            List<string> main = new List<string>();
            foreach (var conn in conns.Where(c => c.Type == "Association"))
            {
                if (!main.Contains(conn.IdTo))
                {
                    main.Add(conn.IdTo);
                }
            }

            int n = useCases.Count;
            int m = actors.Count + main.Count;
            int[,] mtrig = new int[n, m];

            for (int i = 0; i < n; i++)
                for (int j = 0; j < m; j++)
                    mtrig[i, j] = 0;

            foreach (var conn in conns)
            {
                if (conn.Type == "Association")
                {
                    int actorInd = actors.FindIndex(a => a.Id == conn.IdFrom || a.Id == conn.IdTo);
                    int ucInd = useCases.FindIndex(uc => uc.Id == conn.IdTo || uc.Id == conn.IdFrom);
                    if (ucInd != -1 && actorInd != -1)
                        mtrig[ucInd, actorInd] = 3;
                }
                else if (conn.Type == "Include")
                {
                    int includingInd = main.FindIndex(id => id == conn.IdFrom) + actors.Count;
                    int includedInd = useCases.FindIndex(uc => uc.Id == conn.IdTo);
                    if (includedInd != -1 && includingInd != -1)
                        mtrig[includedInd, includingInd] = 2;
                }
                else if (conn.Type == "Extend")
                {
                    int extendedInd = main.FindIndex(id => id == conn.IdTo) + actors.Count;
                    int extendingInd = useCases.FindIndex(uc => uc.Id == conn.IdFrom);
                    if (extendingInd != -1 && extendedInd != -1)
                        mtrig[extendingInd, extendedInd] = 1;
                }
            }

            int[,] minit = GetTransparentMatrix(mtrig, n, m);

            CTE = 0;
            for (int i = 0; i < n; i++)
            {
                int tmp = 0;
                bool one = false;
                for (int j = 0; j < m; j++)
                {
                    if (mtrig[i, j] != 1)
                        tmp += mtrig[i, j];
                    else
                        one = true;
                }
                CTE += tmp + (one ? 1 : 0);
            }

            CIE = 0;
            for (int i = 0; i < m; i++)
            {
                int tmp = 0;
                bool one = false;
                for (int j = 0; j < n; j++)
                {
                    if (minit[i, j] != 1)
                        tmp += minit[i, j];
                    else
                        one = true;
                }
                CIE += tmp + (one ? 1 : 0);
            }

            cUcd = CTE + CIE;
        }
        public int[,] GetTransparentMatrix(int[,] mat, int n, int m)
        {
            int[,] t = new int[m, n];
            for (int i = 0; i < n; i++)
                for (int j = 0; j < m; j++)
                    t[j, i] = mat[i, j];
            return t;
        }
        public List<int> getListOfIncluded(int i)
        {
            List<int> result = new List<int>();
            string idFrom = useCases[i].Id;

            foreach (var conn in conns.Where(c => c.Type == "Include"))
            {
                if (conn.IdFrom == idFrom)
                    result.Add(useCases.FindIndex(a => a.Id == conn.IdTo));
            }

            return result;
        }
        public int getExtededUcIndex(int i)
        {
            string idTo = useCases[i].Id;
            Connection tmp = conns.FirstOrDefault(c => c.Type == "Extend" && c.IdTo == idTo);
            if (tmp == null)
                return -1;
            else
                return useCases.FindIndex(uc => uc.Id == tmp.IdFrom);
        }
    }
    class ADMetricCalculator : IMetricsCalculator
    {
        StreamWriter logWriter;
        ADFileFixer aff;
        ADNodesList adNodesList;
        public int nosw;
        public int noact;
        public int nodn;
        public int nof;
        public int noj;
        public int noe;
        public int comp;
        private int noip;
        public static List<string> metricsNames = new List<string> { "NoSw", "NoAct", "NoDN", "NoF", "NoJ", "NoE", "C" };
        public int totalFixes = 0;
        public ADMetricCalculator(ADModel model, StreamWriter lw)
        {
            logWriter = lw;
            adNodesList = model.adNodeList;
        }
        public void Calculate()
        {
            aff = new ADFileFixer(adNodesList);
            totalFixes = aff.Fix(logWriter);

            nosw = calcNosw();
            noact = calcNoact();
            nodn = calcNodn();
            nof = calcNof();
            noj = calcNoj();
            calcC();
        }
        private int calcNosw()
        {
            return adNodesList.getAllSwimlanes().Count;
        }
        private int calcNoact()
        {
            return adNodesList.getAllActivities().Count;
        }
        private int calcNodn()
        {
            return adNodesList.getAllDecisionNodes().Count;
        }
        private int calcNof()
        {
            return adNodesList.getAllForkNodes().Count;
        }
        private int calcNoj()
        {
            return adNodesList.getAllJoinNodes().Count;
        }
        private void calcC()
        {
            int n = noact + nodn + nof + noj;
            noe = adNodesList.getAllContrFlows().Count;
            noip = noe - n + 2;
            comp = noip + n;
        }
    }
}
