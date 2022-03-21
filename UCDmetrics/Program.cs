using System.IO;
using System.Collections.Generic;
class Actor
{
    public string Name { get; set; }
    public string Id { get; set; }
    public Actor(string name, string id)
    {
        Name = name;   
        Id = id;
    }
}
class UseCase
{
    public string Name { get; set; }
    public string Id { get; set; }
    public UseCase(string name, string id)
    {
        Name = name;
        Id = id;
    }
}
class Connection
{
    public string Type { get; set; }
    public string IdFrom { get; set; }
    public string IdTo { get; set; }
    public Connection(string type, string from, string to)
    {
        Type = type;
        IdFrom = from;
        IdTo = to;
    }
}
class UCDModel
{
    public string FilePath;
    public List<Actor> Actors { get; set; }
    public List<UseCase> UseCases { get; set; }
    public List<Connection> Conns { get; set; }
    public UCDModel(string filePath)
    {
        FilePath = filePath;
        Actors = new List<Actor>();
        UseCases = new List<UseCase>();
        Conns = new List<Connection>();

        XMItoCSharp(FilePath);
    }
    private void XMItoCSharp(string path)
    {
        using StreamReader file = new StreamReader(path);
        string fullText = file.ReadToEnd();
        foreach(string row in fullText.Split('\n'))
        {
            if(row.Contains("<packagedElement"))
            {
                string[] attrStr = row.Trim().Split(' ');
                switch(attrStr[1].Split('"')[1])
                {
                    case "uml:Actor":
                        string actorId = attrStr[2].Split('"')[1];
                        string actorName = attrStr[3].Split('"')[1];
                        Actor newActor = new Actor(actorName, actorId);
                        Actors.Add(newActor);
                        break;
                    case "uml:UseCase":
                        string useCaseId = attrStr[2].Split('"')[1];
                        string useCaseName = attrStr[3].Split('"')[1];
                        UseCase newUseCase = new UseCase(useCaseName, useCaseId);
                        UseCases.Add(newUseCase);
                        break; ;
                    default:
                        break;
                }
                continue;
            }
            
            if(row.Contains("<ownedEnd"))
            {
                if (Conns.Count == 0 || Conns[Conns.Count - 1].IdTo != "")
                {
                    string idFrom = row.Trim().Split(' ')[3].Split('"')[1];
                    Conns.Add(new Connection("Association", idFrom, ""));
                } 
                else
                {
                    var a = row.Trim().Split(' ');
                    var b = a[3].Split('"');
                    string idTo = row.Trim().Split(' ')[3].Split('"')[1];
                    Conns[Conns.Count - 1].IdTo = idTo;
                }
                continue;
            }

            if(row.Contains("<extend"))
            {
                string idFrom = row.Trim().Split(' ')[4].Split('"')[1];
                string idTo = row.Trim().Split(' ')[2].Split('"')[1];
                Connection newConn = new Connection("Extend", idFrom, idTo);
                Conns.Add(newConn);
                continue;
            }

            if (row.Contains("<include"))
            {
                string idFrom = row.Trim().Split(' ')[3].Split('"')[1];
                string idTo = row.Trim().Split(' ')[2].Split('"')[1];
                Connection newConn = new Connection("Include", idFrom, idTo);
                Conns.Add(newConn);
            }
        }
    }
}
class MetricCalculator
{
    public UCDModel model;
    public int nouc;
    public int noa;
    public double nouca;
    public int ucSecond;
    public double ucThird;
    public double ucFourth;
    public int cUcd;
    public int CTE;
    public int CIE;
    public MetricCalculator(UCDModel ucdModel)
    {
        model = ucdModel;
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
       return model.UseCases.Count;
    }
    public int CalcNoa()
    {
        return model.Actors.Count;
    }
    public double CalcNouca()
    {
        return (double) noa / nouc;
    }
    public void CalcUcNth()
    {
        int n = model.UseCases.Count;
        int m = model.Actors.Count;
        int[,] c = new int[n, m];

        for (int i = 0; i < n; i++)
            for (int j = 0; i < m; i++)
                c[i, j] = 0;

        foreach(var conn in model.Conns.Where(c => c.Type == "Association"))
        {
            c[model.UseCases.FindIndex(uc => uc.Id == conn.IdTo), model.Actors.FindIndex(a => a.Id == conn.IdFrom)] = 1;
        }

        ucSecond = 0;
        for (int i = 0; i < n; i++)
            for (int j = 0; j < m; j++)
                ucSecond += c[i, j];

        int[,] d = new int[n, m];
        for(int i = 0; i < n; i++)
            for(int j = 0; j < m; j++)
            {
                if(c[i, j] == 0)
                { 
                    d[i, j] = 0;
                    continue;
                }

                int possIndex = getExtededUcIndex(i);
                if (possIndex == -1)
                    d[i, j] = c[i, j];
                else
                    d[i, j] = c[i, j] - c[possIndex, j];
            }

        int[,] e = new int[n, m];
        for(int i = 0; i < n; i++)
            for(int j = 0; j < m; j++)
            {
                if(d[i, j] == 0)
                {
                    e[i, j] = 0;
                    continue;
                }

                List<int> tmp = getListOfIncluded(i);
                bool redundancy = false;
                foreach(var index in tmp)
                {
                    if(d[index, j] == 1)
                    {
                        redundancy = true;
                        break;
                    }
                }

                if(redundancy)
                    e[i, j] = 0;
                else
                    e[i, j] = 1;

            }

        ucThird = 0;
        for(int i = 0; i < n; i++)
        {
            int tempSumm = 0;
            for(int j = 0; j < m; j++)
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
        foreach(var conn in model.Conns.Where(c => c.Type == "Association"))
        {
            if(!main.Contains(conn.IdTo))
            {
                main.Add(conn.IdTo);
            }
        }

        int n = model.UseCases.Count;
        int m = model.Actors.Count + main.Count;
        int[,] mtrig = new int[n, m];

        for (int i = 0; i < n; i++)
            for (int j = 0; j < m; j++)
                mtrig[i, j] = 0;
        
        foreach(var conn in model.Conns)
        {
            if(conn.Type == "Association")
            {
                int actorInd = model.Actors.FindIndex(a => a.Id == conn.IdFrom);
                int ucInd = model.UseCases.FindIndex(uc => uc.Id == conn.IdTo);
                mtrig[ucInd, actorInd] = 3;
            }
            else if(conn.Type == "Include")
            {
                int includingInd = main.FindIndex(id => id == conn.IdFrom) + model.Actors.Count;
                int includedInd = model.UseCases.FindIndex(uc => uc.Id == conn.IdTo);
                mtrig[includedInd, includingInd] = 2;
            }
            else if(conn.Type == "Extend")
            {
                int extendedInd = main.FindIndex(id => id == conn.IdTo) + model.Actors.Count;
                int extendingInd = model.UseCases.FindIndex(uc => uc.Id == conn.IdFrom);
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
        string idFrom = model.UseCases[i].Id;
        
        foreach(var conn in model.Conns.Where(c => c.Type == "Include"))
        {
            if (conn.IdFrom == idFrom)
                result.Add(model.UseCases.FindIndex(a => a.Id == conn.IdTo));
        }

        return result;
    }
    public int getExtededUcIndex(int i)
    {
        string idFrom = model.UseCases[i].Id;
        Connection? tmp = model.Conns.FirstOrDefault(c => c.Type == "Extend" && c.IdFrom == idFrom);
        if (tmp == null)
            return -1;
        else
            return model.UseCases.FindIndex(uc => uc.Id == tmp.IdTo);
    }
}
class Program
{
    static void ConsoleOutput(MetricCalculator mc)
    {
        Console.WriteLine("Выбранный файл: " + mc.model.FilePath);
        Console.WriteLine("\nВычисленные метрики:\nNOUC: " + mc.nouc + "\nNOA: " + mc.noa + "\nNOUCA: " + mc.nouca
            + "\nUC2: " + mc.ucSecond + "\nUC3: " + mc.ucThird + "\nUC4: " + mc.ucFourth + "\nCTE: " + mc.CTE
            + "\nCIE: " + mc.CIE + "\nCucd: " + mc.cUcd);
    }
    static void Main(string[] args)
    {

        //C:\Users\efalk\Desktop\Курсовая\Library.xmi
        Console.WriteLine("Введите путь к файлу");
        string? path = Console.ReadLine();
        UCDModel curModel = new UCDModel(path);

        foreach(var act in curModel.Actors)
            Console.WriteLine(act.Name + ":  " + act.Id + "\n");

        foreach (var uc in curModel.UseCases)
            Console.WriteLine(uc.Name + ":  " + uc.Id + "\n");

        foreach (var conn in curModel.Conns)
            Console.WriteLine(conn.Type + "  :  " + conn.IdFrom + "  :  " + conn.IdTo + "\n");

        MetricCalculator mc = new MetricCalculator(curModel);
        mc.Calculate();
        ConsoleOutput(mc);
    }
}