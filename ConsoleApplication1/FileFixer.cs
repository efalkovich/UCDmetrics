using System.IO;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Xml;
using System;
using System.Linq;
using System.Xml.Linq;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApplication1
{
    interface IFileFixer
    {
        int Fix(StreamWriter lw);
    }
    class UCDFileFixer : IFileFixer
    {
        public List<Connection> conns;
        public List<Element> elems;

        public UCDFileFixer(List<Connection> conns, List<Element> elems)
        {
            this.conns = conns;
            this.elems = elems;
        }

        public int Fix(StreamWriter lw)
        {
            int summCount = RemoveDuplicatesElems(lw).Count;
            summCount += RemoveDuplicatesConns(lw).Count;

            lw.WriteLine("\r\n\tОБЩЕЕ КОЛИЧЕСТВО УДАЛЕННЫХ ДУБЛИКАТОВ: " + summCount);

            summCount += RemoveMissusedAssociations(lw).Count;
            summCount += RemoveLoopedConns(lw).Count;
            summCount += RemoveIsolatedActors(lw).Count;

            var temp = RemoveIsolatedUcs(lw);
            summCount += temp.Item1.Count + temp.Item2.Count;

            summCount += RemoveIncompleteConns(lw).Count;

            lw.WriteLine("\r\nИТОГ: " + summCount + "\r\n\r\n");

            return summCount;
        }
        public List<Element> RemoveDuplicatesElems(StreamWriter lw)
        {
            List<Element> delEl = new List<Element>();
            if (lw != null)
                lw.WriteLine("\r\nУдаленные дубликаты:");
            for (int i = 0; i < elems.Count - 1; i++)
                for (int j = i + 1; j < elems.Count; j++)
                {
                    if (elems[i].Id == elems[j].Id)
                    {
                        if (lw != null)
                            lw.WriteLine("\r\n\tТип елемента: " + elems[j].Type + "\r\n\tНазвание елемента: " + elems[j].Name);
                        delEl.Add(elems[i]);
                        elems.Remove(elems[j]);
                    }
                }
            return delEl;
        }
        public List<Connection> RemoveDuplicatesConns(StreamWriter lw)
        {
            List<Connection> delCon = new List<Connection>();
            for (int i = 0; i < conns.Count - 1; i++)
                for (int j = i + 1; j < conns.Count; j++)
                {
                    if (conns[i].IdFrom == conns[j].IdFrom && conns[i].IdTo == conns[j].IdTo)
                    {
                        Element elFrom = elems.Find(el => el.Id == conns[j].IdFrom);
                        Element elTo = elems.Find(el => el.Id == conns[j].IdTo);
                        if (lw != null)
                            lw.WriteLine("\r\n\tТип звязи: " + conns[j].Type + "\r\n\tКонцы связи: " + (elFrom != null ? elFrom.Name : "?") + " - " + (elTo != null ? elTo.Name : "?"));
                        delCon.Add(conns[j]);
                        conns.Remove(conns[j]);
                    }
                }
            return delCon;
        }
        public List<Connection> RemoveIncompleteConns(StreamWriter lw)
        {
            List<Connection> delCon = new List<Connection>();
            if (lw != null)
                lw.WriteLine("\r\nУдаленные незавершенные связи:");
            for (int i = 0; i < conns.Count; i++)
            {
                if (!(elems.Select(x => x.Id).Contains(conns[i].IdFrom) && elems.Select(x => x.Id).Contains(conns[i].IdTo)))
                {
                    Element elFrom = elems.Find(el => el.Id == conns[i].IdFrom);
                    Element elTo = elems.Find(el => el.Id == conns[i].IdTo);
                    if (lw != null)
                        lw.WriteLine("\r\n\tТип звязи: " + conns[i].Type + "\r\n\tКонцы связи: " + (elFrom != null ? elFrom.Name : "?") + " - " + (elTo != null ? elTo.Name : "?"));
                    delCon.Add(conns[i]);
                    conns.Remove(conns[i]);
                    i--;
                }
            }
            if (lw != null)
                lw.WriteLine("\r\n\tОБЩЕЕ КОЛИЧЕСТВО УДАЛЕННЫХ НЕЗАВЕРШЕННЫХ СВЯЗЕЙ: " + delCon.Count);
            return delCon;
        }
        public Tuple<List<Element>, List<Connection>> RemoveIsolatedUcs(StreamWriter lw)
        {
            List<Element> delEl = new List<Element>();
            List<Connection> delCon = new List<Connection>();
            if (lw != null)
                lw.WriteLine("\r\nУдаленные изолированные прецеденты:");
            List<Element> actors = elems.Where(e => e.Type == "uml:Actor").ToList();
            List<Element> ucs = elems.Where(e => e.Type == "uml:UseCase").ToList();
            for (int i = 0; i < ucs.Count; i++)
            {

                List<Element> ucForDel = new List<Element>();
                if (!FindPathToActor(actors, ucs, ucs[i], null, ucForDel))
                {
                    if (!elems.Exists(e => e.Id == ucs[i].Id))
                        continue;
                    foreach (var del in ucForDel)
                    {
                        List<Connection> connsForDel = conns.Where(c => c.IdFrom == del.Id || c.IdTo == del.Id).ToList();
                        foreach (var conn in connsForDel)
                        {
                            Element elFrom = ucs.Find(el => el.Id == conn.IdFrom);
                            Element elTo = ucs.Find(el => el.Id == conn.IdTo);
                            if (lw != null)
                                lw.WriteLine("\r\n\tТип звязи: " + conn.Type + "\r\n\tКонцы связи: " + (elFrom != null ? elFrom.Name : "?") + " - " + (elTo != null ? elTo.Name : "?"));
                            delCon.Add(conn);
                            conns.Remove(conn);
                        }
                        if (lw != null)
                            lw.WriteLine("\r\n\tТип елемента: " + del.Type + "\r\n\tНазвание елемента: " + del.Name);
                        delEl.Add(del);
                        elems.Remove(del);
                    }
                }
            }
            if (lw != null)
                lw.WriteLine("\r\n\tОБЩЕЕ КОЛИЧЕСТВО УДАЛЕННЫХ ИЗОЛИРОВАННЫХ ПРЕЦЕДЕНТОВ И СВЯЗЕЙ: " + (delCon.Count + delEl.Count));
            return Tuple.Create(delEl, delCon);
        }
        public List<Element> RemoveIsolatedActors(StreamWriter lw)
        {
            List<Element> delEl = new List<Element>();
            if (lw != null)
                lw.WriteLine("\r\nУдаленные изолированные акторы: ");
            List<Element> actors = elems.Where(e => e.Type == "uml:Actor").ToList();
            for (int i = 0; i < actors.Count; i++)
                if (!conns.Exists(c => c.IdFrom == actors[i].Id || c.IdTo == actors[i].Id))
                {
                    if (lw != null)
                        lw.WriteLine("\r\n\tТип елемента: " + actors[i].Type + "\r\n\tНазвание елемента: " + actors[i].Name);
                    delEl.Add(actors[i]);
                    elems.Remove(actors[i]);
                }
            if (lw != null)
                lw.WriteLine("\r\n\tОБЩЕЕ КОЛИЧЕСТВО УДАЛЕННЫХ ИЗОЛИРОВАННЫХ АКТОРОВ: " + delEl.Count);
            return delEl;
        }
        public List<Connection> RemoveMissusedAssociations(StreamWriter lw)
        {
            List<Connection> delCon = new List<Connection>();
            if (lw != null)
                lw.WriteLine("\r\nУдаленные неверно использованные связи типа ассоциации: ");
            List<Element> acts = elems.Where(el => el.Type == "uml:Actor").ToList();
            List<Element> ucs = elems.Where(el => el.Type == "uml:UseCase").ToList();
            for (int i = 0; i < conns.Count; i++)
            {
                if (acts.Select(a => a.Id).Contains(conns[i].IdFrom) && acts.Select(a => a.Id).Contains(conns[i].IdTo))
                {
                    Element act = elems.Find(el => el.Id == conns[i].IdFrom || el.Id == conns[i].IdTo);
                    if (lw != null)
                        lw.WriteLine("\r\n\tТип звязи: " + conns[i].Type + "\r\n\tКонцы связи: " + (act != null ? act.Name : "?") + " - " + (act != null ? act.Name : "?" + "   (Актор - Актор)"));
                    delCon.Add(conns[i]);
                    conns.Remove(conns[i]);
                    i--;
                }
                else if (ucs.Select(u => u.Id).Contains(conns[i].IdFrom) && ucs.Select(u => u.Id).Contains(conns[i].IdTo) && conns[i].Type == "uml:Association")
                {
                    Element act = elems.Find(el => el.Id == conns[i].IdFrom || el.Id == conns[i].IdTo);
                    if (lw != null)
                        lw.WriteLine("\r\n\tТип звязи: " + conns[i].Type + "\r\n\tКонцы связи: " + (act != null ? act.Name : "?") + " - " + (act != null ? act.Name : "?" + "   (Прецедент - Прецедент)"));
                    delCon.Add(conns[i]);
                    conns.Remove(conns[i]);
                    i--;
                }
            }

            if (lw != null)
                lw.WriteLine("\r\n\tОБЩЕЕ КОЛИЧЕСТВО УДАЛЕННЫХ НЕВЕРНО ИСПОЛЬЗОВАННЫХ СВЯЗЕЙ ТИПА АССОЦИАЦИИ: " + delCon.Count);
            return delCon;
        }
        public List<Connection> RemoveLoopedConns(StreamWriter lw)
        {
            List<Connection> delCon = new List<Connection>();
            if (lw != null)
                lw.WriteLine("\r\nУдаленные циклические связи: ");
            for (int i = 0; i < conns.Count; i++)
            {
                if (conns[i].IdFrom == conns[i].IdTo)
                {
                    Element loopEl = elems.Find(el => el.Id == conns[i].IdFrom || el.Id == conns[i].IdTo);
                    delCon.Add(conns[i]);
                    conns.Remove(conns[i]);
                    if (lw != null)
                        lw.WriteLine("\r\n\tТип звязи: " + conns[i].Type + "\r\n\tКонцы связи: " + (loopEl != null ? loopEl.Name : "?") + " - " + (loopEl != null ? loopEl.Name : "?"));
                    i--;
                }
            }
            if (lw != null)
                lw.WriteLine("\r\n\tОБЩЕЕ КОЛИЧЕСТВО УДАЛЕННЫХ ЦИКЛИЧЕСКИХ СВЯЗЕЙ: " + delCon.Count);
            return delCon;
        }
        private bool FindPathToActor(List<Element> actors, List<Element> ucs, Element curUc, Connection usedConn, List<Element> UcForDeleting)
        {
            List<Connection> cons = conns.Where(c => c.IdFrom == curUc.Id || c.IdTo == curUc.Id).ToList();
            UcForDeleting.Add(curUc);
            foreach (var conn in cons)
            {
                if (conn.Equals(usedConn))
                    continue;
                if (actors.Exists(a => a.Id == conn.IdFrom || a.Id == conn.IdTo))
                    return true;
                Element nextUc = ucs.Find(u => (u.Id == conn.IdFrom || u.Id == conn.IdTo) && !UcForDeleting.Contains(u));
                if (nextUc == null)
                    continue;
                if (FindPathToActor(actors, ucs, nextUc, conn, UcForDeleting))
                    return true;
            }
            return false;
        }
    }
    class ADFileFixer : IFileFixer
    {
        private ADNodesList elems;
        public int totalFixCount = 0;
        public ADFileFixer(ADNodesList elems)
        {
            this.elems = elems;
        }
        public int Fix(StreamWriter lw)
        {

            var remSw = RemoveDublicSwimlanes(lw);
            totalFixCount += remSw.Count;

            if (lw != null)
                lw.WriteLine("\n\tОБЩЕЕ КОЛИЧЕСТВО УДАЛЕННЫХ ДУБЛИКАТОВ УЧАСТНИКОВ (по id): " + totalFixCount);

            var remAct = RemoveDublicActivities(lw);
            int delCount = remAct.Count;
            totalFixCount += delCount;

            if (lw != null)
                lw.WriteLine("\n\tОБЩЕЕ КОЛИЧЕСТВО УДАЛЕННЫХ ДУБЛИКАТОВ АКТИВНОСТЕЙ (по id): " + delCount);

            var remDN = RemoveDublicDesNodes(lw);
            delCount = remDN.Count;
            totalFixCount += delCount;

            if (lw != null)
                lw.WriteLine("\n\tОБЩЕЕ КОЛИЧЕСТВО УДАЛЕННЫХ ДУБЛИКАТОВ УСЛОВНЫХ ПЕРЕХОДОВ: " + delCount);

            var remFork = RemoveDublicForks(lw);
            delCount = remFork.Count;
            totalFixCount += delCount;

            if (lw != null)
                lw.WriteLine("\n\tОБЩЕЕ КОЛИЧЕСТВО УДАЛЕННЫХ ДУБЛИКАТОВ РАЗВЕТВИТЕЛЕЙ: " + delCount);

            var remJoin = RemoveDublicJoins(lw);
            delCount = remJoin.Count;
            totalFixCount += delCount;

            if (lw != null)
                lw.WriteLine("\n\tОБЩЕЕ КОЛИЧЕСТВО УДАЛЕННЫХ ДУБЛИКАТОВ СИНХРОНИЗАТОРОВ: " + delCount);

            var remFlows = RemoveDublicFlows(lw);
            delCount = remFlows.Count;
            totalFixCount += delCount;

            if (lw != null)
                lw.WriteLine("\n\tОБЩЕЕ КОЛИЧЕСТВО УДАЛЕННЫХ ДУБЛИКАТОВ ПЕРЕХОДОВ: " + delCount);

            var remInFlows = RemoveIncompleteFlows(lw);
            delCount = remInFlows.Count;
            totalFixCount += delCount;

            if (lw != null)
                lw.WriteLine("\n\tОБЩЕЕ КОЛИЧЕСТВО УДАЛЕННЫХ НЕЗАВЕРШЕННЫХ ПЕРЕХОДОВ: " + delCount);

            var remDuNaSwim = RemoveDublicNamesSwimlanes(lw);
            delCount = remDuNaSwim.Count;
            totalFixCount += delCount;

            if (lw != null)
                lw.WriteLine("\n\tОБЩЕЕ КОЛИЧЕСТВО УДАЛЕННЫХ ДУБЛИКАТОВ УЧАСТНИКОВ (по имени): " + delCount);

            var remDuNaAct = RemoveDublicNamesSwimlanes(lw);
            delCount = remDuNaAct.Count;
            totalFixCount += delCount;

            if (lw != null)
                lw.WriteLine("\n\tОБЩЕЕ КОЛИЧЕСТВО УДАЛЕННЫХ ДУБЛИКАТОВ АКТИВНОСТЕЙ (по имени): " + delCount);

            if (lw != null)
                lw.WriteLine("\nИТОГ: " + totalFixCount + "\n\n");
            return totalFixCount;
        }
        public List<Swimlane> RemoveDublicSwimlanes(StreamWriter lw)
        {
            List<Swimlane> delElems = new List<Swimlane>();
            List<Swimlane> els = elems.getAllSwimlanes();
            if (lw != null)
                lw.WriteLine("\nУдаленные дубликаты участников:");
            for (int i = 0; i < els.Count - 1; i++)
                for (int j = i + 1; j < els.Count; j++)
                {
                    string temp1 = els[i].getId();
                    string temp2 = els[j].getId();
                    if (els[i].getId() == els[j].getId())
                    {
                        if (lw != null)
                            lw.WriteLine("\n\tТип елемента: " + els[j].getType() + "\n\tНазвание елемента: " + els[j].getName());
                        delElems.Add(els[i]);
                        elems.nodes.RemoveAll(n => n.value.getId() == els[j].getId());
                        els.Remove(els[j]);
                    }
                }
            return delElems;
        }
        public List<ActivityNode> RemoveDublicActivities(StreamWriter lw)
        {
            List<ActivityNode> delElems = new List<ActivityNode>();
            List<ActivityNode> els = elems.getAllActivities();
            if (lw != null)
                lw.WriteLine("\nУдаленные дубликаты активностей:");
            for (int i = 0; i < els.Count - 1; i++)
                for (int j = i + 1; j < els.Count; j++)
                {
                    if (els[i].getId() == els[j].getId())
                    {
                        if (lw != null)
                            lw.WriteLine("\n\tТип елемента: " + els[j].getType() + "\n\tНазвание елемента: " + els[j].getName());
                        delElems.Add(els[i]);
                        elems.nodes.RemoveAll(n => n.value.getId() == els[j].getId());
                        els.Remove(els[j]);
                    }
                }
            return delElems;
        }
        public List<DecisionNode> RemoveDublicDesNodes(StreamWriter lw)
        {
            List<DecisionNode> delElems = new List<DecisionNode>();
            List<DecisionNode> els = elems.getAllDecisionNodes();
            if (lw != null)
                lw.WriteLine("\nУдаленные дубликаты условных переходов:");
            for (int i = 0; i < els.Count - 1; i++)
                for (int j = i + 1; j < els.Count; j++)
                {
                    if (els[i].getId() == els[j].getId())
                    {
                        if (lw != null)
                            lw.WriteLine("\n\tТип елемента: " + els[j].getType() + "\n\tНазвание елемента: " + els[j].getQuestion());
                        delElems.Add(els[i]);
                        elems.nodes.RemoveAll(n => n.value.getId() == els[j].getId());
                        els.Remove(els[j]);
                    }
                }
            return delElems;
        }
        public List<ForkNode> RemoveDublicForks(StreamWriter lw)
        {
            List<ForkNode> delElems = new List<ForkNode>();
            List<ForkNode> els = elems.getAllForkNodes();
            if (lw != null)
                lw.WriteLine("\nУдаленные дубликаты разветвителей:");
            for (int i = 0; i < els.Count - 1; i++)
                for (int j = i + 1; j < els.Count; j++)
                {
                    if (els[i].getId() == els[j].getId())
                    {
                        if (lw != null)
                            lw.WriteLine("\n\tТип елемента: " + els[j].getType() + "\n\tId елемента: " + els[j].getId());
                        delElems.Add(els[i]);
                        elems.nodes.RemoveAll(n => n.value.getId() == els[j].getId());
                        els.Remove(els[j]);
                    }
                }
            return delElems;
        }
        public List<JoinNode> RemoveDublicJoins(StreamWriter lw)
        {
            List<JoinNode> delElems = new List<JoinNode>();
            List<JoinNode> els = elems.getAllJoinNodes();
            if (lw != null)
                lw.WriteLine("\nУдаленные дубликаты синхронизаторов:");
            for (int i = 0; i < els.Count - 1; i++)
                for (int j = i + 1; j < els.Count; j++)
                {
                    if (els[i].getId() == els[j].getId())
                    {
                        if (lw != null)
                            lw.WriteLine("\n\tТип елемента: " + els[j].getType() + "\n\tId елемента: " + els[j].getId());
                        delElems.Add(els[i]);
                        elems.nodes.RemoveAll(n => n.value.getId() == els[j].getId());
                        els.Remove(els[j]);
                    }
                }
            return delElems;
        }
        public List<ControlFlow> RemoveDublicFlows(StreamWriter lw)
        {
            List<ControlFlow> delElems = new List<ControlFlow>();
            List<ControlFlow> els = elems.getAllContrFlows();
            if (lw != null)
                lw.WriteLine("\nУдаленные дубликаты переходов:");
            for (int i = 0; i < els.Count - 1; i++)
                for (int j = i + 1; j < els.Count; j++)
                {
                    if (els[i].getSrc() == els[j].getSrc() && els[i].getTarget() == els[j].getTarget() || els[i].getId() == els[j].getId())
                    {
                        if (lw != null)
                            lw.WriteLine("\n\tТип елемента: " + els[j].getType() + "\n\tId источника: " + els[j].getSrc() + "\n\tId цели: " + els[j].getTarget());
                        delElems.Add(els[i]);
                        elems.nodes.RemoveAll(n => n.value.getId() == els[j].getId());
                        els.Remove(els[j]);
                    }
                }
            return delElems;
        }
        public List<ControlFlow> RemoveIncompleteFlows(StreamWriter lw)
        {
            List<ControlFlow> delElems = new List<ControlFlow>();
            List<ControlFlow> els = elems.getAllContrFlows();
            if (lw != null)
                lw.WriteLine("\nУдаленные незавершенные переходы:");
            for (int i = 0; i < els.Count - 1; i++)
            {
                var nonFlows = elems.nodes.Where(e => e.getValue().getType() != ElementType.FLOW);
                if (!(nonFlows.Select(n => n.value.getId()).Contains(els[i].getSrc()) && nonFlows.Select(n => n.value.getId()).Contains(els[i].getTarget())))
                {
                    if (lw != null)
                        lw.WriteLine("\n\tТип елемента: " + els[i].getType() + "\n\tId источника: " + els[i].getSrc() + "\n\tId цели: " + els[i].getTarget());
                    delElems.Add(els[i]);
                    elems.nodes.RemoveAll(n => n.value.getId() == els[i].getId());
                    els.Remove(els[i]);
                }
            }
            return delElems;
        }
        public List<Swimlane> RemoveDublicNamesSwimlanes(StreamWriter lw)
        {
            List<Swimlane> delElems = new List<Swimlane>();
            List<Swimlane> els = elems.getAllSwimlanes();
            if (lw != null)
                lw.WriteLine("\nУдаленные дубликаты участников (по имени):");
            for (int i = 0; i < els.Count - 1; i++)
                for (int j = i + 1; j < els.Count; j++)
                {
                    if (els[i].name == els[j].name && !(els[i].name != "" || els[i].getId().Count() < 5))
                    {
                        if (lw != null)
                            lw.WriteLine("\n\tТип елемента: " + els[j].getType() + "\n\tНазвание елемента: " + els[j].getName());
                        delElems.Add(els[i]);
                        elems.nodes.RemoveAll(n => n.value.getId() == els[j].getId());
                        els.Remove(els[j]);
                    }
                }
            return delElems;
        }
        public List<ActivityNode> RemoveDublicNamesActivities(StreamWriter lw)
        {
            List<ActivityNode> delElems = new List<ActivityNode>();
            List<ActivityNode> els = elems.getAllActivities();
            if (lw != null)
                lw.WriteLine("\nУдаленные дубликаты активностей:");
            for (int i = 0; i < els.Count - 1; i++)
                for (int j = i + 1; j < els.Count; j++)
                {
                    if (els[i].getName() == els[j].getName())
                    {
                        if (lw != null)
                            lw.WriteLine("\n\tТип елемента: " + els[j].getType() + "\n\tНазвание елемента: " + els[j].getName());
                        delElems.Add(els[i]);
                        elems.nodes.RemoveAll(n => n.value.getId() == els[j].getId());
                        els.Remove(els[j]);
                    }
                }
            return delElems;
        }
    }
}
