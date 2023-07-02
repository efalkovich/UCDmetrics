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
    class ADModel : Model
    {
        public ADNodesList adNodeList;
        public ADModel(string filePath) : base(filePath)
        {
            adNodeList = new ADNodesList();
            FilePath = filePath;
            XMItoCSharp();
        }
        public override void XMItoCSharp()
        { 
            ////ПИШИ ЗДЕСЬ/////
            bool hasJoinOrFork = false;
            XmiParser xp = new XmiParser(adNodeList);
            xp.Parse(this, ref hasJoinOrFork);
        }
    }
    public enum ElementType
    {
        UNKNOWN,
        FLOW,
        ACTIVITY,
        FORK,
        JOIN,
        DECISION,
        MERGE,
        INITIAL_NODE,
        FINAL_NODE,
        SWIMLANE
    }
    public abstract class BaseNode
    {
        protected string id;
        protected ElementType type;
        public int x = -1;
        public int y = -1;
        public int width = -1;
        public int height = -1;

        public override bool Equals(object obj)
        {
            if (this == obj) return true;
            if (obj == null) return false;
            BaseNode baseNode = (BaseNode)obj;
            return id.Equals(baseNode.id);
        }

        public override int GetHashCode()
        {
            return id.GetHashCode();
        }
        public string getId()
        {
            return id;
        }
        public abstract string getName();
        public abstract string getSrc();
        public abstract string getTarget();
        public BaseNode(string id)
        {
            this.id = id;
        }

        //region Getter-Setter
        public ElementType getType()
        {
            return type;
        }

        public void setType(ElementType type)
        {
            this.type = type;
        }
        //endregion
    }
    public class ElementTypeAdapter
    {
        public static string toString(ElementType type)
        {
            switch (type)
            {
                case ElementType.DECISION: return "Условный переход";
                case ElementType.ACTIVITY: return "Активность";
                case ElementType.MERGE: return "Узел слияния";
                case ElementType.JOIN: return "Синхронизатор";
                case ElementType.FORK: return "Разветвитель";
                case ElementType.FLOW: return "Переход";
                case ElementType.FINAL_NODE: return "Конечное состояние";
                case ElementType.INITIAL_NODE: return "Начальное состояние";
                case ElementType.SWIMLANE: return "Дорожка участника";
                case ElementType.UNKNOWN: return "";
                default: throw new ArgumentException();
            }
        }
    }
    public class MergeNode : DiagramElement
    {
        public MergeNode(string id, string inPartition)
            : base(id, inPartition, "")
        {
        }
    }
    public class ActivityNode : DiagramElement
    {
        private readonly string name;

        public ActivityNode(string id, string inPartition, string name)
            : base(id, inPartition, name)
        {
            this.name = name;
        }
        public override string getName()
        {
            return name;
        }
    }
    public class DiagramElement : BaseNode
    {
        protected string inPartition = "";
        protected List<string> idsOut = new List<string>();       // массив ид входящих переходов
        protected List<string> idsIn = new List<string>();        // массив ид выходящих переходов
        protected string description = "";

        public int petriId;


        public DiagramElement(string id, string inPartition, string description)
            : base(id)
        {
            this.inPartition = inPartition;
            this.description = description;
        }
        public override string getSrc()
        {
            throw new System.NotImplementedException();
        }
        public override string getTarget()
        {
            throw new System.NotImplementedException();
        }
        public string getInPartition()
        {
            return inPartition;
        }
        public string getDescription()
        {
            return description;
        }

        public void addIn(string allId)
        {
            string[] ids = allId.Split(' ');
            foreach (string id in ids)
            {
                if (!id.Equals("")) idsIn.Add(id);
            }
        }
        public void addOut(string allId)
        {
            string[] ids = allId.Split(' ');
            foreach (string id in ids)
            {
                if (!id.Equals("")) idsOut.Add(id);
            }
        }

        public string getInId(int index)
        {
            return idsIn[index];
        }

        public string getOutId(int index)
        {
            return idsOut[index];
        }

        public int inSize()
        {
            return idsIn.Count;
        }

        public int outSize()
        {
            return idsOut.Count;
        }
        public override string getName()
        {
            return "";
        }
    }
    public class ControlFlow : BaseNode
    {
        private string src = "";
        private string targets = "";
        private readonly string text;
        public ControlFlow(string id) : base(id) { }

        public ControlFlow(string id, string text)
            : base(id)
        {
            this.text = text;
        }

        public string getText()
        {
            return text;
        }

        public override string getSrc()
        {
            return src;
        }

        public void setSrc(string src)
        {
            this.src = src;
        }

        public override string getTarget()
        {
            return targets;
        }

        public void setTarget(string targets)
        {
            this.targets = targets;
        }
        public override string getName()
        {
            return "";
        }
    }
    public class DecisionNode : DiagramElement
    {
        private string question;
        private readonly List<string> alternatives = new List<string>();     // хранит названия альтернатив

        public DecisionNode(string id, string inPartition, string question)
            : base(id, inPartition, question)
        {
            this.question = question;
        }

        public List<string> findEqualAlternatives()
        {
            List<string> equals = new List<string>();
            for (int i = 0; i < alternatives.Count - 1; i++)
            {
                for (int j = i + 1; j < alternatives.Count; j++)
                {
                    if (alternatives[i].Equals(alternatives[j]) && alternatives[i] != "")
                        equals.Add(alternatives[i]);
                }
            }
            return equals;
        }

        public bool findEmptyAlternative()
        {
            for (int i = 0; i < alternatives.Count; i++)
            {
                if (alternatives[i].Equals("")) return true;
            }
            return false;
        }

        public string getQuestion()
        {
            return question;
        }
        public override string getName()
        {
            return question;
        }
        public void setQuestion(string question)
        {
            this.question = question;
        }

        public void addAlternative(string alternative)
        {
            alternatives.Add(alternative);
        }
        public string getAlternative(int index)
        {
            return alternatives[index];
        }
        public int alternativeSize()
        {
            return alternatives.Count;
        }
    }
    public class ForkNode : DiagramElement
    {
        public ForkNode(string id, string inPartition)
            : base(id, inPartition, "")
        {
        }
    }
    public class JoinNode : DiagramElement
    {
        public JoinNode(string id, string inPartition)
            : base(id, inPartition, "")
        {
        }
    }
    public class Swimlane : BaseNode
    {
        public readonly string name;
        public int childCount = 0;
        public Swimlane(string id, string name) : base(id)
        {
            this.name = name;
        }
        public override string getSrc()
        {
            throw new System.NotImplementedException();
        }
        public override string getTarget()
        {
            throw new System.NotImplementedException();
        }

        public override string getName()
        {
            return name;
        }
    }
    public class ADNodesList
    {
        public List<ADNode> nodes;
        private int diagramElementId = 0;     // Петри ид, присваиваемый элементу

        public ADNodesList()
        {
            nodes = new List<ADNode>();
        }
        /**
         * Возвращает колво элементов, используемых для проверки сетью Петри
         * @return
         */
        public int getPetriElementsCount()
        {
            return diagramElementId;
        }

        /**
         * получить все активности из массива
         * @return
         */
        public List<ADNode> getAllNodes()
        {
            return nodes;
        }
        public List<ActivityNode> getAllActivities()
        {
            List<int> temp = new List<int>();
            return nodes.Where(x => x.getValue().getType() == ElementType.ACTIVITY).ToList().Select(x => (ActivityNode)x.getValue()).ToList();
        }
        public List<Swimlane> getAllSwimlanes()
        {
            return nodes.Where(x => x.getValue().getType() == ElementType.SWIMLANE).ToList().Select(x => (Swimlane)x.getValue()).ToList();
        }
        public List<DecisionNode> getAllDecisionNodes()
        {
            return nodes.Where(x => x.getValue().getType() == ElementType.DECISION).ToList().Select(x => (DecisionNode)x.getValue()).ToList();
        }
        public List<ForkNode> getAllForkNodes()
        {
            return nodes.Where(x => x.getValue().getType() == ElementType.FORK).ToList().Select(x => (ForkNode)x.getValue()).ToList();
        }
        public List<JoinNode> getAllJoinNodes()
        {
            return nodes.Where(x => x.getValue().getType() == ElementType.JOIN).ToList().Select(x => (JoinNode)x.getValue()).ToList();
        }
        public List<ControlFlow> getAllContrFlows()
        {
            return nodes.Where(x => x.getValue().getType() == ElementType.FLOW).ToList().Select(x => (ControlFlow)x.getValue()).ToList();
        }

        /**
         * Найти начальное состояние
         * @return ссылка на узел начального состояние
         */
        public ADNode findInitial()
        {
            for (int i = 0; i < nodes.Count; i++)
            {
                if (nodes[i].getValue().getType() == ElementType.INITIAL_NODE)
                {
                    return nodes[i];
                }
            }
            return null;
        }
        /**
         * Найти конеченое состояние
         * @return ссылка на узел конеченого состояния
         */
        public List<ADNode> findFinal()
        {
            var finalNodes = new List<ADNode>();
            for (int i = 0; i < nodes.Count; i++)
            {
                if (nodes[i].getValue().getType() == ElementType.FINAL_NODE)
                {
                    finalNodes.Add(nodes[i]);
                }
            }
            return finalNodes;
        }

        /**
         * Установить связи между элементами ДА
         */
        public void connect()
        {
            foreach (var node in nodes)
            {
                // связываем все элементы, кроме переходов
                if (node.getValue() is DiagramElement)
                {
                    findNext((DiagramElement)node.getValue(), node);
                }
            }
        }


        /**
         * Найти элементы для связи
         * @param cur текущий элемент, кот надо связать
         * @param curNode
         */
        private void findNext(DiagramElement cur, ADNode curNode)
        {
            // для всех выходный переходов находим таргеты и добавляем ссылки в текущий элемент на таргеты
            for (int i = 0; i < cur.outSize(); i++)
            {
                ControlFlow flow = (ControlFlow)get(cur.getOutId(i));
                ADNode target = getNode(flow.getTarget());
                curNode.next.Add(target);       // прямая связь
                target.prev.Add(curNode);        // обратная связь
            }
        }

        /**
         * Печать связей между элементами
         */
        public void print()
        {
            foreach (ADNode node in nodes)
            {
                if (node.getValue() is DiagramElement)
                {
                    Console.WriteLine("Cur: [" + ((DiagramElement)node.getValue()).petriId + "] " + ((DiagramElement)node.getValue()).getDescription() + " " + node.getValue().getType() + " | ");
                    for (int i = 0; i < node.next.Count; i++)
                    {
                        Console.WriteLine(node.getNext(i).getValue().getType() + " ");
                    }
                    Console.WriteLine(" || ");
                    for (int i = 0; i < node.prev.Count; i++)
                    {
                        Console.WriteLine(node.prev[i].getValue().getType() + " ");
                    }
                    Console.WriteLine("");
                }
            }
        }

        public int size()
        {
            return nodes.Count;
        }
        public void addLast(BaseNode node)
        {
            if (node is DiagramElement)
            {
                ((DiagramElement)node).petriId = diagramElementId;
                diagramElementId++;
            }
            nodes.Add(new ADNode(node));
        }
        public BaseNode get(int index)
        {
            return nodes[index].getValue();
        }

        public BaseNode get(string id)
        {

            ADNode node = nodes.Where(x => x.getValue().getId().Equals(id)).FirstOrDefault();
            if (node == default(ADNode))
                return null;

            return node.getValue() != null ? node.getValue() : null;
        }

        public ADNode getNode(string id)
        {
            var node = nodes.Where(x => x.getValue().getId().Equals(id)).FirstOrDefault();
            return node == default(ADNode) ? null : node;

        }
        public ADNode getNode(int index)
        {
            return nodes[index];
        }

        public ADNode getNodeByPetriIndex(int id)
        {
            ADNode node = nodes.Where(x =>
            {
                if (x.getValue() is DiagramElement)
                    return ((DiagramElement)x.getValue()).petriId == id;
                return false;
            }).FirstOrDefault();
            return node == default(ADNode) ? null : node;
        }
        //endregion
    }
    public class ADNode
    {
        public BaseNode value;
        public List<ADNode> next = new List<ADNode>();
        public List<ADNode> prev = new List<ADNode>();


        public ADNode(BaseNode value)
        {
            this.value = value;
        }

        public ADNode(ADNode old)
        {
            value = old.value;
            next = old.next;
            prev = old.prev;
        }

        //region Getter-Setter
        public BaseNode getValue()
        {
            return value;
        }

        public void setValue(BaseNode value)
        {
            this.value = value;
        }

        public int prevSize() { return prev.Count; }
        public int nextSize() { return next.Count; }

        public ADNode getNext(int index)
        {
            return next[index];
        }
        public ADNode getPrev(int index) { return prev[index]; }

        public List<int> getNextPetriIds()
        {
            return next.Select(x => ((DiagramElement)x.getValue()).petriId).ToList();
        }
        public List<int> getPrevPetriIds()
        {
            return prev.Select(x => ((DiagramElement)x.getValue()).petriId).ToList();
        }
    }
    public class FinalNode : DiagramElement
    {
        public FinalNode(string id, string inPartition)
            : base(id, inPartition, "")
        {
        }
    }
    public class InitialNode : DiagramElement
    {
        public InitialNode(string id, string inPartition)
            : base(id, inPartition, "")
        {

        }
    }
    internal class XmiParser
    {
        private XmlDocument xmlFile = null;
        private readonly ADNodesList adNodesList;
        private readonly List<BaseNode> unknownNodes = new List<BaseNode>();

        public XmiParser(ADNodesList adNodesList)
        {
            this.adNodesList = adNodesList;
        }
        private XmlNode FindActivePackageEl(XmlNodeList xPackagedList, int indSearch)
        {
            int realInd = 0;
            foreach (XmlNode node in xPackagedList)
            {
                var attr = node.Attributes["xsi:type"];
                if (attr == null) continue;
                if (attr.Value.Equals("uml:Activity"))
                {
                    if (realInd == indSearch)
                        return node;
                    else
                        realInd++;
                }
            }
            return null;
        }
        public bool Parse(ADModel diagram, ref bool hasJoinOrFork) {

            xmlFile = new XmlDocument();
            xmlFile.Load(diagram.FilePath);

            XmlNodeList xPackagedList;
            try {
                xPackagedList = xmlFile.GetElementsByTagName("packagedElement");
            } catch (NullReferenceException) {
                //Console.WriteLine("[x] Тег packagedElement не найден");
                return false;
            }


            // получим корневой элемент
            int nTimes = 0;
            XmlNode xRoot = null;
            int nCh = 0;
            while (nTimes == 0 || (xRoot != null && nCh == 0))
            {
                nCh = 0;
                xRoot = FindActivePackageEl(xPackagedList, nTimes);
                if(xRoot != null)
                    foreach (XmlNode node in xRoot.ChildNodes)
                    {
                        if (node.Name != "xmi:Extension")
                            nCh++;
                    }
                nTimes++;
            }

            if (xRoot == null)
                xRoot = FindActivePackageEl(xPackagedList, nTimes - 1);
            if (xRoot == null) {
                //Console.WriteLine("[x] Вид диаграммы не AD");
                return false;
            }

            var attr = xRoot.Attributes["xsi:type"];
            if (attr == null) {
                //Console.WriteLine("[x] Не удалось распарсить xmi файл");
                return false;
            }
            if (!attr.Value.Equals("uml:Activity")) {
                //Console.WriteLine("[x] Вид диаграммы не AD");
                return false;
            }

            // пройтись по всем тегам и создать объекты
            foreach (XmlNode node in xRoot.ChildNodes) {
                if (node.NodeType.ToString() == "Comment")
                    continue;
                var elAttr = node.Attributes["xsi:type"];
                if (elAttr == null) continue;

                if (elAttr.Value == "uml:OpaqueAction" || elAttr.Value == "uml:InitialNode" || elAttr.Value == "uml:ActivityFinalNode" ||
                    elAttr.Value == "uml:FlowFinalNode" || elAttr.Value == "uml:DecisionNode" || elAttr.Value == "uml:MergeNode" ||
                    elAttr.Value == "uml:ForkNode" || elAttr.Value == "uml:JoinNode") {
                    DiagramElement nodeFromXMI = null;
                    switch (elAttr.Value) {
                        // активность
                        case "uml:OpaqueAction":
                            nodeFromXMI = new ActivityNode(node.Attributes["xmi:id"].Value,
                                    AttrAdapter(node.Attributes["inPartition"]), AttrAdapter(node.Attributes["name"]));
                            nodeFromXMI.setType(ElementType.ACTIVITY);
                            adNodesList.addLast(nodeFromXMI);
                            break;
                        // конечное состояние
                        case "uml:ActivityFinalNode":
                        case "uml:FlowFinalNode":
                            nodeFromXMI = new FinalNode(node.Attributes["xmi:id"].Value, AttrAdapter(node.Attributes["inPartition"]));
                            nodeFromXMI.setType(ElementType.FINAL_NODE);
                            adNodesList.addLast(nodeFromXMI);
                            break;
                        // начальное состояние
                        case "uml:InitialNode":
                            nodeFromXMI = new InitialNode(node.Attributes["xmi:id"].Value, AttrAdapter(node.Attributes["inPartition"]));
                            nodeFromXMI.setType(ElementType.INITIAL_NODE);
                            adNodesList.addLast(nodeFromXMI);
                            break;
                        // условный переход
                        case "uml:DecisionNode":
                            nodeFromXMI = new DecisionNode(node.Attributes["xmi:id"].Value, AttrAdapter(node.Attributes["inPartition"]), AttrAdapter(node.Attributes["question"]));
                            nodeFromXMI.setType(ElementType.DECISION);
                            adNodesList.addLast(nodeFromXMI);
                            break;
                        // узел слияния
                        case "uml:MergeNode":
                            nodeFromXMI = new MergeNode(node.Attributes["xmi:id"].Value, AttrAdapter(node.Attributes["inPartition"]));
                            nodeFromXMI.setType(ElementType.MERGE);
                            adNodesList.addLast(nodeFromXMI);
                            break;
                        // разветвитель
                        case "uml:ForkNode":
                            nodeFromXMI = new ForkNode(node.Attributes["xmi:id"].Value, AttrAdapter(node.Attributes["inPartition"]));
                            nodeFromXMI.setType(ElementType.FORK);
                            adNodesList.addLast(nodeFromXMI);
                            hasJoinOrFork = true;
                            break;
                        // синхронизатор
                        case "uml:JoinNode":
                            nodeFromXMI = new JoinNode(node.Attributes["xmi:id"].Value, AttrAdapter(node.Attributes["inPartition"]));
                            nodeFromXMI.setType(ElementType.JOIN);
                            adNodesList.addLast(nodeFromXMI);
                            hasJoinOrFork = true;
                            break;
                    }
                    // добавляем ид входящих и выходящих переходов
                    if (nodeFromXMI != null) {
                        string idsIn = node.Attributes["incoming"] == null ? null : node.Attributes["incoming"].Value;
                        string idsOut = node.Attributes["outgoing"] == null ? null : node.Attributes["outgoing"].Value;
                        nodeFromXMI.addIn(idsIn ?? "");
                        nodeFromXMI.addOut(idsOut ?? "");
                    }
                }
                // создаем переход
                else if (node.Attributes["xsi:type"].Value.Equals("uml:ControlFlow")) {
                    // находим подпись перехода
                    //var markNode = node.ChildNodes[1];
                    //string mark = markNode.Attributes["value"].Value.Trim();        // если подпись является "yes", значит это подпись по умолчанию
                    //mark.Equals("true") ? "" : mark
                    ControlFlow temp = new ControlFlow(node.Attributes["xmi:id"].Value, "");
                    temp.setType(ElementType.FLOW);
                    temp.setSrc(AttrAdapter(node.Attributes["source"]));
                    temp.setTarget(AttrAdapter(node.Attributes["target"]));
                    adNodesList.addLast(temp);
                }
                // создаем дорожку
                else if (node.Attributes["xsi:type"].Value.Equals("uml:ActivityPartition")) {
                    Swimlane temp = new Swimlane(node.Attributes["xmi:id"].Value, AttrAdapter(node.Attributes["name"])) {
                        childCount = node.Attributes["node"] == null ? 0 : node.Attributes["node"].Value.Split().Length
                    };
                    temp.setType(ElementType.SWIMLANE);
                    adNodesList.addLast(temp);

                }
            }

            return true;
        }

        private string AttrAdapter(XmlAttribute attr)
        {
            if (attr == null)
                return "";
            else
            {
                string temp = Regex.Replace(attr.Value.Trim(), @"\s+", " ");
                return temp;
            }
        }
    }
}
