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
    class Element
    {
        public string Type { get; set; }
        public string Name { get; set; }
        public string Id { get; set; }
        public Element(string type, string name, string id)
        {
            Type = type;
            Name = name;
            Id = id;
        }
    }
    class Connection
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public string IdFrom { get; set; }
        public string IdTo { get; set; }
        public Connection(string type, string from, string to, string name)
        {
            Type = type;
            IdFrom = from;
            IdTo = to;
            Name = name;
        }
    }
    class UCDModel : Model
    {
        public List<Element> Elems { get; set; }
        public List<Connection> Conns { get; set; }
        public UCDModel(string filePath) : base(filePath)
        {
            FilePath = filePath;
            Elems = new List<Element>();
            Conns = new List<Connection>();

            var doc = new XmlDocument();
            doc.Load(FilePath);
            if (doc.DocumentElement != null)
                XMItoCSharp(doc.DocumentElement, false);
        }
        public override void XMItoCSharp(XmlElement root, bool isPackage)
        {
            XmlElement rootOfEls = root;
            if (!isPackage)
            {
                if (root.FirstChild != null)
                    rootOfEls = (XmlElement)root.FirstChild;
            }

            foreach (XmlNode childnode in rootOfEls.ChildNodes)
            {
                if (childnode.Name == "packageImport" || childnode.Name == "xmi:Extension" || childnode.Name == "#comment")
                    continue;

                string id = getId(childnode);
                string name = getName(childnode);
                string type = getType(childnode);

                if (childnode.Name == "packagedElement" && (type == "uml:UseCase" || type == "uml:Actor"))
                {
                    Element newElem = new Element(type, name, id);
                    Elems.Add(newElem);
                    if (type == "uml:UseCase")
                        ReadUseCase(childnode);
                }
                else if (childnode.Name == "packagedElement" && (type == "uml:Component" || type == "uml:Class"))
                {
                    ReadPackage(childnode);
                }
                else if (childnode.Name == "packagedElement" && type == "uml:Package")
                {
                    XMItoCSharp((XmlElement)childnode, true);
                }
                else if (childnode.Name == "packagedElement" && type == "uml:Association")
                {
                    List<XmlNode> owEnd = new List<XmlNode>();
                    foreach (XmlNode node in childnode.ChildNodes)
                        if (node.Name == "ownedEnd")
                            owEnd.Add(node);

                    string idFrom = owEnd[0].Attributes.GetNamedItem("type").Value;
                    string idTo = owEnd[1].Attributes.GetNamedItem("type").Value;
                    Connection conn = new Connection("Association", idFrom, idTo, name);
                    Conns.Add(conn);
                }
            }
        }
        private void ReadPackage(XmlNode packRoot)
        {
            foreach (XmlNode childnode in packRoot.ChildNodes)
            {
                if (childnode.Name == "xmi:Extension")
                    continue;

                string id = getId(childnode);
                string name = getName(childnode);
                string type = getType(childnode);

                if (childnode.Name == "ownedUseCase")
                {
                    Element newElem = new Element("uml:UseCase", name, id);
                    Elems.Add(newElem);
                    ReadUseCase(childnode);
                }
                else if (childnode.Name == "packagedElement" && type == "uml:Actor")
                {
                    Element newElem = new Element(type, name, id);
                    Elems.Add(newElem);
                }
                else if (childnode.Name == "packagedElement" && type == "uml:Association")
                {
                    string idFrom = childnode.ChildNodes[1].Attributes.GetNamedItem("type").Value;
                    string idTo = childnode.ChildNodes[2].Attributes.GetNamedItem("type").Value;
                    Connection conn = new Connection("Association", idFrom, idTo, name);
                    Conns.Add(conn);
                }
            }
        }
        private void ReadUseCase(XmlNode ucRoot)
        {
            foreach (XmlNode childnode in ucRoot.ChildNodes)
            {
                if (childnode.Name == "extend")
                {
                    string idFrom = childnode.Attributes.GetNamedItem("extension").Value;
                    string idTo = childnode.Attributes.GetNamedItem("extendedCase").Value;
                    Connection conn = new Connection("Extend", idFrom, idTo, "");
                    Conns.Add(conn);
                }
                else if (childnode.Name == "include")
                {
                    string idTo = childnode.Attributes.GetNamedItem("includingCase").Value;
                    string idFrom = childnode.Attributes.GetNamedItem("addition").Value;
                    Connection conn = new Connection("Include", idFrom, idTo, "");
                    Conns.Add(conn);
                }
            }
        }
        private string getId(XmlNode item)
        {
            XmlNode temp = item.Attributes.GetNamedItem("xmi:id");
            if (temp != null)
                return temp.Value;
            return null;
        }
        private string getType(XmlNode item)
        {
            XmlNode temp = item.Attributes.GetNamedItem("xsi:type");
            if (temp != null)
                return temp.Value;
            return null;
        }
        private string getName(XmlNode item)
        {
            XmlNode temp = item.Attributes.GetNamedItem("name");
            if (temp != null)
                return temp.Value;
            return null;
        }
    }
}
