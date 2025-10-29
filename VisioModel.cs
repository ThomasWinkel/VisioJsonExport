using System;
using System.Collections.Generic;

namespace Geradeaus.VisioJsonExport
{
    public class VisioModel
    {
        public Document Document { get; set; } = new Document();
        public String ExportTime { get; set; }
    }

    public class ConnectionPoint
    {
        public string D { get; set; }
    }

    public class UserRow
    {
        public string Value { get; set; }
        public string Prompt { get; set; }
    }

    public class PropRow
    {
        public string Label { get; set; }
        public string Prompt { get; set; }
        public int Type { get; set; }
        public string Format { get; set; }
        public string Value { get; set; }
    }

    public class Connect
    {
        public int Index { get; set; }
        public int ToShape { get; set; }
        public string ToPoint { get; set; }
        public string ToPointD { get; set; }
    }

    public class Shape
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public string NameU { get; set; }
        public string NameID { get; set; }
        public string Master { get; set; }
        public string Text { get; set; }
        public bool OneD { get; set; }
        public Dictionary<string, UserRow> UserRows { get; set; } = new Dictionary<string, UserRow>();
        public Dictionary<string, PropRow> PropRows { get; set; } = new Dictionary<string, PropRow>();
        public Dictionary<string, ConnectionPoint> ConnectionPoints { get; set; } = new Dictionary<string, ConnectionPoint>();
        public List<Connect> Connects { get; set; } = new List<Connect>();
        public List<int> Layers { get; set; } = new List<int>();
    }

    public class Layer
    {
        public int Index { get; set; }
        public string Name { get; set; }
        public string NameU { get; set; }
        public bool Visible { get; set; }
        public bool Print { get; set; }
        public bool Active { get; set; }
        public bool Lock { get; set; }
        public bool Snap { get; set; }
        public bool Glue { get; set; }
        public int Color { get; set; }
        public double ColorTrans { get; set; }
    }

    public class Page
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public string NameU { get; set; }
        public Dictionary<string, UserRow> UserRows { get; set; } = new Dictionary<string, UserRow>();
        public Dictionary<string, PropRow> PropRows { get; set; } = new Dictionary<string, PropRow>();
        public Dictionary<int, Shape> Shapes { get; set; } = new Dictionary<int, Shape>();
        public Dictionary<int, Layer> Layers { get; set; } = new Dictionary<int, Layer>();
    }

    public class Master
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public string NameU { get; set; }
        public bool OneD { get; set; }
    }

    public class Document
    {
        public string Name { get; set; }
        public string FullName { get; set; }
        public string Path { get; set; }
        public string Title { get; set; }
        public string Subject { get; set; }
        public string Description { get; set; }
        public string Creator { get; set; }
        public string Manager { get; set; }
        public string Company { get; set; }
        public string Category { get; set; }
        public string Keywords { get; set; }
        public string Language { get; set; }
        public string TimeCreated { get; set; }
        public string TimeEdited { get; set; }
        public string TimeSaved { get; set; }
        public Dictionary<string, UserRow> UserRows { get; set; } = new Dictionary<string, UserRow>();
        public Dictionary<string, PropRow> PropRows { get; set; } = new Dictionary<string, PropRow>();
        public Dictionary<string, Master> Masters { get; set; } = new Dictionary<string, Master>();
        public Dictionary<string, Page> Pages { get; set; } = new Dictionary<string, Page>();
    }
}