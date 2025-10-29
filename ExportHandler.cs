using System;
using System.IO;
using Visio = Microsoft.Office.Interop.Visio;

namespace Geradeaus.VisioJsonExport
{
    public class ExportHandler
    {
        public Visio.Document Document { get; set; }
        public VisioModel VisioModel { get; set; } = new VisioModel();

        public ExportHandler(Visio.Document document)
        {
            Document = document;
        }

        public void Parse()
        {
            VisioModel.ExportTime = DateTime.Now.ToString();

            VisioModel.Document.Name = Document.Name;
            VisioModel.Document.FullName = Document.FullName;
            VisioModel.Document.Path = Document.Path;
            VisioModel.Document.Title = Document.Title;
            VisioModel.Document.Subject = Document.Subject;
            VisioModel.Document.Description = Document.Description;
            VisioModel.Document.Creator = Document.Creator;
            VisioModel.Document.Manager = Document.Manager;
            VisioModel.Document.Company = Document.Company;
            VisioModel.Document.Category = Document.Category;
            VisioModel.Document.Keywords = Document.Keywords;
            VisioModel.Document.Language = Document.Language.ToString();
            VisioModel.Document.TimeCreated = Document.TimeCreated.ToString();
            VisioModel.Document.TimeEdited = Document.TimeEdited.ToString();
            VisioModel.Document.TimeSaved = Document.TimeSaved.ToString();

            if (Convert.ToBoolean(Document.DocumentSheet.SectionExists[(short)Visio.VisSectionIndices.visSectionUser, (short)Visio.VisExistsFlags.visExistsAnywhere]))
            {
                Visio.Section section = Document.DocumentSheet.Section[(short)Visio.VisSectionIndices.visSectionUser];
                for (short i = 0; i < section.Count; i++)
                {
                    UserRow userRow = new UserRow();
                    Visio.Row row = section[i];

                    userRow.Value = row[(short)Visio.VisCellIndices.visUserValue].ResultStr[""];
                    userRow.Prompt = row[(short)Visio.VisCellIndices.visUserPrompt].ResultStr[""];
                    VisioModel.Document.UserRows[row.Name] = userRow;
                }
            }

            if (Convert.ToBoolean(Document.DocumentSheet.SectionExists[(short)Visio.VisSectionIndices.visSectionProp, (short)Visio.VisExistsFlags.visExistsAnywhere]))
            {
                Visio.Section section = Document.DocumentSheet.Section[(short)Visio.VisSectionIndices.visSectionProp];
                for (short i = 0; i < section.Count; i++)
                {
                    PropRow propRow = new PropRow();
                    Visio.Row row = section[i];

                    propRow.Label = row[(short)Visio.VisCellIndices.visCustPropsLabel].ResultStr[""];
                    propRow.Prompt = row[(short)Visio.VisCellIndices.visCustPropsPrompt].ResultStr[""];
                    propRow.Type = Convert.ToInt32(row[(short)Visio.VisCellIndices.visCustPropsType].ResultStr[""]);
                    propRow.Format = row[(short)Visio.VisCellIndices.visCustPropsFormat].ResultStr[""];
                    propRow.Value = row[(short)Visio.VisCellIndices.visCustPropsValue].ResultStr[""];

                    VisioModel.Document.PropRows[row.Name] = propRow;
                }
            }

            foreach (Visio.Master master in Document.Masters)
            {
                Master master1 = new Master();
                master1.Name = master.Name;
                master1.NameU = master.NameU;
                master1.ID = master.ID;
                master1.OneD = Convert.ToBoolean(master.OneD);

                VisioModel.Document.Masters[master1.Name] = master1;
            }

            foreach (Visio.Page page in Document.Pages)
            {
                Page page1 = new Page();
                page1.Name = page.Name;
                page1.NameU = page.NameU;
                page1.ID = page.ID;

                if (Convert.ToBoolean(page.PageSheet.SectionExists[(short)Visio.VisSectionIndices.visSectionUser, (short)Visio.VisExistsFlags.visExistsAnywhere]))
                {
                    Visio.Section section = page.PageSheet.Section[(short)Visio.VisSectionIndices.visSectionUser];
                    for (short i = 0; i < section.Count; i++)
                    {
                        UserRow userRow = new UserRow();
                        Visio.Row row = section[i];

                        userRow.Value = row[(short)Visio.VisCellIndices.visUserValue].ResultStr[""];
                        userRow.Prompt = row[(short)Visio.VisCellIndices.visUserPrompt].ResultStr[""];

                        page1.UserRows[row.Name] = userRow;
                    }
                }

                if (Convert.ToBoolean(page.PageSheet.SectionExists[(short)Visio.VisSectionIndices.visSectionProp, (short)Visio.VisExistsFlags.visExistsAnywhere]))
                {
                    Visio.Section section = page.PageSheet.Section[(short)Visio.VisSectionIndices.visSectionProp];
                    for (short i = 0; i < section.Count; i++)
                    {
                        PropRow propRow = new PropRow();
                        Visio.Row row = section[i];

                        propRow.Label = row[(short)Visio.VisCellIndices.visCustPropsLabel].ResultStr[""];
                        propRow.Prompt = row[(short)Visio.VisCellIndices.visCustPropsPrompt].ResultStr[""];
                        propRow.Type = Convert.ToInt32(row[(short)Visio.VisCellIndices.visCustPropsType].ResultStr[""]);
                        propRow.Format = row[(short)Visio.VisCellIndices.visCustPropsFormat].ResultStr[""];
                        propRow.Value = row[(short)Visio.VisCellIndices.visCustPropsValue].ResultStr[""];

                        page1.PropRows[row.Name] = propRow;
                    }
                }

                foreach (Visio.Shape shape in page.Shapes)
                {
                    Shape shape1 = new Shape();
                    shape1.Name = shape.Name;
                    shape1.NameU = shape.NameU;
                    shape1.NameID = shape.NameID;
                    shape1.ID = shape.ID;
                    shape1.Text = shape.Text;
                    shape1.OneD = Convert.ToBoolean(shape.OneD);
                    if (shape.Master != null)
                    {
                        shape1.Master = shape.Master.Name;
                    }

                    if (Convert.ToBoolean(shape.SectionExists[(short)Visio.VisSectionIndices.visSectionUser, (short)Visio.VisExistsFlags.visExistsAnywhere]))
                    {
                        Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionUser];
                        for (short i = 0; i < section.Count; i++)
                        {
                            UserRow userRow = new UserRow();
                            Visio.Row row = section[i];

                            userRow.Value = row[(short)Visio.VisCellIndices.visUserValue].ResultStr[""];
                            userRow.Prompt = row[(short)Visio.VisCellIndices.visUserPrompt].ResultStr[""];

                            shape1.UserRows[row.Name] = userRow;
                        }
                    }

                    if (Convert.ToBoolean(shape.SectionExists[(short)Visio.VisSectionIndices.visSectionProp, (short)Visio.VisExistsFlags.visExistsAnywhere]))
                    {
                        Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionProp];
                        for (short i = 0; i < section.Count; i++)
                        {
                            PropRow propRow = new PropRow();
                            Visio.Row row = section[i];

                            propRow.Label = row[(short)Visio.VisCellIndices.visCustPropsLabel].ResultStr[""];
                            propRow.Prompt = row[(short)Visio.VisCellIndices.visCustPropsPrompt].ResultStr[""];
                            propRow.Type = Convert.ToInt32(row[(short)Visio.VisCellIndices.visCustPropsType].ResultStr[""]);
                            propRow.Format = row[(short)Visio.VisCellIndices.visCustPropsFormat].ResultStr[""];
                            propRow.Value = row[(short)Visio.VisCellIndices.visCustPropsValue].ResultStr[""];

                            shape1.PropRows[row.Name] = propRow;
                        }
                    }

                    if (Convert.ToBoolean(shape.SectionExists[(short)Visio.VisSectionIndices.visSectionConnectionPts, (short)Visio.VisExistsFlags.visExistsAnywhere]))
                    {
                        Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionConnectionPts];
                        for (short i = 0; i < section.Count; i++)
                        {
                            ConnectionPoint connectionPoint = new ConnectionPoint();
                            Visio.Row row = section[i];

                            if (row.Name == "")
                            {
                                shape1.ConnectionPoints[row.Index.ToString()] = connectionPoint;
                            }
                            else
                            {
                                connectionPoint.D = row[(short)Visio.VisCellIndices.visCnnctD].ResultStr[""];
                                shape1.ConnectionPoints[row.Name] = connectionPoint;
                            }
                        }
                    }

                    for (short i = 1; i <= shape.LayerCount; i++)
                    {
                        shape1.Layers.Add(shape.Layer[i].Index);
                    }
                    
                    foreach (Visio.Connect connect in shape.Connects)
                    {
                        Connect connect1 = new Connect();
                        connect1.Index = connect.Index;

                        connect1.ToShape = connect.ToSheet.ID;

                        try
                        {
                            if (connect.ToCell.RowName == "")
                            {
                                connect1.ToPoint = connect.ToCell.ContainingRow.Index.ToString();
                            }
                            else
                            {
                                connect1.ToPoint = connect.ToCell.RowName;
                                connect1.ToPointD = connect.ToCell.ContainingRow[(short)Visio.VisCellIndices.visCnnctD].ResultStr[""];
                            }
                        }
                        catch { }
                        
                        shape1.Connects.Add(connect1);
                    }

                    foreach (Visio.Layer layer in page.Layers)
                    {
                        Layer layer1 = new Layer();
                        layer1.Index = layer.Index;

                        layer1.Name = layer.Name;
                        layer1.NameU = layer.NameU;
                        layer1.Visible = Convert.ToBoolean(layer.CellsC[(short)Visio.VisCellIndices.visLayerVisible].ResultIU);
                        layer1.Print = Convert.ToBoolean(layer.CellsC[(short)Visio.VisCellIndices.visLayerPrint].ResultIU);
                        layer1.Active = Convert.ToBoolean(layer.CellsC[(short)Visio.VisCellIndices.visLayerActive].ResultIU);
                        layer1.Lock = Convert.ToBoolean(layer.CellsC[(short)Visio.VisCellIndices.visLayerLock].ResultIU);
                        layer1.Snap = Convert.ToBoolean(layer.CellsC[(short)Visio.VisCellIndices.visLayerSnap].ResultIU);
                        layer1.Glue = Convert.ToBoolean(layer.CellsC[(short)Visio.VisCellIndices.visLayerGlue].ResultIU);
                        layer1.Color = Convert.ToInt32(layer.CellsC[(short)Visio.VisCellIndices.visLayerColor].ResultIU);
                        layer1.ColorTrans = Convert.ToDouble(layer.CellsC[(short)Visio.VisCellIndices.visLayerColorTrans].ResultIU);

                        page1.Layers[layer1.Index] = layer1;
                    }

                    page1.Shapes[shape1.ID] = shape1;
                }

                VisioModel.Document.Pages[page1.Name] = page1;
            }
        }

        public void Export(string path)
        {
            string json = Newtonsoft.Json.JsonConvert.SerializeObject(VisioModel);
            File.WriteAllText(path, json);
        }
    }
}