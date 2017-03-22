namespace Sitecore.Support.WFFM.Services.Pipelines.ExportToExcel
{
    using Sitecore;
    using Sitecore.Diagnostics;
    using Sitecore.Jobs;
    using Sitecore.Security.Accounts;
    using Sitecore.WFFM.Abstractions.Analytics;
    using Sitecore.WFFM.Abstractions.Data;
    using Sitecore.WFFM.Abstractions.Dependencies;
    using Sitecore.WFFM.Services.Pipelines;
    using Sitecore.WFFM.Speak.ViewModel;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Xml;

    public class ExportFormDataToExcel
    {
        private void BuildBody(XmlDocument doc, IFormItem item, FormPacket packet, XmlElement root)
        {
            List<FormData> list = (from entry in packet.Entries orderby entry.Timestamp select entry).ToList<FormData>(); //added orderby to sort the entries before export
            foreach (FormData current in list)
            {
                root.AppendChild(this.BuildRow(current, item, doc));
            }
        }

        private void BuildHeader(XmlDocument doc, IFormItem item, XmlElement root)
        {
            XmlElement newChild = doc.CreateElement("Row");
            string exportRestriction = DependenciesManager.FormRegistryUtil.GetExportRestriction(item.ID.ToString(), string.Empty);
            if (exportRestriction.IndexOf("created", StringComparison.Ordinal) == -1)
            {
                XmlElement element2 = this.CreateHeaderCell("String", "Created", doc);
                newChild.AppendChild(element2);
            }
            foreach (IFieldItem item2 in item.Fields)
            {
                if (exportRestriction.IndexOf(item2.ID.ToString(), StringComparison.Ordinal) == -1)
                {
                    XmlElement element3 = this.CreateHeaderCell("String", item2.FieldDisplayName, doc);
                    newChild.AppendChild(element3);
                }
            }
            root.AppendChild(newChild);
        }

        private XmlElement BuildRow(FormData entry, IFormItem item, XmlDocument xd)
        {
            XmlElement element = xd.CreateElement("Row");
            string exportRestriction = DependenciesManager.FormRegistryUtil.GetExportRestriction(item.ID.ToString(), string.Empty);
            if (exportRestriction.IndexOf("created") == -1)
            {
                XmlElement newChild = this.CreateCell("String", entry.Timestamp.ToLocalTime().ToString("G"), xd);
                element.AppendChild(newChild);
            }
            IFieldItem[] fields = item.Fields;
            for (int i = 0; i < fields.Length; i++)
            {
                Func<FieldData, bool> predicate = null;
                IFieldItem field = fields[i];
                if (exportRestriction.IndexOf(field.ID.ToString(), StringComparison.Ordinal) == -1)
                {
                    if (predicate == null)
                    {
                        predicate = f => f.FieldId == field.ID.Guid;
                    }
                    FieldData data = entry.Fields.FirstOrDefault<FieldData>(predicate);
                    XmlElement element3 = this.CreateCell("String", (data != null) ? data.Value : string.Empty, xd);
                    element.AppendChild(element3);
                }
            }
            return element;
        }

        private XmlElement CreateCell(string sType, string sValue, XmlDocument doc)
        {
            XmlElement element = doc.CreateElement("Cell");
            XmlAttribute node = doc.CreateAttribute("ss", "StyleID", "xmlns");
            node.Value = "xVerdana";
            element.Attributes.Append(node);
            XmlElement newChild = doc.CreateElement("Data");
            XmlAttribute attribute2 = doc.CreateAttribute("ss", "Type", "xmlns");
            attribute2.Value = sType;
            newChild.Attributes.Append(attribute2);
            newChild.InnerText = sValue;
            element.AppendChild(newChild);
            return element;
        }

        private XmlElement CreateHeaderCell(string sType, string sValue, XmlDocument doc)
        {
            XmlElement element = doc.CreateElement("Cell");
            XmlAttribute node = doc.CreateAttribute("ss", "StyleID", "xmlns");
            node.Value = "xBoldVerdana";
            element.Attributes.Append(node);
            XmlElement newChild = doc.CreateElement("Data");
            XmlAttribute attribute2 = doc.CreateAttribute("ss", "Type", "xmlns");
            attribute2.Value = sType;
            newChild.Attributes.Append(attribute2);
            newChild.InnerText = sValue;
            element.AppendChild(newChild);
            return element;
        }

        public void Process(FormExportArgs args)
        {
            Job job = Context.Job;
            if (job != null)
            {
                job.Status.LogInfo(DependenciesManager.ResourceManager.Localize("EXPORTING_DATA"));
            }
            string str = args.Parameters["contextUser"];
            Assert.IsNotNullOrEmpty(str, "contextUser");
            using (new UserSwitcher(str, true))
            {
                XmlDocument doc = new XmlDocument();
                XmlElement newChild = doc.CreateElement("ss:Workbook");
                XmlAttribute node = doc.CreateAttribute("xmlns");
                node.Value = "urn:schemas-microsoft-com:office:spreadsheet";
                newChild.Attributes.Append(node);
                XmlAttribute attribute2 = doc.CreateAttribute("xmlns:o");
                attribute2.Value = "urn:schemas-microsoft-com:office:office";
                newChild.Attributes.Append(attribute2);
                XmlAttribute attribute3 = doc.CreateAttribute("xmlns:x");
                attribute3.Value = "urn:schemas-microsoft-com:office:excel";
                newChild.Attributes.Append(attribute3);
                XmlAttribute attribute4 = doc.CreateAttribute("xmlns:ss");
                attribute4.Value = "urn:schemas-microsoft-com:office:spreadsheet";
                newChild.Attributes.Append(attribute4);
                XmlAttribute attribute5 = doc.CreateAttribute("xmlns:html");
                attribute5.Value = "http://www.w3.org/TR/REC-html40";
                newChild.Attributes.Append(attribute5);
                doc.AppendChild(newChild);
                XmlElement element2 = doc.CreateElement("Styles");
                newChild.AppendChild(element2);
                XmlElement element3 = doc.CreateElement("Style");
                XmlAttribute attribute6 = doc.CreateAttribute("ss", "ID", "xmlns");
                attribute6.Value = "xBoldVerdana";
                element3.Attributes.Append(attribute6);
                element2.AppendChild(element3);
                XmlElement element4 = doc.CreateElement("Font");
                XmlAttribute attribute7 = doc.CreateAttribute("ss", "Bold", "xmlns");
                attribute7.Value = "1";
                element4.Attributes.Append(attribute7);
                XmlAttribute attribute8 = doc.CreateAttribute("ss", "FontName", "xmlns");
                attribute8.Value = "verdana";
                element4.Attributes.Append(attribute8);
                element3.AppendChild(element4);
                element3 = doc.CreateElement("Style");
                attribute6 = doc.CreateAttribute("ss", "ID", "xmlns");
                attribute6.Value = "xVerdana";
                element3.Attributes.Append(attribute6);
                element2.AppendChild(element3);
                element4 = doc.CreateElement("Font");
                attribute8 = doc.CreateAttribute("ss", "FontName", "xmlns");
                attribute8.Value = "verdana";
                element4.Attributes.Append(attribute8);
                element3.AppendChild(element4);
                XmlElement element5 = doc.CreateElement("Worksheet");
                XmlAttribute attribute9 = doc.CreateAttribute("ss", "Name", "xmlns");
                attribute9.Value = "Sheet1";
                element5.Attributes.Append(attribute9);
                newChild.AppendChild(element5);
                XmlElement element6 = doc.CreateElement("Table");
                XmlAttribute attribute10 = doc.CreateAttribute("ss", "DefaultColumnWidth", "xmlns");
                attribute10.Value = "130";
                element6.Attributes.Append(attribute10);
                element5.AppendChild(element6);
                this.BuildHeader(doc, args.Item, element6);
                this.BuildBody(doc, args.Item, args.Packet, element6);
                XmlElement element7 = doc.CreateElement("WorksheetOptions");
                XmlElement element8 = doc.CreateElement("Selected");
                XmlElement element9 = doc.CreateElement("Panes");
                XmlElement element10 = doc.CreateElement("Pane");
                XmlElement element11 = doc.CreateElement("Number");
                element11.InnerText = "1";
                XmlElement element12 = doc.CreateElement("ActiveCol");
                element12.InnerText = "1";
                element10.AppendChild(element12);
                element10.AppendChild(element11);
                element9.AppendChild(element10);
                element7.AppendChild(element9);
                element7.AppendChild(element8);
                element5.AppendChild(element7);
                args.Result = "<?xml version=\"1.0\"?>" + doc.InnerXml.Replace("xmlns:ss=\"xmlns\"", "");
            }
        }
    }
}