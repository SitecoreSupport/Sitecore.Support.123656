namespace Sitecore.Support.WFFM.Services.Pipelines.ExportToXml
{
    using Sitecore;
    using Sitecore.Diagnostics;
    using Sitecore.Jobs;
    using Sitecore.Security.Accounts;
    using Sitecore.WFFM.Speak.ViewModel;
    using Sitecore.Text;
    using Sitecore.WFFM.Abstractions.Dependencies;
    using Sitecore.WFFM.Services.Pipelines;
    using System.Linq;
    using System.Text.RegularExpressions;
    using System.Xml;

    public class ExportFormDataToXml
    {
        public void Process(FormExportArgs args)
        {
            Job job = Context.Job;
            if (job != null)
            {
                job.Status.LogInfo(DependenciesManager.ResourceManager.Localize("EXPORTING_DATA"));
            }
            XmlDocument document = new XmlDocument
            {
                InnerXml = new FormPacket(from entry in args.Packet.Entries orderby entry.Timestamp select entry).ToXml() //added orderby to sort the entries before export
            };
            string str = args.Parameters["contextUser"];
            Assert.IsNotNullOrEmpty(str, "contextUser");
            using (new UserSwitcher(str, true))
            {
                ListString str3 = new ListString(Regex.Replace(DependenciesManager.FormRegistryUtil.GetExportRestriction(args.Item.ID.ToString(), string.Empty), "{|}", string.Empty));
                XmlNodeList list = document.SelectNodes("packet/formentry");
                Assert.IsNotNull(list, "roots");
                foreach (string str4 in str3)
                {
                    foreach (XmlNode node in list)
                    {
                        Assert.IsNotNull(node.Attributes, "Attributes");
                        XmlAttribute attribute = node.Attributes[str4];
                        if (attribute != null)
                        {
                            node.Attributes.Remove(attribute);
                        }
                        XmlNodeList list2 = node.SelectNodes($"field[@fieldid='{str4.ToLower()}']");
                        Assert.IsNotNull(list2, "nodeList");
                        foreach (XmlNode node2 in list2)
                        {
                            node.RemoveChild(node2);
                        }
                    }
                }
                args.Result = document.DocumentElement.OuterXml;
            }
        }
    }
}