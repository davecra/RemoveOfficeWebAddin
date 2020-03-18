using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace RemoveOfficeWebAddin
{
    [ComVisible(true)]
    public class DisableWebAddinsRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public DisableWebAddinsRibbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            try
            {
                // the property tag for the user profile entry
                const string PR_EMSMDB_SECTION_UID = "http://schemas.microsoft.com/mapi/proptag/0x3D150102";
                // laod the base ribbon
                string LstrRibbonXml = GetResourceText("RemoveOfficeWebAddin.DisableWebAddinsRibbon.xml");
                // code base is the install directory where the vsto file is located
                string LstrPath = new Uri(Assembly.GetExecutingAssembly().CodeBase).LocalPath.ToString();
                LstrPath = LstrPath.Substring(0, LstrPath.LastIndexOf("\\"));
                string LstrSettingFile = Path.Combine(LstrPath, "RemovedCustomizations.txt");
                int iCount = 0; // for namespaces

                // open the settings file and get the list of customiztions:
                //  - type,TabName,Type_AppID_Name. 
                // For example:
                //  - customgroup,TabNewMailMessage,Group_5febe0ec-e536-4275-bd02-66818bf9e191_msgEditGroup_TabNewMailMessage
                StreamReader LobjReader = new StreamReader(LstrSettingFile);
                List<string> LobjLines = LobjReader.ReadToEnd().Split('\n').ToList<string>();
                LobjReader.Close();

                // build the List from the data and load into array of class objects
                List<RibbonRemovalData> LobjData = new List<RibbonRemovalData>();
                string LstrWebAddinID = LobjLines[0].Trim();
                int LintCount = 0;
                foreach (string LstrLine in LobjLines)
                {
                    LintCount++;
                    if (LintCount == 1) continue; // skip the first one, because it is our ID
                    if (string.IsNullOrEmpty(LstrLine)) continue;
                    string[] LstrParts = LstrLine.Split(',');
                    LobjData.Add(new RibbonRemovalData(
                                        (RibbonRemovalData.ItemTypesEnum)Enum.Parse(typeof(RibbonRemovalData.ItemTypesEnum), LstrParts[0].Trim()),
                                        LstrParts[1].Trim(),
                                        LstrParts[2].Trim()));
                }
                // build and create a list of namespaces for each account loaded in Outlook
                // what we will do is create an entry for every account since we do not know if
                // which one the add-in is loaded in. It could be one, or all
                Dictionary<string, string> LobjAccountUIDs = new Dictionary<string, string>();
                // loop through each account and add a ribbon NS element for each one
                foreach (Outlook.Account account in Globals.ThisAddIn.Application.Session.Accounts)
                {
                    iCount++;
                    Outlook.PropertyAccessor propertyAccessor = account.CurrentUser.PropertyAccessor;
                    string ns = "x" + iCount.ToString();
                    LobjAccountUIDs.Add(ns, LstrWebAddinID + "_" + propertyAccessor.BinaryToString(propertyAccessor.GetProperty(PR_EMSMDB_SECTION_UID)));
                    break;
                }
                // now build custom ribbon tabs into a list
                string LstrTabsToAdd = "";
                foreach (string LstrTab in LobjData.Where(item => item.ItemType == RibbonRemovalData.ItemTypesEnum.customtab)
                                                  .Select(item => item.TabName).Distinct().ToList<string>())
                {
                    // add a copy of the group for each account
                    foreach (KeyValuePair<string, string> LobjUID in LobjAccountUIDs)
                    {
                        // <tab idQ="xyz1:Tab_5febe0ec-e536-4275-bd02-66818bf9e191_MyTab" visible="false" />
                        LstrTabsToAdd += "<tab idQ='" + LobjUID.Key + ":" + LstrTab + "' visible='false' />\n";
                    }
                }
                // now build the built-in tabs and add the groups
                foreach (string LstrTab in LobjData.Where(item => item.ItemType == RibbonRemovalData.ItemTypesEnum.customgroup)
                                                  .Select(item => item.TabName).Distinct().ToList<string>())
                {
                    string LstrTabInfo = "<tab idMso=\"" + LstrTab + "\">\n";
                    List<string> LobjGroups = new List<string>();
                    // get each group in this tab
                    foreach (string LstrGroup in LobjData.Where(item => item.ItemType == RibbonRemovalData.ItemTypesEnum.customgroup &&
                                                                        item.TabName == LstrTab)
                                                        .Select(item => item.GroupName).ToList<string>())
                    {
                        // add a copy of the group for each account
                        foreach (KeyValuePair<string, string> LobjUID in LobjAccountUIDs)
                        {
                            // <group idQ="xyz1:Group_5febe0ec-e536-4275-bd02-66818bf9e191_msgEditGroup_TabNewMailMessage" visible="false" />
                            LstrTabInfo += "<group idQ=\"" + LobjUID.Key + ":" + LstrGroup + "\" visible=\"false\" />\n";
                        }
                    }
                    LstrTabInfo += "</tab>\n";
                    // now add the tab to the collection
                    LstrTabsToAdd += LstrTabInfo;
                }

                // add the custom ribbon tabs for each
                LstrRibbonXml = LstrRibbonXml.Replace("<!--{{ADDIN_RIBBON_REMOVE_PART}}-->", LstrTabsToAdd);
                // add the namespaces to the root of the ribbon note
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(LstrRibbonXml);
                XmlNode ele = doc.DocumentElement.GetElementsByTagName("ribbon")[0];
                // append the attributes for each namespace for each UID we found / per account
                foreach (KeyValuePair<string, string> kvp in LobjAccountUIDs)
                {
                    XmlAttribute attr = doc.CreateAttribute("xmlns:" + kvp.Key);
                    attr.Value = kvp.Value;
                    ele.Attributes.Append(attr);
                }
                LstrRibbonXml = doc.OuterXml;

                // all done - return the results
                return LstrRibbonXml;
            }
            catch(Exception PobjEx)
            {
                MessageBox.Show("Ribbon customization failed to load: \n\n" + PobjEx.ToString(), 
                                "Disable Web Add-ins", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            try
            {
                this.ribbon = ribbonUI;
            }
            catch(Exception PobjEx)
            {
                MessageBox.Show(PobjEx.ToString());
            }
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }

    public class RibbonRemovalData
    {
        public enum ItemTypesEnum { customtab, customgroup }
        public ItemTypesEnum ItemType { get; private set; }
        public string TabName { get; private set; }
        public string GroupName { get; private set; }

        public RibbonRemovalData(ItemTypesEnum PobjType, string PstrTabName, string PstrGroupName)
        {
            ItemType = PobjType;
            TabName = PstrTabName;
            GroupName = PstrGroupName;
        }
    }
}
