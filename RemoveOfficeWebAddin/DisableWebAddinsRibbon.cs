using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace RemoveOfficeWebAddin
{
    [ComVisible(true)]
    public class DisableWebAddinsRibbon : Office.IRibbonExtensibility
    {
        public Office.IRibbonUI Ribbon { get; private set; }
        // build the List from the data and load into array of class objects
        private List<RibbonRemovalData> MobjData = new List<RibbonRemovalData>();
        private Dictionary<string, string> MobjAccountUIDs = new Dictionary<string, string>();
        private string MstrWebAddinID = null;
        // the property tag for the user profile entry
        private const string PR_EMSMDB_SECTION_UID = "http://schemas.microsoft.com/mapi/proptag/0x3D150102";
        // remove parts in the XML
        private const string REMOVE_TABS_COMMENT = "<!--{{ADDIN_TABS_REMOVE_PART}}-->";
        private const string REMOVE_CONTEXTTABS_COMMENT = "<!--{{ADDIN_CONTEXTUALTABS_REMOVE_PART}}-->";
        
        /// <summary>
        /// CTOR - Do nothing here
        /// </summary>
        public DisableWebAddinsRibbon()
        {
        }

        #region IRibbonExtensibility Members
        public string GetCustomUI(string ribbonID)
        {
            try
            {
                LoadConfiguration(); // load the RemovedCustomizations.txt
                LoadNamespaceInfo(); // load each Outlook account namespace

                // load the base ribbon from resources
                string LstrRibbonXml = GetResourceText("RemoveOfficeWebAddin.DisableWebAddinsRibbon.xml");

                // -------------------
                // CREATE THE TABSETS
                // -------------------
                string LstrTabSetsXml = WriteTabSetXml();
                if(!string.IsNullOrEmpty(LstrTabSetsXml))
                {
                    // wrap in the contextual tab tab
                    LstrTabSetsXml = "<contextualTabs>\n" 
                                        + LstrTabSetsXml + 
                                     "</contextualTabs>\n";
                    LstrRibbonXml = LstrRibbonXml.Replace(REMOVE_CONTEXTTABS_COMMENT, LstrTabSetsXml);
                }

                // -------------------
                // CREATE TABS/GROUPS
                // -------------------
                // Now write all the tabs alone without groups
                string LstrTabsXml = WriteTabsXml(null, true);
                // Now write all the tabs with groups
                LstrTabsXml += WriteTabsXml();
                if(!string.IsNullOrEmpty(LstrTabsXml))
                {
                    LstrTabsXml = "<tabs>\n" +
                                    LstrTabsXml +
                                  "</tabs>\n";
                    LstrRibbonXml = LstrRibbonXml.Replace(REMOVE_TABS_COMMENT, LstrTabsXml);
                }

                // -------------------
                // NAMESPACE ADD
                // -------------------
                // add the namespaces to the root of the ribbon note
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(LstrRibbonXml);
                XmlNode ele = doc.DocumentElement.GetElementsByTagName("ribbon")[0];
                // append the attributes for each namespace for each UID we found / per account
                foreach (KeyValuePair<string, string> kvp in MobjAccountUIDs)
                {
                    XmlAttribute attr = doc.CreateAttribute("xmlns:" + kvp.Key);
                    attr.Value = kvp.Value;
                    ele.Attributes.Append(attr);
                }
                LstrRibbonXml = doc.OuterXml;

                // -------------------
                // RETURN XML
                // -------------------
                // all done - return the results
                return LstrRibbonXml;
            }
            catch(Exception PobjEx)
            {
                // NOTE: All functions throw exceptions up here, thus this is
                //       the only error handler. Show the exception info:
                MessageBox.Show("Ribbon customization failed to load: \n\n" + PobjEx.ToString(), 
                                "Disable Web Add-ins", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        /// <summary>
        /// Gets each unique TabSet and then adds all the ribbons and groups
        /// for each item to be disbaled in the Ribbon
        /// </summary>
        /// <returns></returns>
        private string WriteTabSetXml()
        {
            try
            {
                string LstrReturnInfo = "";
                // get a unique list of each tabset needed
                List<string> LobjTabSetNames = MobjData.Where(item => item.ItemType == RibbonRemovalData.ItemTypesEnum.customgroup &&
                                                                     !string.IsNullOrEmpty(item.TabSet))
                                                       .Select(item => item.TabSet).Distinct().ToList<string>();
                // first build any tabsets
                foreach (string LstrTabSet in LobjTabSetNames)
                {
                    LstrReturnInfo += "<tabSet idMso='" + LstrTabSet + "'>";
                    LstrReturnInfo += WriteTabsXml(LstrTabSet);
                    LstrReturnInfo += "</tabSet>";
                }
                return LstrReturnInfo;
            }
            catch(Exception PobjEx)
            {
                throw new Exception("Unable to write TabSet XML: " + PobjEx.Message);
            }
        }

        /// <summary>
        /// Writes out each tab alone, tab + groups or tabset tab+groups for each
        /// item to be disabled in the Ribbon
        /// </summary>
        /// <param name="PstrTabSet"></param>
        /// <param name="PbolTabsOnly"></param>
        /// <returns></returns>
        private string WriteTabsXml(string PstrTabSet = "", bool PbolTabsOnly = false)
        {
            try
            {
                List<string> LobjTabNames = null;
                if (!string.IsNullOrEmpty(PstrTabSet))
                {
                    // get a list of tabs for the specific tabset
                    LobjTabNames = MobjData.Where(item => item.ItemType == RibbonRemovalData.ItemTypesEnum.customgroup &&
                                                          !string.IsNullOrEmpty(item.TabSet) &&
                                                          item.TabSet == PstrTabSet)
                                           .Select(item => item.TabName).Distinct().ToList<string>();
                }
                else
                {
                    if (PbolTabsOnly)
                    {
                        // get a list of all the top level tabs
                        LobjTabNames = MobjData.Where(item => item.ItemType == RibbonRemovalData.ItemTypesEnum.customtab &&
                                                              string.IsNullOrEmpty(item.TabSet))
                                               .Select(item => item.TabName).ToList<string>();
                    }
                    else
                    {
                        // get a list of all the top level tabs with groups
                        LobjTabNames = MobjData.Where(item => item.ItemType == RibbonRemovalData.ItemTypesEnum.customgroup &&
                                                              string.IsNullOrEmpty(item.TabSet))
                                               .Select(item => item.TabName).Distinct().ToList<string>();
                    }
                }

                string LstrReturnInfo = "";
                // now build the tab
                foreach (string LstrTab in LobjTabNames)
                {
                    if (PbolTabsOnly)
                    {
                        // We remove custom tabs added by a web add-in here
                        foreach (KeyValuePair<string, string> LobjUID in MobjAccountUIDs)
                        {
                            // <tab idQ="xyz1:Tab_5febe0ec-e536-4275-bd02-66818bf9e191_MyTab" visible="false" />
                            LstrReturnInfo += "<tab idQ='" + LobjUID.Key + ":" + LstrTab + "' visible='false' />\n";
                        }
                    }
                    else
                    {
                        // otherwise we remove items added to an existing Office tab here
                        LstrReturnInfo += "<tab idQ=\"" + LstrTab + "\">";
                        // now add the groups
                        LstrReturnInfo += WriteGroupsXml(LstrTab);
                        LstrReturnInfo += "</tab>";
                    }
                }

                return LstrReturnInfo;
            }
            catch(Exception PobjEx)
            {
                throw new Exception("Unable to write tab data: " + PobjEx.Message);
            }
        }

        /// <summary>
        /// Writes out each Group XML to be disabled in the Ribbon
        /// </summary>
        /// <param name="LstrTab"></param>
        /// <returns></returns>
        private string WriteGroupsXml(string LstrTab)
        {
            try
            {
                string LstrReturnInfo = "";
                // get each group in this tab
                foreach (string LstrGroup in MobjData.Where(item => item.ItemType == RibbonRemovalData.ItemTypesEnum.customgroup &&
                                                                    item.TabName == LstrTab)
                                                    .Select(item => item.GroupName).ToList<string>())
                {
                    // add a copy of the group for each account
                    foreach (KeyValuePair<string, string> LobjUID in MobjAccountUIDs)
                    {
                        // <group idQ="xyz1:Group_5febe0ec-e536-4275-bd02-66818bf9e191_msgEditGroup_TabNewMailMessage" visible="false" />
                        LstrReturnInfo += "<group idQ=\"" + LobjUID.Key + ":" + LstrGroup + "\" visible=\"false\" />\n";
                    }
                }
                return LstrReturnInfo;
            }
            catch(Exception PobjEx)
            {
                // pass it along
                throw new Exception(PobjEx.Message);
            }
        }

        /// <summary>
        /// Loads the configuraiton for items to remove into memory
        /// </summary>
        private void LoadConfiguration()
        {
            try
            {
                // code base is the install directory where the vsto file is located
                string LstrPath = new Uri(Assembly.GetExecutingAssembly().CodeBase).LocalPath.ToString();
                LstrPath = LstrPath.Substring(0, LstrPath.LastIndexOf("\\"));
                string LstrSettingFile = Path.Combine(LstrPath, "RemovedCustomizations.txt");
                // open the settings file and get the list of customiztions:
                //  - type,TabName,Type_AppID_Name. 
                // For example:
                //  - customgroup,TabNewMailMessage,Group_5febe0ec-e536-4275-bd02-66818bf9e191_msgEditGroup_TabNewMailMessage
                StreamReader LobjReader = new StreamReader(LstrSettingFile);
                List<string> LobjLines = LobjReader.ReadToEnd().Split('\n').ToList<string>();
                LobjReader.Close();
                MstrWebAddinID = LobjLines[0].Trim(); // set the ID
                int LintCount = 0;
                foreach (string LstrLine in LobjLines)
                {
                    LintCount++;
                    if (LintCount == 1) continue; // skip the first one, because it is our ID
                    if (string.IsNullOrEmpty(LstrLine)) continue;
                    string[] LstrParts = LstrLine.Split(',');
                    MobjData.Add(new RibbonRemovalData(
                                        (RibbonRemovalData.ItemTypesEnum)Enum.Parse(typeof(RibbonRemovalData.ItemTypesEnum), LstrParts[0].Trim()),
                                        LstrParts[1].Trim(),
                                        LstrParts[2].Trim()));
                }
            }
            catch (Exception PobjEx)
            {
                throw new Exception("Loading configuration failed: " + PobjEx.Message);
            }
        }

        /// <summary>
        /// Loads the namespaces into memory
        /// </summary>
        private void LoadNamespaceInfo()
        {
            int iCount = 0; // for namespaces
                            // build and create a list of namespaces for each account loaded in Outlook
                            // what we will do is create an entry for every account since we do not know if
                            // which one the add-in is loaded in. It could be one, or all
            MobjAccountUIDs = new Dictionary<string, string>();
            // loop through each account and add a ribbon NS element for each one
            foreach (Outlook.Account account in Globals.ThisAddIn.Application.Session.Accounts)
            {
                iCount++;
                Outlook.PropertyAccessor propertyAccessor = account.CurrentUser.PropertyAccessor;
                string ns = "x" + iCount.ToString();
                MobjAccountUIDs.Add(ns, MstrWebAddinID + "_" + propertyAccessor.BinaryToString(propertyAccessor.GetProperty(PR_EMSMDB_SECTION_UID)));
                break;
            }
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            try
            {
                this.Ribbon = ribbonUI;
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
        public string TabSet { get; private set; }

        public RibbonRemovalData(ItemTypesEnum PobjType, string PstrTabName, string PstrGroupName)
        {
            ItemType = PobjType;
            GroupName = PstrGroupName;
            if (PstrTabName.Contains("\\"))
            {
                string[] LobjParts = PstrTabName.Split('\\');
                TabSet = LobjParts[0];
                TabName = LobjParts[1];
            }
            else
            {
                TabName = PstrTabName;
            }
        }
    }
}
