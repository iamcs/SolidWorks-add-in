using System;
using System.Runtime.InteropServices;
using System.Collections;
using System.Reflection;
using System.IO;

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using SolidWorks.Interop.swconst;
using SolidWorksTools;
using SolidWorksTools.File;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Configuration;
using System.Windows.Forms;

namespace SwCSharpAddinByStanley
{
    /// <summary>
    /// Summary description for SwCSharpAddinByStanley.
    /// </summary>
    [Guid("041bb175-3ce5-4124-b0bf-e31364f98f06"), ComVisible(true)]
    [SwAddin(
        Description = "FAST AND EASY",
        Title = "WATEREASY",
        LoadAtStartup = true
        )]
    public class SwAddin : ISwAddin
    {
        #region Local Variables
        ISldWorks iSwApp = null;
        ICommandManager iCmdMgr = null;
        int addinID = 0;
        BitmapHandler iBmp;

        public const int mainCmdGroupID = 5;
        public const int mainItemID1 = 0;
        public const int mainItemID2 = 1;
        public const int mainItemID3 = 2;
        public const int mainItemID4 = 3;
        public const int mainItemID5 = 4;
        public const int mainItemID6 = 5;
        public const int mainItemID7 = 6;
        public const int mainItemID8 = 7;
        public const int flyoutGroupID = 91;

        #region Event Handler Variables
        Hashtable openDocs = new Hashtable();
        SolidWorks.Interop.sldworks.SldWorks SwEventPtr = null;
        #endregion

        #region Property Manager Variables
        UserPMPage ppage = null;        
        #endregion


        // Public Properties
        public ISldWorks SwApp
        {
            get { return iSwApp; }
        }
        public ICommandManager CmdMgr
        {
            get { return iCmdMgr; }
        }

        public Hashtable OpenDocs
        {
            get { return openDocs; }
        }

        #endregion

        #region SolidWorks Registration
        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type t)
        {
            #region Get Custom Attribute: SwAddinAttribute
            SwAddinAttribute SWattr = null;
            Type type = typeof(SwAddin);

            foreach (System.Attribute attr in type.GetCustomAttributes(false))
            {
                if (attr is SwAddinAttribute)
                {
                    SWattr = attr as SwAddinAttribute;
                    break;
                }
            }

            #endregion

            try
            {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\AddIns\\{" + t.GUID.ToString() + "}";
                Microsoft.Win32.RegistryKey addinkey = hklm.CreateSubKey(keyname);
                addinkey.SetValue(null, 0);

                addinkey.SetValue("Description", SWattr.Description);
                addinkey.SetValue("Title", SWattr.Title);

                keyname = "SOFTWARE\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                addinkey = hkcu.CreateSubKey(keyname);
                addinkey.SetValue(null, Convert.ToInt32(SWattr.LoadAtStartup), Microsoft.Win32.RegistryValueKind.DWord);
            }
            catch (System.NullReferenceException nl)
            {
                Console.WriteLine("There was a problem registering this dll: SWattr is null. \n\"" + nl.Message + "\"");
                System.Windows.Forms.MessageBox.Show("There was a problem registering this dll: SWattr is null.\n\"" + nl.Message + "\"");
            }

            catch (System.Exception e)
            {
                Console.WriteLine(e.Message);

                System.Windows.Forms.MessageBox.Show("There was a problem registering the function: \n\"" + e.Message + "\"");
            }
        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type t)
        {
            try
            {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\AddIns\\{" + t.GUID.ToString() + "}";
                hklm.DeleteSubKey(keyname);

                keyname = "SOFTWARE\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                hkcu.DeleteSubKey(keyname);
            }
            catch (System.NullReferenceException nl)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + nl.Message);
                System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + nl.Message + "\"");
            }
            catch (System.Exception e)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + e.Message);
                System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + e.Message + "\"");
            }
        }

        #endregion

        #region ISwAddin Implementation
        public SwAddin()
        {
        }

        public bool ConnectToSW(object ThisSW, int cookie)
        {
            iSwApp = (ISldWorks)ThisSW;
            addinID = cookie;

            //Setup callbacks
            iSwApp.SetAddinCallbackInfo(0, this, addinID);

            #region Setup the Command Manager
            iCmdMgr = iSwApp.GetCommandManager(cookie);
            AddCommandMgr();
            #endregion

            #region Setup the Event Handlers
            SwEventPtr = (SolidWorks.Interop.sldworks.SldWorks)iSwApp;
            openDocs = new Hashtable();
            AttachEventHandlers();
            #endregion

            #region Setup Sample Property Manager
            AddPMP();
            #endregion

            return true;
        }

        public bool DisconnectFromSW()
        {
            RemoveCommandMgr();
            RemovePMP();
            DetachEventHandlers();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(iCmdMgr);
            iCmdMgr = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(iSwApp);
            iSwApp = null;
            //The addin _must_ call GC.Collect() here in order to retrieve all managed code pointers 
            GC.Collect();
            GC.WaitForPendingFinalizers();

            GC.Collect();
            GC.WaitForPendingFinalizers();

            return true;
        }
        #endregion

        #region UI Methods
        public void AddCommandMgr()
        {
            ICommandGroup cmdGroup;
            if (iBmp == null)
                iBmp = new BitmapHandler();
            Assembly thisAssembly;
            int cmdIndex0, cmdIndex1, cmdIndex2, cmdIndex3, cmdIndex4, cmdIndex5, cmdIndex6,cmdIndex7;
            string Title = "WATEASY", ToolTip = "fast and easy";


            int[] docTypes = new int[]{(int)swDocumentTypes_e.swDocASSEMBLY,
                                       (int)swDocumentTypes_e.swDocDRAWING,
                                       (int)swDocumentTypes_e.swDocPART};

            thisAssembly = System.Reflection.Assembly.GetAssembly(this.GetType());


            int cmdGroupErr = 0;
            bool ignorePrevious = false;

            object registryIDs;
            //get the ID information stored in the registry
            bool getDataResult = iCmdMgr.GetGroupDataFromRegistry(mainCmdGroupID, out registryIDs);

            int[] knownIDs = new int[2] { mainItemID1, mainItemID2 };

            if (getDataResult)
            {
                if (!CompareIDs((int[])registryIDs, knownIDs)) //if the IDs don't match, reset the commandGroup
                {
                    ignorePrevious = true;
                }
            }

            cmdGroup = iCmdMgr.CreateCommandGroup2(mainCmdGroupID, Title, ToolTip, "", -1, ignorePrevious, ref cmdGroupErr);
            cmdGroup.LargeIconList = iBmp.CreateFileFromResourceBitmap("SwCSharpAddinByStanley.ToolbarLarge.bmp", thisAssembly);
            cmdGroup.SmallIconList = iBmp.CreateFileFromResourceBitmap("SwCSharpAddinByStanley.ToolbarSmall.bmp", thisAssembly);
            cmdGroup.LargeMainIcon = iBmp.CreateFileFromResourceBitmap("SwCSharpAddinByStanley.MainIconLarge.bmp", thisAssembly);
            cmdGroup.SmallMainIcon = iBmp.CreateFileFromResourceBitmap("SwCSharpAddinByStanley.MainIconSmall.bmp", thisAssembly);

            int menuToolbarOption = (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem);
            cmdIndex0 = cmdGroup.AddCommandItem2("清空属性", -1, "清空属性", "清空属性", 0, "PropertyClear", "", mainItemID1, menuToolbarOption);
            //cmdIndex0 = cmdGroup.AddCommandItem2("CreateCube", -1, "Create a cube", "Create cube", 0, "CreateCube", "", mainItemID1, menuToolbarOption);
            //cmdIndex1 = cmdGroup.AddCommandItem2("Show PMP", -1, "Display sample property manager", "Show PMP", 2, "ShowPMP", "EnablePMP", mainItemID2, menuToolbarOption);
            cmdIndex1 = cmdGroup.AddCommandItem2("填写图号", -1, "填写图号", "填写图号", 1, "DrawingNoInput", "", mainItemID2, menuToolbarOption);
            cmdIndex2 = cmdGroup.AddCommandItem2("配置插件", -1, "配置插件", "配置插件", 2, "ConfigModify", "", mainItemID3, menuToolbarOption);
            cmdIndex3 = cmdGroup.AddCommandItem2("随机颜色", -1, "随机颜色", "随机颜色", 3, "PartDye", "", mainItemID4, menuToolbarOption);
            cmdIndex4 = cmdGroup.AddCommandItem2("检查配孔", -1, "检查配孔", "检查配孔", 4, "HoleCheck", "", mainItemID5, menuToolbarOption);
            cmdIndex5 = cmdGroup.AddCommandItem2("外包尺寸", -1, "外包尺寸", "外包尺寸", 5, "GetBoundingBox", "", mainItemID6, menuToolbarOption);
            cmdIndex6 = cmdGroup.AddCommandItem2("关联打开", -1, "关联打开", "关联打开", 7, "OpenRelvantFile", "", mainItemID7, menuToolbarOption);
            cmdIndex7 = cmdGroup.AddCommandItem2("关于我", -1, "关于我", "关于我", 6, "SaveAs", "", mainItemID8, menuToolbarOption);

            cmdGroup.HasToolbar = true;
            cmdGroup.HasMenu = true;
            cmdGroup.Activate();

            bool bResult;



            FlyoutGroup flyGroup = iCmdMgr.CreateFlyoutGroup(flyoutGroupID, "Dynamic Flyout", "Flyout Tooltip", "Flyout Hint",
              cmdGroup.SmallMainIcon, cmdGroup.LargeMainIcon, cmdGroup.SmallIconList, cmdGroup.LargeIconList, "FlyoutCallback", "FlyoutEnable");


            flyGroup.AddCommandItem("FlyoutCommand 1", "test", 0, "FlyoutCommandItem1", "FlyoutEnableCommandItem1");

            flyGroup.FlyoutType = (int)swCommandFlyoutStyle_e.swCommandFlyoutStyle_Simple;


            foreach (int type in docTypes)
            {
                CommandTab cmdTab;

                cmdTab = iCmdMgr.GetCommandTab(type, Title);

                if (cmdTab != null & !getDataResult | ignorePrevious)//if tab exists, but we have ignored the registry info (or changed command group ID), re-create the tab.  Otherwise the ids won't matchup and the tab will be blank
                {
                    bool res = iCmdMgr.RemoveCommandTab(cmdTab);
                    cmdTab = null;
                }

                //if cmdTab is null, must be first load (possibly after reset), add the commands to the tabs
                //TODO: 将命令添加到tab上
                if (cmdTab == null)
                {
                    cmdTab = iCmdMgr.AddCommandTab(type, Title);

                    CommandTabBox cmdBox = cmdTab.AddCommandTabBox();

                    int[] cmdIDs = new int[8];
                    int[] TextType = new int[8];

                    cmdIDs[0] = cmdGroup.get_CommandID(cmdIndex0);

                    TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;

                    cmdIDs[1] = cmdGroup.get_CommandID(cmdIndex1);

                    TextType[1] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;
                    cmdIDs[2] = cmdGroup.get_CommandID(cmdIndex2);

                    TextType[2] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;
                    cmdIDs[3] = cmdGroup.get_CommandID(cmdIndex3);
                    
                    TextType[3] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;
                    cmdIDs[4] = cmdGroup.get_CommandID(cmdIndex4);

                    TextType[4] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;
                    cmdIDs[5] = cmdGroup.get_CommandID(cmdIndex5);

                    TextType[5] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;
                    cmdIDs[6] = cmdGroup.get_CommandID(cmdIndex6);

                    TextType[6] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;
                    cmdIDs[7] = cmdGroup.get_CommandID(cmdIndex7);

                    TextType[7] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;
                    
                    //cmdIDs[2] = cmdGroup.ToolbarId;

                    //TextType[2] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal | (int)swCommandTabButtonFlyoutStyle_e.swCommandTabButton_ActionFlyout;

                    bResult = cmdBox.AddCommands(cmdIDs, TextType);



                    CommandTabBox cmdBox1 = cmdTab.AddCommandTabBox();
                    cmdIDs = new int[1];
                    TextType = new int[1];

                    cmdIDs[0] = flyGroup.CmdID;
                    TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow | (int)swCommandTabButtonFlyoutStyle_e.swCommandTabButton_ActionFlyout;

                    bResult = cmdBox1.AddCommands(cmdIDs, TextType);

                    cmdTab.AddSeparator(cmdBox1, cmdIDs[0]);

                }

            }
            thisAssembly = null;

        }

        public void RemoveCommandMgr()
        {
            iBmp.Dispose();

            iCmdMgr.RemoveCommandGroup(mainCmdGroupID);
            iCmdMgr.RemoveFlyoutGroup(flyoutGroupID);
        }

        public bool CompareIDs(int[] storedIDs, int[] addinIDs)
        {
            List<int> storedList = new List<int>(storedIDs);
            List<int> addinList = new List<int>(addinIDs);

            addinList.Sort();
            storedList.Sort();

            if (addinList.Count != storedList.Count)
            {
                return false;
            }
            else
            {

                for (int i = 0; i < addinList.Count; i++)
                {
                    if (addinList[i] != storedList[i])
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        public Boolean AddPMP()
        {
            ppage = new UserPMPage(this);
            return true;
        }

        public Boolean RemovePMP()
        {
            ppage = null;
            return true;
        }

        #endregion

        #region UI Callbacks
        public void CreateCube()
        {
            //make sure we have a part open
            string partTemplate = iSwApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            if ((partTemplate != null) && (partTemplate != ""))
            {
                IModelDoc2 modDoc = (IModelDoc2)iSwApp.NewDocument(partTemplate, (int)swDwgPaperSizes_e.swDwgPaperA2size, 0.0, 0.0);

                modDoc.InsertSketch2(true);
                modDoc.SketchRectangle(0, 0, 0, .1, .1, .1, false);
                //Extrude the sketch
                IFeatureManager featMan = modDoc.FeatureManager;
                featMan.FeatureExtrusion(true,
                    false, false,
                    (int)swEndConditions_e.swEndCondBlind, (int)swEndConditions_e.swEndCondBlind,
                    0.1, 0.0,
                    false, false,
                    false, false,
                    0.0, 0.0,
                    false, false,
                    false, false,
                    true,
                    false, false);
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("There is no part template available. Please check your options and make sure there is a part template selected, or select a new part template.");
            }
        }

        public void PropertyClear()
        {
            IModelDoc2 modDoc = (IModelDoc2)iSwApp.ActiveDoc;
            string[] InfoNames;
            string[] InfoNames2;
            string[] ConfigNames;

            InfoNames = modDoc.GetCustomInfoNames();
            foreach (string x in InfoNames)
            {
                modDoc.DeleteCustomInfo(x);
            }
            ConfigNames = modDoc.GetConfigurationNames();
            foreach (string x in ConfigNames)
            {
                InfoNames2 = modDoc.GetCustomInfoNames2(x);                
                foreach (string y in InfoNames2)
                {                    
                    modDoc.DeleteCustomInfo2(x,y);
                }
            } 
        }

        public void DrawingNoInput()
        {
            IModelDoc2 modDoc = (IModelDoc2)iSwApp.ActiveDoc;                       
            
            //获取当前打开文件类型：1-part，2-assembly
            int modleType = modDoc.GetType();
            //将本体命名
            GetAndSetDrawingNo(modDoc);//待完善

            //装配体则对下层命名
            if (modleType == 2)
            {
                AssemblyDoc asmDoc = (AssemblyDoc)modDoc;
                SelectionMgr selectionMgr = modDoc.SelectionManager;
                IModelDoc2 partDoc;
                Component2 component;
                //asmDoc.getsel
                int looptime = selectionMgr.GetSelectedObjectCount2(-1);
                while (looptime >= 1)
                {
                    //todo：校核是否是零件类,现在是直接抛出    
                    try
                    {
                        looptime--;
                        component = selectionMgr.GetSelectedObject6(looptime+1, -1);
                        partDoc = (IModelDoc2)component.GetModelDoc2();
                        GetAndSetDrawingNo(partDoc);                                            
                    }
                    catch
                    { }
                }
            }            
        }

        private static void GetAndSetDrawingNo(IModelDoc2 modDoc)
        {
            string fileName;
            string pattern;
            string drawingNoEnd;
            string drawingNoPre;
            fileName = modDoc.GetTitle();

            //去除后缀
            fileName = fileName.Substring(0, fileName.LastIndexOf("."));
            //正则提取5位图号
            pattern = "((a|A)?[0-9]{3,5})(a|A)?$";
            drawingNoEnd = Regex.Match(fileName, pattern).Value;
            //从配置文件获取前缀
            drawingNoPre = ConfigurationManipulate.GetConfigValue("图号前缀");
            //将提取的5位图号写入自定义属性
            if (drawingNoEnd != "")
            {
                modDoc.AddCustomInfo("图号", "文字", drawingNoPre + drawingNoEnd);
                modDoc.set_CustomInfo("图号", drawingNoPre + drawingNoEnd);
            }
            else
            {
                modDoc.AddCustomInfo("图号", "文字", "外购件");
                modDoc.set_CustomInfo("图号", "外购件");
            }
            
        }


                
        public void ConfigModify()
        {
            Form1 form = new Form1();
            form.Show();
        }

        public void PartDye()
        {
            IModelDoc2 modDoc = (IModelDoc2)iSwApp.ActiveDoc;
            //获取当前打开文件类型：1-part，2-assembly
            int modleType = modDoc.GetType();
            dynamic materialPropertyValues;
            PartDoc partDoc;
            Component2 component;

            if (modleType == 1)
            {
                partDoc = (PartDoc)modDoc;
                materialPropertyValues = partDoc.MaterialPropertyValues;

                partDoc.MaterialPropertyValues = RandomColorValues(materialPropertyValues);                
            }
            else if(modleType == 2)
            {
                AssemblyDoc asmDoc = (AssemblyDoc)modDoc;
                SelectionMgr selectionMgr = modDoc.SelectionManager;
                //asmDoc.getsel
                int looptime = selectionMgr.GetSelectedObjectCount2(-1);
                while(looptime>=1)
                {
                    try
                    {
                        looptime--;
                        component = selectionMgr.GetSelectedObject6(looptime+1, -1);
                        partDoc = (PartDoc)component.GetModelDoc2();                        
                        //todo：校核是否是零件类,现在是直接抛出
                        materialPropertyValues = partDoc.MaterialPropertyValues;
                        partDoc.MaterialPropertyValues = RandomColorValues(materialPropertyValues);
                    }
                    catch
                    { }
                }
            }
            modDoc.GraphicsRedraw2();
        }

        public dynamic RandomColorValues(dynamic materialPropertyValues)
        {
            Random random = new Random();
            float red, green, blue;
            string dyeMode;

            red = ((float)random.Next(0, 256)) / 256;
            green = ((float)random.Next(0, 256)) / 256;
            blue = ((float)random.Next(0, 256)) / 256;
            dyeMode = ConfigurationManipulate.GetConfigValue("随机颜色");
            if (dyeMode == "False")
            {
                materialPropertyValues[0] = red;
                materialPropertyValues[1] = green;
                materialPropertyValues[2] = blue;
            }
            else if (dyeMode == "True")
            {
                materialPropertyValues[0] = 1;
                materialPropertyValues[1] = 1;
                materialPropertyValues[2] = 1;
            }
            

            return materialPropertyValues;
        }

        public void HoleCheck()
        {
            //MessageBox.Show("此功能计划中...");
            IModelDoc2 modDoc = (IModelDoc2)iSwApp.ActiveDoc;
            //获取当前打开文件类型：1-part，2-assembly
            int modleType = modDoc.GetType();
            dynamic materialPropertyValues;
            AssemblyDoc asmDoc;
            PartDoc partDoc;
            Component2 component;
            WizardHoleFeatureData2 holeFeatureData;
            Feature feature;
            
            int featureCount;
            ModelDocExtension modDocExtension;
            object[] arrBody = null;
            Body2 swBody = default(Body2);
            Face2[] arrface;
            Edge[] edges;


            
            partDoc = (PartDoc)modDoc;
            feature = partDoc.FirstFeature();
            while(feature != null)
            {
                if(feature.GetTypeName() == "HoleWzd")
                {
                    holeFeatureData = feature.GetDefinition();
                    foreach(SketchPoint x in holeFeatureData.GetSketchPoints())
                    {
                        MessageBox.Show("发现孔坐标\nx: "+(x.X*1000).ToString()+"\n"+
                                        "y: " + (x.Y * 1000).ToString() + "\n" +
                                        "z: " + (x.Z * 1000).ToString() + "\n");
                    }
                };
                feature= feature.GetNextFeature();
            }
            /*WizardHoleFeatureData2 holes;
            modDoc.feature
            partDoc.FeatureById
            arrBody = (Body2[])partDoc.GetBodies2((int)swBodyType_e.swSolidBody, true);
            foreach (Body2 x in arrBody)
            {
                arrface = (Face2[])x.GetFaces();
                foreach (Face2 y in arrface)
                {
                    y.is
                    edges = (Edge[])y.GetEdges();
                    foreach (Edge z in edges)
                    { 
                      
                    }
                }
            }
            partDoc.GetNamedEntities();
            featureCount = modDoc.GetFeatureCount();
            while(featureCount>=0)
            {
                featureCount--;                
                //feature = partDoc.;
            }*/
        }

        public void GetBoundingBox()
        {
            IModelDoc2 modDoc = (IModelDoc2)iSwApp.ActiveDoc;
            //获取当前打开文件类型：1-part，2-assembly
            int modleType = modDoc.GetType();
            AssemblyDoc asmDoc;
            PartDoc partDoc;            
            double[] boundingBox = null;            

            switch (modleType)
            { 
                case 1:
                    partDoc = (PartDoc)modDoc;
                    boundingBox = partDoc.GetPartBox(false);
                    break;
                case 2:
                    asmDoc = (AssemblyDoc)modDoc;
                    boundingBox = asmDoc.GetBox(1);
                    int i = 0;
                    foreach (double x in boundingBox)
                    {
                        boundingBox[i] = x * 1000;
                        i++;
                    }
                    break;
            }
            if(boundingBox != null)
            {
                MessageBox.Show("外形尺寸为：\n" +
                    "X方向: " + (boundingBox[3] - boundingBox[0]).ToString("f2") + "\n" +
                    "Y方向: " + (boundingBox[4] - boundingBox[1]).ToString("f2") + "\n" +
                    "Z方向: " + (boundingBox[5] - boundingBox[2]).ToString("f2"));
            };
            
        }

        public void OpenRelvantFile()
        {
            IModelDoc2 modDoc = (IModelDoc2)iSwApp.ActiveDoc;
            //获取当前打开文件类型：1-part，2-assembly, 3-drawing
            int modleType = modDoc.GetType();
            string fileName;
            string targetName;
            string readyName;
            string path;
            string[] siblingNames;
            int errors = 0;
            int warnings = 0;

            path = modDoc.GetPathName().Substring(0, modDoc.GetPathName().LastIndexOf("\\"));
            fileName = modDoc.GetPathName().Substring(modDoc.GetPathName().LastIndexOf("\\") + 1, modDoc.GetPathName().LastIndexOf(".") - modDoc.GetPathName().LastIndexOf("\\") - 1);
            siblingNames = Directory.GetFiles(path,"*.sld");

            foreach (string x in siblingNames)
            {
                //获取不包括路径的文件名
                readyName = x.ToLower().Replace("~$", "").Substring(x.LastIndexOf("\\") + 1, x.LastIndexOf(".") - x.LastIndexOf("\\") - 1);
                if (x.Replace("~$", "") != modDoc.GetPathName() & (readyName.LastIndexOf(fileName.ToLower()) != -1 || fileName.LastIndexOf(readyName.ToLower()) != -1))
                {
                    targetName = x.Replace("~$","");
                    iSwApp.OpenDoc6(targetName, 1, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);
                    iSwApp.OpenDoc6(targetName, 2, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);
                    iSwApp.OpenDoc6(targetName, 3, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);
                    iSwApp.ActivateDoc3(x.ToLower().Replace("~$", "").Substring(x.LastIndexOf("\\") + 1, x.Length - x.LastIndexOf("\\") - 1), false, (int)swRebuildOnActivation_e.swUserDecision, ref errors);                    
                }
            }            
        }

        public void SaveAs()
        {
            IModelDoc2 modDoc = (IModelDoc2)iSwApp.ActiveDoc;
            ModelDocExtension mDocExten;
            IModelDoc2 swModel;
            AssemblyDoc asmDoc = (AssemblyDoc)modDoc;
            SelectionMgr selectionMgr = modDoc.SelectionManager;            
            Component2 component;            
            string path;
            string recentPath;
            int errors = 0;
            int warnings = 0;

            recentPath = iSwApp.GetRecentFiles()[0].Substring(0, modDoc.GetPathName().LastIndexOf("\\"));
            //asmDoc.getsel
            int looptime = selectionMgr.GetSelectedObjectCount2(-1);
            while (looptime >= 1)
            {
                try
                {
                    looptime--;
                    component = selectionMgr.GetSelectedObject6(looptime + 1, -1);
                    swModel = (IModelDoc2)component.GetModelDoc2();
                    mDocExten = swModel.Extension;
                    //todo：校核是否是零件类,现在是直接抛出
                    //另存为对话框
                    SaveFileDialog sfd = new SaveFileDialog();
                    sfd.Title = "另存为";
                    sfd.InitialDirectory = recentPath;//更改为最近使用
                    sfd.Filter = "零件| *.sldprt";
                    sfd.ShowDialog();

                    path = sfd.FileName;
                    if (path == "")
                    {
                        return;
                    }
                    //另存为新零件
                    mDocExten.SaveAs(path, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, ref errors, ref warnings);
                    //替换现有零件
                    asmDoc.ReplaceComponents(path,"",false,true);                    
                }
                catch
                { }
            }
            
        }

        public void AboutMe()
        {
            AboutMe form = new AboutMe();
            form.Show();
        }

        public void ShowPMP()
        {
            if (ppage != null)
                ppage.Show();
        }

        public int EnablePMP()
        {
            if (iSwApp.ActiveDoc != null)
                return 1;
            else
                return 0;
        }

        public void FlyoutCallback()
        {
            FlyoutGroup flyGroup = iCmdMgr.GetFlyoutGroup(flyoutGroupID);
            flyGroup.RemoveAllCommandItems();

            flyGroup.AddCommandItem(System.DateTime.Now.ToLongTimeString(), "test", 0, "FlyoutCommandItem1", "FlyoutEnableCommandItem1");
            
        }
        public int FlyoutEnable()
        {
            return 1;
        }

        public void FlyoutCommandItem1()
        {
            iSwApp.SendMsgToUser("Flyout command 1");
        }

        public int FlyoutEnableCommandItem1()
        {
            return 1;
        }
        #endregion

        #region Event Methods
        public bool AttachEventHandlers()
        {
            AttachSwEvents();
            //Listen for events on all currently open docs
            AttachEventsToAllDocuments();
            return true;
        }

        private bool AttachSwEvents()
        {
            try
            {
                SwEventPtr.ActiveDocChangeNotify += new DSldWorksEvents_ActiveDocChangeNotifyEventHandler(OnDocChange);
                SwEventPtr.DocumentLoadNotify2 += new DSldWorksEvents_DocumentLoadNotify2EventHandler(OnDocLoad);
                SwEventPtr.FileNewNotify2 += new DSldWorksEvents_FileNewNotify2EventHandler(OnFileNew);
                SwEventPtr.ActiveModelDocChangeNotify += new DSldWorksEvents_ActiveModelDocChangeNotifyEventHandler(OnModelChange);
                SwEventPtr.FileOpenPostNotify += new DSldWorksEvents_FileOpenPostNotifyEventHandler(FileOpenPostNotify);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
        }



        private bool DetachSwEvents()
        {
            try
            {
                SwEventPtr.ActiveDocChangeNotify -= new DSldWorksEvents_ActiveDocChangeNotifyEventHandler(OnDocChange);
                SwEventPtr.DocumentLoadNotify2 -= new DSldWorksEvents_DocumentLoadNotify2EventHandler(OnDocLoad);
                SwEventPtr.FileNewNotify2 -= new DSldWorksEvents_FileNewNotify2EventHandler(OnFileNew);
                SwEventPtr.ActiveModelDocChangeNotify -= new DSldWorksEvents_ActiveModelDocChangeNotifyEventHandler(OnModelChange);
                SwEventPtr.FileOpenPostNotify -= new DSldWorksEvents_FileOpenPostNotifyEventHandler(FileOpenPostNotify);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }

        }

        public void AttachEventsToAllDocuments()
        {
            ModelDoc2 modDoc = (ModelDoc2)iSwApp.GetFirstDocument();
            while (modDoc != null)
            {
                if (!openDocs.Contains(modDoc))
                {
                    AttachModelDocEventHandler(modDoc);
                }
                else if (openDocs.Contains(modDoc))
                {
                    bool connected = false;
                    DocumentEventHandler docHandler = (DocumentEventHandler)openDocs[modDoc];
                    if (docHandler != null)
                    {
                        connected = docHandler.ConnectModelViews();
                    }
                }

                modDoc = (ModelDoc2)modDoc.GetNext();
            }
        }

        public bool AttachModelDocEventHandler(ModelDoc2 modDoc)
        {
            if (modDoc == null)
                return false;

            DocumentEventHandler docHandler = null;

            if (!openDocs.Contains(modDoc))
            {
                switch (modDoc.GetType())
                {
                    case (int)swDocumentTypes_e.swDocPART:
                        {
                            docHandler = new PartEventHandler(modDoc, this);
                            break;
                        }
                    case (int)swDocumentTypes_e.swDocASSEMBLY:
                        {
                            docHandler = new AssemblyEventHandler(modDoc, this);
                            break;
                        }
                    case (int)swDocumentTypes_e.swDocDRAWING:
                        {
                            docHandler = new DrawingEventHandler(modDoc, this);
                            break;
                        }
                    default:
                        {
                            return false; //Unsupported document type
                        }
                }
                docHandler.AttachEventHandlers();
                openDocs.Add(modDoc, docHandler);
            }
            return true;
        }

        public bool DetachModelEventHandler(ModelDoc2 modDoc)
        {
            DocumentEventHandler docHandler;
            docHandler = (DocumentEventHandler)openDocs[modDoc];
            openDocs.Remove(modDoc);
            modDoc = null;
            docHandler = null;
            return true;
        }

        public bool DetachEventHandlers()
        {
            DetachSwEvents();

            //Close events on all currently open docs
            DocumentEventHandler docHandler;
            int numKeys = openDocs.Count;
            object[] keys = new Object[numKeys];

            //Remove all document event handlers
            openDocs.Keys.CopyTo(keys, 0);
            foreach (ModelDoc2 key in keys)
            {
                docHandler = (DocumentEventHandler)openDocs[key];
                docHandler.DetachEventHandlers(); //This also removes the pair from the hash
                docHandler = null;
            }
            return true;
        }
        #endregion

        #region Event Handlers
        //Events
        public int OnDocChange()
        {
            return 0;
        }

        public int OnDocLoad(string docTitle, string docPath)
        {
            return 0;
        }

        int FileOpenPostNotify(string FileName)
        {
            AttachEventsToAllDocuments();
            return 0;
        }

        public int OnFileNew(object newDoc, int docType, string templateName)
        {
            AttachEventsToAllDocuments();
            return 0;
        }

        public int OnModelChange()
        {
            return 0;
        }

        #endregion
    }

}
