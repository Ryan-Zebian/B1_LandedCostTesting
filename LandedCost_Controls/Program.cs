using System;
using System.Collections.Generic;
using SAPbouiCOM.Framework;
using SAPbobsCOM;
using System.Xml;
using System.Xml.XPath;
using System.IO;


namespace LandedCost_Controls
{
    /// <summary>
    /// This Application Demo's how to Create a LandedCost and Link to Document ->A/P Invoice,GRPO, and LandedCost
    /// AddLanedCostXML - Copies and Links a new LandedCost through XML
    /// testLandedCost - Creates a new Landed Cost using the code required: BaseDocumentType,BaseEntry,LandedCost_costLines amount and code
    /// </summary>
    class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                Application oApp = null;
                if (args.Length < 1)
                {
                    oApp = new Application();
                }
                else
                {
                    oApp = new Application(args[0]);
                }
               
                Menu MyMenu = new Menu();
                MyMenu.AddMenuItems();
                oApp.RegisterMenuEventHandler(MyMenu.SBO_Application_MenuEvent);
                Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
               // openGRPO();
             //  updateLandedCost();
                //testLandedCost();
                AddLandedCostXML();
                oApp.Run();
            }
            catch (Exception ex)
            {
                //SAPbobsCOM.Company oCompany = Application.SBO_Application.Company.GetDICompany() as SAPbobsCOM.Company;
                //string s = oCompany.GetLastErrorDescription();
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        private static void AddLandedCostXML()
        {
            SAPbobsCOM.Company oCompany = Application.SBO_Application.Company.GetDICompany() as SAPbobsCOM.Company;
            SAPbobsCOM.CompanyService oCompanyService = oCompany.GetCompanyService() as SAPbobsCOM.CompanyService;
            LandedCostsService LandedService = oCompanyService.GetBusinessService(ServiceTypes.LandedCostsService) as LandedCostsService;
            LandedCost oLandedCost = LandedService.GetDataInterface(LandedCostsServiceDataInterfaces.lcsLandedCost) as LandedCost;
            LandedCostParams oCopiedLCParams = LandedService.GetDataInterface(LandedCostsServiceDataInterfaces.lcsLandedCostParams) as LandedCostParams;
            oCopiedLCParams.LandedCostNumber = 17;
            LandedCost oCopiedLC = LandedService.GetDataInterface(LandedCostsServiceDataInterfaces.lcsLandedCost) as LandedCost;
            oCopiedLC = LandedService.GetLandedCost(oCopiedLCParams);

           string k = XMLCleaner(oCopiedLC.ToXMLString(),oCopiedLC.LandedCostNumber+"",oCopiedLC.DocEntry+"");

           oLandedCost.FromXMLString(k);
           //Adding Costs
           LandedCost_CostLine oLandedCost_CostLine = oLandedCost.LandedCost_CostLines.Add();

           oLandedCost_CostLine.LandedCostCode = "TI";
           oLandedCost_CostLine.amount = 10;

           
            LandedService.AddLandedCost(oLandedCost);
        

        }

        private static void updateLandedCost()
        {
            SAPbobsCOM.Company oCompany = Application.SBO_Application.Company.GetDICompany() as SAPbobsCOM.Company;
            SAPbobsCOM.CompanyService oCompanyService = oCompany.GetCompanyService() as SAPbobsCOM.CompanyService;
            LandedCostsService LandedService = oCompanyService.GetBusinessService(ServiceTypes.LandedCostsService) as LandedCostsService;
            
            LandedCostParams oLandedCostUpdateParams = LandedService.GetDataInterface(LandedCostsServiceDataInterfaces.lcsLandedCostParams) as LandedCostParams;
            LandedCost oLandedCostUpdate = LandedService.GetDataInterface(LandedCostsServiceDataInterfaces.lcsLandedCost) as LandedCost;
            oLandedCostUpdateParams.LandedCostNumber = 10;
            //LandedService.CancelLandedCost(oLandedCostUpdateParams);
            oLandedCostUpdate = LandedService.GetLandedCost(oLandedCostUpdateParams);
            
            // oLandedCostUpdate.LandedCost_ItemLines.Remove(0);
            // LandedCost_ItemLine oLandedCost_ItemLine;
            //oLandedCost_ItemLine = oLandedCostUpdate.LandedCost_ItemLines.Add();
            //oLandedCost_ItemLine.BaseDocumentType = LandedCostBaseDocumentTypeEnum.asGoodsReceiptPO;
            //oLandedCost_ItemLine.BaseEntry = 521;
            //oLandedCost_ItemLine.BaseLine = 1;
            

            ////LandedCost_CostLine oLandedCostUpdate_CostLine = oLandedCostUpdate.LandedCost_CostLines.Add();
            ////oLandedCostUpdate_CostLine.LandedCostCode = "CM";
            ////oLandedCostUpdate_CostLine.amount = 11;
            //LandedService.UpdateLandedCost(oLandedCostUpdate);
            
        }

        private static void openGRPO()
        {
            Company oCompany = Application.SBO_Application.Company.GetDICompany() as Company;
            Documents oGRPODoc = oCompany.GetBusinessObject(BoObjectTypes.oPurchaseInvoices) as Documents;
            oGRPODoc.GetByKey(2474);
            oGRPODoc.OpenForLandedCosts = BoYesNoEnum.tYES;
            oGRPODoc.Update();
        }

        private static void testLandedCost()
        {
           SAPbobsCOM.Company oCompany = Application.SBO_Application.Company.GetDICompany() as SAPbobsCOM.Company;
           SAPbobsCOM.CompanyService oCompanyService = oCompany.GetCompanyService() as SAPbobsCOM.CompanyService;
           LandedCostsService LandedService = oCompanyService.GetBusinessService(ServiceTypes.LandedCostsService) as LandedCostsService;
           LandedCost oLandedCost = LandedService.GetDataInterface(LandedCostsServiceDataInterfaces.lcsLandedCost) as LandedCost;
           LandedCostParams oCopiedLCParams = LandedService.GetDataInterface(LandedCostsServiceDataInterfaces.lcsLandedCostParams) as LandedCostParams;
           oCopiedLCParams.LandedCostNumber = 13;
            LandedCost oCopiedLC =  LandedService.GetDataInterface(LandedCostsServiceDataInterfaces.lcsLandedCost) as LandedCost;
           oCopiedLC = LandedService.GetLandedCost(oCopiedLCParams);
            //Testing
            int total_lineNum = oCopiedLC.LandedCost_ItemLines.Count;
            long oLandedCostEntry = 0;//Try 10
            int GRPOEntry = 2474;

            //Adding Items
            for (int i = 0; i < 2; i++) { 
           LandedCost_ItemLine oLandedCost_ItemLine = oLandedCost.LandedCost_ItemLines.Add();
           oLandedCost_ItemLine.BaseDocumentType = LandedCostBaseDocumentTypeEnum.asPurchaseInvoice;
           oLandedCost_ItemLine.BaseEntry = GRPOEntry;
           oLandedCost_ItemLine.BaseLine = i;
           }
          
            //Adding Costs
           LandedCost_CostLine oLandedCost_CostLine = oLandedCost.LandedCost_CostLines.Add();
            
           oLandedCost_CostLine.LandedCostCode = "CD";
           oLandedCost_CostLine.amount = 99;
          //Extra Cost
            oLandedCost_CostLine = oLandedCost.LandedCost_CostLines.Add();
           oLandedCost_CostLine.LandedCostCode = "CM";
           oLandedCost_CostLine.amount = 150;



           LandedCostParams oLandedCostParams;
           oLandedCostParams = LandedService.AddLandedCost(oLandedCost);
           oLandedCostEntry = oLandedCostParams.LandedCostNumber;
           

        }

        private static string XMLCleaner(string xmlString,string baseDocNum, string baseEntry)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(xmlString);
            XPathNavigator xNav = xmlDoc.CreateNavigator();
            string query = "//DocEntry";
            xNav.MoveToRoot();
            
            XPathNodeIterator xPathIt = xNav.Select(query);
            CleanAttributes(xPathIt);
            query = "//LandedCostNumber";
            xPathIt = xNav.Select(query);
            CleanAttributes(xPathIt);
            query = "//Number";
            xPathIt = xNav.Select(query);
            CleanAttributes(xPathIt);
            //query = "//Reference";
            //xPathIt = xNav.Select(query);
           // CleanAttributes(xPathIt);
            query = "//LandedCost/LandedCost_ItemLines/LandedCost_ItemLine/Reference";
            xPathIt = xNav.Select(query);
            ReplaceAttributes(xPathIt,baseDocNum);
            query = "//LandedCost/LandedCost_ItemLines/LandedCost_ItemLine/BaseDocumentType";
            xPathIt = xNav.Select(query);
            ReplaceAttributes(xPathIt, "asLandedCosts");
            query = "//LandedCost/LandedCost_ItemLines/LandedCost_ItemLine/BaseEntry";
            xPathIt = xNav.Select(query);
            ReplaceAttributes(xPathIt, baseEntry);

            StringWriter stringWriter = new StringWriter();
            XmlTextWriter xmlTextWriter = new XmlTextWriter(stringWriter);
            xmlDoc.WriteTo(xmlTextWriter);
       
            xmlString = stringWriter.ToString();

            xmlTextWriter.Flush();
            stringWriter.Flush();
            return xmlString;
        }

        private static void ReplaceAttributes(XPathNodeIterator xPathIt, string baseDocNum)
        {
            string temp;
            if (xPathIt.Count > 0)
            {
                while (xPathIt.MoveNext())
                {
                    temp = xPathIt.Current.Value + " - " + xPathIt.CurrentPosition;
                    xPathIt.Current.SetValue(baseDocNum);
                    //xPathIt.Current.CreateAttribute(String.Empty, "nil", String.Empty, "true");
                }
            }
        }
        private static void CleanAttributes(XPathNodeIterator xPathIt)
        {
            string temp;
            if (xPathIt.Count > 0)
            {
                while (xPathIt.MoveNext())
                {
                    temp = xPathIt.Current.Value + " - " + xPathIt.CurrentPosition;
                    xPathIt.Current.SetValue("");
                    xPathIt.Current.CreateAttribute(String.Empty, "nil", String.Empty, "true");
                }
            }

        }

        static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //Exit Add-On
                    System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    break;
                default:
                    break;
            }
        }
    }
}
