using System;
using SAPbouiCOM;


namespace OsiAddon
{
    public class AddonController
    {
        private GuiController guiController;

        private String lastError;

        private Boolean isAttached;

        public Boolean IsAttached // Acoplamento a interface do SAP B1 bem sucedida
        {
            get { return isAttached; }
        }


        public AddonController()
        {
            guiController = new GuiController();
            guiController.InitializeGui(AttachToSBO());
        }

        private SAPbouiCOM.Application AttachToSBO()
        {
            SAPbouiCOM.Application sboApplication = GetSBOApplication();
            if (sboApplication == null)
            {
                isAttached = false;
                return null;
            }

            sboApplication.AppEvent += new _IApplicationEvents_AppEventEventHandler(SboApplicationEvent);
            sboApplication.MenuEvent += new _IApplicationEvents_MenuEventEventHandler(SboMenuEvent);
            sboApplication.FormDataEvent += new _IApplicationEvents_FormDataEventEventHandler(SboFormDataEvent);
            sboApplication.ItemEvent += new _IApplicationEvents_ItemEventEventHandler(SboItemEvent);

            isAttached = true;
            return sboApplication;
        }

        private void SboApplicationEvent(BoAppEventTypes eventType)
        {
            // Monitora o evento de shut down da aplicação
            if (eventType == BoAppEventTypes.aet_ShutDown)
            {
                System.Windows.Forms.Application.Exit();
            }
        }

        private void SboMenuEvent(ref MenuEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            try
            {
                guiController.MenuEvent(ref pVal, out bubbleEvent);
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        private void SboFormDataEvent(ref BusinessObjectInfo businessObjectInfo, out bool bubbleEvent)
        {
            bubbleEvent = true;
            try
            {
                guiController.FormDataEvent(ref businessObjectInfo, out bubbleEvent);
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        private void SboItemEvent(String formUID, ref ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            try
            {
                guiController.ItemEvent(formUID, ref pVal, out bubbleEvent);
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        // Obtem o objeto de aplicação associado ao SAP Business One
        private SAPbouiCOM.Application GetSBOApplication()
        {
            SboGuiApi sboGuiApi;
            try
            {
                String[] args = Environment.GetCommandLineArgs();
                String connectionString = args[1];

                sboGuiApi = new SboGuiApi();
                sboGuiApi.Connect(connectionString);
            }
            catch
            {
                return null;
            }

            return sboGuiApi.GetApplication(-1);
        }
    }

}
