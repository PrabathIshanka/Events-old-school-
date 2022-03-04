using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM;

namespace TestEvent
{
    class Event
    {
        private SAPbouiCOM.Application Oapp;

        public _IApplicationEvents_ProgressBarEventEventHandler Oapp_ProgressBarEvent { get; private set; }

        private void setApplication()
        {
            SAPbouiCOM.SboGuiApi sboGuiApi = null;
            String sConnectionString = null;
            sboGuiApi = new SAPbouiCOM.SboGuiApi();
            sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));
            sboGuiApi.Connect(sConnectionString);
            Oapp = sboGuiApi.GetApplication(-1);
        }

        public Event()
        {
            setApplication();
            Oapp.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(Oapp_AppEvent);
            Oapp.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(Oapp_MenuEvent);
            Oapp.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(Oapp_ItemEvent);
            Oapp.StatusBarEvent += new SAPbouiCOM._IApplicationEvents_StatusBarEventEventHandler(Oapp_StatusBarEvent);
            Oapp.ProgressBarEvent += new SAPbouiCOM._IApplicationEvents_ProgressBarEventEventHandler(Oapp_ProgressBarEvent);

        }

        private void Oapp_StatusBarEvent(string Text, BoStatusBarMessageType messageType)
        {
            //throw new NotImplementedException();
            Oapp.MessageBox("Message :" +Text,1,"ok","","");
        }

        private void Oapp_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            //throw new NotImplementedException();
            BubbleEvent = true;
            if (pVal.FormType != 0)
            {
                SAPbouiCOM.BoEventTypes Eventnum = 0;
                Eventnum = pVal.EventType;
                if ((Eventnum != SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE) || (Eventnum != SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD))

                {
                    Oapp.MessageBox("Item event :" + Eventnum.ToString() + "on_" + pVal.FormUID, 1, "ok", "", "");
                }

            }
        }

        private void Oapp_MenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            // throw new NotImplementedException();
            BubbleEvent = true;
            if (pVal.BeforeAction == true)
            {
                Oapp.SetStatusBarMessage("Before Menu : " + pVal.MenuUID,SAPbouiCOM.BoMessageTime.bmt_Medium,false );
            }
            else 
            {
                Oapp.SetStatusBarMessage("After Menu : " + pVal.MenuUID, SAPbouiCOM.BoMessageTime.bmt_Medium, false); 
            }
        }

        private void Oapp_AppEvent(BoAppEventTypes EventType)
        {
           // throw new NotImplementedException();
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    Oapp.MessageBox("Company change event", 1, "ok", "", "");
                    break;

                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    Oapp.MessageBox("Language change event", 1,"ok", "", "");
                    break;

                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    Oapp.MessageBox("Shutdown event", 1, "ok", "", "");
                    break;

            }
        }
    }
}
