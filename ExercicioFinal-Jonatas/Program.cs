using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;

namespace ExercicioFinal_Jonatas
{
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
                    //If you want to use an add-on identifier for the development license, you can specify an add-on identifier string as the second parameter.
                    //oApp = new Application(args[0], "XXXXX");
                    oApp = new Application(args[0]);
                }
                Menu MyMenu = new Menu();
                MyMenu.AddMenuItems();
                oApp.RegisterMenuEventHandler(MyMenu.SBO_Application_MenuEvent);
                Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);

                oApp.AfterInitialized += OApp_AfterInitialized;

                oApp.Run();

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        private static void OApp_AfterInitialized(object sender, EventArgs e)
        {
            try
            {
                string server = Application.SBO_Application.Company.ServerName.Replace("NDB@", "");

                if (server.IndexOf(':') != 0)
                {
                    string[] st = server.Split(':');

                    server = st[0];
                }

                string url = $@"https://{server}:50000/b1s/v1";

                //ServicePointManager.ServerCertificateValidationCallback += delegate { return true; };

                ServicePointManager.ServerCertificateValidationCallback += BypassSslCallback;

                Connection.Context(Application.SBO_Application.Company.GetServiceLayerConnectionContext(url), url,server);

                Application.SBO_Application.StatusBar.SetText("Add-on Jonatas - Exercício Final conectado com sucesso!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {

                throw ex;
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

        private static bool BypassSslCallback(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors)
        {
            return true;
        }
    }
}
