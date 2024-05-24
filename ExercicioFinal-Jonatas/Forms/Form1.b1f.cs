using ExercicioFinal_Jonatas.Models;
using Newtonsoft.Json;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace ExercicioFinal_Jonatas
{
    [FormAttribute("ExercicioFinal_Jonatas.Form1", "Forms/Form1.b1f")]
    class Form1 : UserFormBase
    {
        public Form1()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("stCodPN").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("edCodPN").Specific));
            this.EditText0.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.EditText0_ChooseFromListAfter);
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("stNomPN").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("edNomPN").Specific));
            this.EditText1.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.EditText1_ChooseFromListAfter);
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("stData01").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("edData01").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("stData02").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("edData02").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("stVend").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("cbVend").Specific));
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("stEnt01").Specific));
            this.EditText4 = ((SAPbouiCOM.EditText)(this.GetItem("edEnt01").Specific));
            this.EditText4.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.EditText4_ChooseFromListAfter);
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("stEnt02").Specific));
            this.EditText5 = ((SAPbouiCOM.EditText)(this.GetItem("edEnt02").Specific));
            this.EditText5.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.EditText5_ChooseFromListAfter);
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("mtxWiz").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("btList").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("btFaturar").Specific));
            this.Button2.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button2_PressedAfter);
            this.Button2.PressedBefore += new SAPbouiCOM._IButtonEvents_PressedBeforeEventHandler(this.Button2_PressedBefore);
            this.LinkedButton0 = ((SAPbouiCOM.LinkedButton)(this.GetItem("lbCodPN").Specific));
            this.LinkedButton1 = ((SAPbouiCOM.LinkedButton)(this.GetItem("lbEnt01").Specific));
            this.LinkedButton2 = ((SAPbouiCOM.LinkedButton)(this.GetItem("lbEnt02").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {

        }

        private SAPbouiCOM.StaticText StaticText0;

        private void OnCustomInitialize()
        {
            try
            {
                ((SAPbouiCOM.Matrix)this.UIAPIRawForm.Items.Item("mtxWiz").Specific).AutoResizeColumns();
                SAPbouiCOM.DBDataSource oslp = this.UIAPIRawForm.DataSources.DBDataSources.Item("OSLP");
                oslp.Query();

                for (int i = 0; i < ComboBox0.ValidValues.Count; i++)
                {
                    ComboBox0.ValidValues.Remove(i);
                }

                for (int i = 0; i < Matrix0.Columns.Item("SlpCode").ValidValues.Count; i++)
                {
                    Matrix0.Columns.Item("SlpCode").ValidValues.Remove(i);
                }

                for (int i = 0; i < oslp.Size; i++)
                {
                    ComboBox0.ValidValues.Add(oslp.GetValue("SlpCode", i), oslp.GetValue("SlpName", i));
                    Matrix0.Columns.Item("SlpCode").ValidValues.Add(oslp.GetValue("SlpCode", i), oslp.GetValue("SlpName", i));
                }

                SAPbouiCOM.Conditions cons1 = new SAPbouiCOM.Conditions();
                SAPbouiCOM.Condition con1 = cons1.Add();

                con1.Alias = "CardType";
                con1.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                con1.CondVal = "S";

                this.UIAPIRawForm.ChooseFromLists.Item("cflCodPN").SetConditions(cons1);
                this.UIAPIRawForm.ChooseFromLists.Item("cflNomPN").SetConditions(cons1);

                SAPbouiCOM.Conditions cons2 = new SAPbouiCOM.Conditions();
                SAPbouiCOM.Condition con2 = cons2.Add();

                con2.Alias = "DocStatus";
                con2.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                con2.CondVal = "O";

                this.UIAPIRawForm.ChooseFromLists.Item("cflEnt01").SetConditions(cons2);
                this.UIAPIRawForm.ChooseFromLists.Item("cflEnt02").SetConditions(cons2);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oslp);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(cons1);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(con1);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(cons2);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(con2);

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            catch (Exception)
            {

                throw;
            }

        }

        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.StaticText StaticText4;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.StaticText StaticText5;
        private SAPbouiCOM.EditText EditText4;
        private SAPbouiCOM.StaticText StaticText6;
        private SAPbouiCOM.EditText EditText5;
        private SAPbouiCOM.Matrix Matrix0;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Button Button2;
        private SAPbouiCOM.LinkedButton LinkedButton0;
        private SAPbouiCOM.LinkedButton LinkedButton1;
        private SAPbouiCOM.LinkedButton LinkedButton2;
        private bool Sucess;
        private int cnt;



        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.UserDataSource udCodPN = this.UIAPIRawForm.DataSources.UserDataSources.Item("udCodPN"),
                          udData01 = this.UIAPIRawForm.DataSources.UserDataSources.Item("udData01"),
                          udData02 = this.UIAPIRawForm.DataSources.UserDataSources.Item("udData02"),
                          udVend = this.UIAPIRawForm.DataSources.UserDataSources.Item("udVend"),
                          udEnt01 = this.UIAPIRawForm.DataSources.UserDataSources.Item("udEnt01"),
                          udEnt02 = this.UIAPIRawForm.DataSources.UserDataSources.Item("udEnt02");
                SAPbouiCOM.DBDataSource odln = this.UIAPIRawForm.DataSources.DBDataSources.Item("ODLN");
                SAPbouiCOM.Matrix mtx = (SAPbouiCOM.Matrix)this.UIAPIRawForm.Items.Item("mtxWiz").Specific;

                SAPbouiCOM.Conditions conditions = new SAPbouiCOM.Conditions();
                SAPbouiCOM.Condition condition = null;

                condition = conditions.Add();
                condition.Alias = "DocStatus";
                condition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                condition.CondVal = "O";


                if (!String.IsNullOrEmpty(udCodPN.ValueEx))
                {
                    condition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                    condition = conditions.Add();
                    condition.Alias = "CardCode";
                    condition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    condition.CondVal = udCodPN.ValueEx;
                }

                if (!String.IsNullOrEmpty(udData01.ValueEx))
                {
                    condition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                    condition = conditions.Add();
                    condition.Alias = "DocDate";
                    condition.Operation = SAPbouiCOM.BoConditionOperation.co_GRATER_EQUAL;
                    condition.CondVal = udData01.ValueEx;
                }

                if (!String.IsNullOrEmpty(udData02.ValueEx))
                {
                    condition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                    condition = conditions.Add();
                    condition.Alias = "DocDate";
                    condition.Operation = SAPbouiCOM.BoConditionOperation.co_LESS_EQUAL;
                    condition.CondVal = udData02.ValueEx;
                }

                if (!String.IsNullOrEmpty(udVend.ValueEx))
                {
                    condition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                    condition = conditions.Add();
                    condition.Alias = "SlpCode";
                    condition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    condition.CondVal = udVend.ValueEx;
                }

                if (!String.IsNullOrEmpty(udEnt01.ValueEx))
                {
                    condition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                    condition = conditions.Add();
                    condition.Alias = "DocEntry";
                    condition.Operation = SAPbouiCOM.BoConditionOperation.co_GRATER_EQUAL;
                    condition.CondVal = udEnt01.ValueEx;
                }

                if (!String.IsNullOrEmpty(udEnt02.ValueEx))
                {
                    condition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                    condition = conditions.Add();
                    condition.Alias = "DocEntry";
                    condition.Operation = SAPbouiCOM.BoConditionOperation.co_LESS_EQUAL;
                    condition.CondVal = udEnt02.ValueEx;
                }


                this.UIAPIRawForm.Freeze(true);

                odln.Clear();
                mtx.Clear();

                odln.Query(conditions);

                mtx.LoadFromDataSource();
                mtx.AutoResizeColumns();

                this.UIAPIRawForm.Freeze(false);

                List<object> list = new List<object>();

                list.Add(udCodPN);
                list.Add(udData01);
                list.Add(udData02);
                list.Add(udVend);
                list.Add(udEnt01);
                list.Add(udEnt02);
                list.Add(odln);
                list.Add(mtx);
                list.Add(conditions);
                list.Add(condition);

                foreach (object item in list)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(item);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            catch (Exception)
            {

                throw;
            }

        }

        private void EditText0_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg chooseEvent = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                SAPbouiCOM.DataTable dt = chooseEvent.SelectedObjects;

                if (dt != null)
                {
                    this.UIAPIRawForm.DataSources.UserDataSources.Item("udCodPN").ValueEx = dt.GetValue("CardCode", 0).ToString();
                    this.UIAPIRawForm.DataSources.UserDataSources.Item("udNomPN").ValueEx = dt.GetValue("CardName", 0).ToString();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(dt);
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(chooseEvent);

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

            }
            catch (Exception)
            {

                throw;
            }
        }

        private void EditText1_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg chooseEvent = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                SAPbouiCOM.DataTable dt = chooseEvent.SelectedObjects;

                if (dt != null)
                {
                    this.UIAPIRawForm.DataSources.UserDataSources.Item("udCodPN").ValueEx = dt.GetValue("CardCode", 0).ToString();
                    this.UIAPIRawForm.DataSources.UserDataSources.Item("udNomPN").ValueEx = dt.GetValue("CardName", 0).ToString();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(dt);
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(chooseEvent);
 
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

            }
            catch (Exception)
            {

                throw;
            }
        }

        private void EditText4_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg chooseEvent = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                SAPbouiCOM.DataTable dt = chooseEvent.SelectedObjects;

                if (dt != null)
                {
                    this.UIAPIRawForm.DataSources.UserDataSources.Item("udEnt01").ValueEx = dt.GetValue("DocEntry", 0).ToString();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(dt);
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(chooseEvent);
              
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

            }
            catch (Exception)
            {

                throw;
            }
        }

        private void EditText5_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg chooseEvent = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                SAPbouiCOM.DataTable dt = chooseEvent.SelectedObjects;

                if (dt != null)
                {
                    this.UIAPIRawForm.DataSources.UserDataSources.Item("udEnt02").ValueEx = dt.GetValue("DocEntry", 0).ToString();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(dt);
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(chooseEvent);
               
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

            }
            catch (Exception)
            {

                throw;
            }
        }

        private void Button2_PressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {

            try
            {

                Matrix0.FlushToDataSource();

                for (int i = 1; i <= Matrix0.RowCount; i++)
                {
                    if (((SAPbouiCOM.CheckBox)Matrix0.Columns.Item("#").Cells.Item(i).Specific).Checked)
                    {
                        cnt++;
                    }
                }

                if (cnt > 0)
                {
                    if (Application.SBO_Application.MessageBox("Essa ação dará início a criação das notas fiscais para as entregas que foram selecionadas! Deseja realmente continuar?", 2, "Sim", "Não") == 1)
                    {
                        BubbleEvent = true;
                    }
                    else
                    {
                        BubbleEvent = false;
                    }

                }
                else
                {
                    Application.SBO_Application.StatusBar.SetText("Antes de prosseguir selecione pelo menos uma entrega!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    BubbleEvent = false;
                }


                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            catch (Exception)
            {
                BubbleEvent = false;
                throw;
            }
        }

        private async void Button2_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ActionSuccess)
                {
                    SAPbouiCOM.ProgressBar progressBar = Application.SBO_Application.StatusBar.CreateProgressBar("Iniciango geração de notas ...", cnt, false);

                    int num = 0;

                    for (int i = 1; i <= Matrix0.RowCount; i++)
                    {                                           

                        if (((SAPbouiCOM.CheckBox)Matrix0.Columns.Item("#").Cells.Item(i).Specific).Checked)
                        {
                            int docEntry = int.Parse(((SAPbouiCOM.EditText)Matrix0.Columns.Item("DocEntry").Cells.Item(i).Specific).Value);

                            num++;

                            progressBar.Text = $"{num}/{cnt} - Processando entrega {docEntry} ...";

                            string content = await GetDelivery(docEntry);

                            if (Sucess)
                            {
                                string cont = await PostInvoice(content);

                                if (Sucess)
                                {
                                    progressBar.Text = $"{num}/{cnt} - Nota fiscal da entrega {docEntry} gerada com sucesso!";
                                    progressBar.Value++;
                                }
                                else
                                {
                                    ErrorSL er = JsonConvert.DeserializeObject<ErrorSL>(cont);
                                    progressBar.Text = $"{num}/{cnt} - Falha ao emitir a nota fiscal da entrega {docEntry}! Error: {er.error.code} - {er.error.message.value}";
                                    progressBar.Value++;
                                }

                            }
                            else
                            {
                                ErrorSL er = JsonConvert.DeserializeObject<ErrorSL>(content);
                                progressBar.Text = $"{num}/{cnt} - Falha ao obter os dados entrega {docEntry}! Error: {er.error.code} - {er.error.message.value}";
                                progressBar.Value++;
                            }

                        }
                    }

                    this.UIAPIRawForm.Items.Item("btList").Click();


                    System.Runtime.InteropServices.Marshal.ReleaseComObject(progressBar);

                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                }
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                Sucess = false;
                cnt = 0;
            }

        }


        private async Task<string> GetDelivery(int docentry)
        {
            try
            {
                string url = $@"{Connection.Url}/DeliveryNotes({docentry})";

                CookieContainer cookiecon = new CookieContainer();
                cookiecon.Add(new Cookie(Connection.CoockieName, Connection.SLToken, Connection.Path,Connection.ServerName));

                HttpClientHandler handler = new HttpClientHandler();
                handler.CookieContainer = cookiecon;

                using (HttpClient client = new HttpClient(handler))
                {
                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, url);

                    HttpResponseMessage response = await client.SendAsync(request);

                    string content = await response.Content.ReadAsStringAsync();

                    Sucess = response.IsSuccessStatusCode;

                    return content;

                }
            }
            catch (Exception)
            {

                throw;
            }

        }


        private async Task<string> PostInvoice(string json)
        {
            try
            {
                DeliveryNotes dn = JsonConvert.DeserializeObject<DeliveryNotes>(json);

                Invoices inv = new Invoices();

                inv.CardCode = dn.CardCode;
                inv.CardName = dn.CardName;
                inv.ContactPersonCode = dn.ContactPersonCode;
                inv.NumAtCard = dn.NumAtCard;
                inv.DocCurrency = dn.DocCurrency;
                inv.DocRate = dn.DocRate;
                inv.BPL_IDAssignedToInvoice = dn.BPL_IDAssignedToInvoice;
                inv.DocDate = DateTime.Now.ToString();
                inv.TaxDate = DateTime.Now.ToString();
                inv.SalesPersonCode = dn.SalesPersonCode;
                inv.DocumentsOwner = dn.DocumentsOwner;
                inv.Reference2 = dn.Reference2;
                inv.Comments = dn.Comments + $" Add-on Jonatas - Minha nota vinculada a entrega número: {dn.DocNum}";
                inv.ShipToCode = dn.ShipToCode;
                inv.PayToCode = dn.PayToCode;
                inv.TransportationCode = dn.TransportationCode;
                inv.AgentCode = dn.AgentCode;
                inv.LanguageCode = dn.LanguageCode;
                inv.TrackingNumber = dn.TrackingNumber;
                inv.BPChannelCode = dn.BPChannelCode;
                inv.BPChannelContact = dn.BPChannelContact;
                inv.ControlAccount = dn.ControlAccount;
                inv.PaymentGroupCode = dn.PaymentGroupCode;
                inv.PaymentMethod = dn.PaymentMethod;
                inv.AttachmentEntry = dn.AttachmentEntry;
                inv.Project = dn.Project;
                inv.Indicator = dn.Indicator;
                inv.ImportFileNum = dn.ImportFileNum;
                inv.SequenceCode = dn.SequenceCode;

                foreach (DocumentLine line in dn.DocumentLines)
                {
                    DocumentLine docline = new DocumentLine();

                    docline.ItemCode = line.ItemCode;
                    docline.ItemDescription = line.ItemDescription;
                    docline.Quantity = line.RemainingOpenInventoryQuantity;
                    docline.UnitPrice = line.UnitPrice;
                    docline.Usage = line.Usage;
                    docline.TaxCode = line.TaxCode;
                    docline.WarehouseCode = line.WarehouseCode;
                    docline.BaseLine = line.LineNum;
                    docline.BaseEntry = line.DocEntry;
                    docline.BaseType = "15";

                    inv.DocumentLines.Add(docline);
                }

                inv.TaxExtension = dn.TaxExtension;
                inv.TaxExtension.DocEntry = null;
                inv.AddressExtension = dn.AddressExtension;
                inv.AddressExtension.DocEntry = null;


                string url = $@"{Connection.Url}/Invoices";

                CookieContainer cookiecon = new CookieContainer();
                cookiecon.Add(new Cookie(Connection.CoockieName, Connection.SLToken, Connection.Path, Connection.ServerName));

                HttpClientHandler handler = new HttpClientHandler();
                handler.CookieContainer = cookiecon;

                using (HttpClient client = new HttpClient(handler))
                {
                    client.DefaultRequestHeaders.ExpectContinue = false;

                    JsonSerializerSettings settings = new JsonSerializerSettings();
                    settings.NullValueHandling = NullValueHandling.Ignore;

                    string body = JsonConvert.SerializeObject(inv, Formatting.Indented, settings);

                    StringContent content = new StringContent(body, Encoding.UTF8, "application/json");

                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, url);
                    request.Content = content;

                    HttpResponseMessage response = await client.SendAsync(request);

                    string cont = await response.Content.ReadAsStringAsync();

                    Sucess = response.IsSuccessStatusCode;

                    return cont;

                }


            }
            catch (Exception)
            {

                throw;
            }
        }

    }
}