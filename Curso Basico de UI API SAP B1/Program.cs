using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;
using SAPbouiCOM;

namespace Curso_Basico_de_UI_API_SAP_B1
{
    class Program
    {
        /*Agregar referencia ui api
          Agregar cadena (0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056)
          agregar referencia para el Application.run (System.Windows.Forms)
          Revisar la arquitectura del proyecto  sea la misma que la del cliente SAP
             */

        public static Application SBO_Application = null;
        public static SAPbobsCOM.Company oCompany = null;
        static void Main(string[] args)
        {
            ConexionUIAPI();
            ConexionSingleSignOn();
            SBO_Application.StatusBar.SetText("Addon iniciado Correctamente",BoMessageTime.bmt_Medium,BoStatusBarMessageType.smt_Success);
            // Cliclo para leer todos los eventos de tipo itemeven
            SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
            System.Windows.Forms.Application.Run();
        } 
        public static void ConexionUIAPI()
        {
            try
            {
                SboGuiApi oSboGuiApi = new SboGuiApi();
                string sConnStr = Environment.GetCommandLineArgs().GetValue(1).ToString();
                oSboGuiApi.Connect(sConnStr);

                SBO_Application = oSboGuiApi.GetApplication(-1);
                SBO_Application.StatusBar.SetText("EXITO - Conexion UI API Exitosa", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success);
                oSboGuiApi = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        //Opcion recomendada y con mejor performance
        public static void ConexionSingleSignOn()
        {
            try
            {
                oCompany = new SAPbobsCOM.Company();
                string sCookie = oCompany.GetContextCookie();
                string SConn = SBO_Application.Company.GetConnectionContext(sCookie);
                int error = oCompany.SetSboLoginContext(SConn);

                if (error == 0)
                {
                    oCompany.Connect();
                    SBO_Application.StatusBar.SetText("Exito - en la conexion DI API",BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success);
                }
                else
                {
                    SBO_Application.StatusBar.SetText("Error - en la conexion DI API", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
                }
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("Error General - " + ex.ToString(), BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
            }

        }
        //Conexion repartida para multiples Add-Ons
        public static void ConexionMultipleAddOn()
        {
            try
            {
                oCompany = SBO_Application.Company.GetDICompany();
                SBO_Application.StatusBar.SetText("Conecion DI API", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("Error General - " + ex.ToString(), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }
        }

        public static void SBO_Application_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                //Pedidos de Venta
                if (pVal.FormTypeEx == "139")
                {
                    if (pVal.EventType == BoEventTypes.et_FORM_LOAD && pVal.BeforeAction == false)
                    {
                        Form oForm = SBO_Application.Forms.Item(FormUID);
                        Item oItem;
                        Button oButton;
                        oItem = oForm.Items.Add("btnEntrega", BoFormItemTypes.it_BUTTON);
                        //Inicializando el objeto boton con la referencia del objeto item
                        oButton = oItem.Specific;
                        //Agregando propiedades al boton
                        oButton.Caption = "Entrega";
                        //agregando posicio del boton
                        oItem.Top = oForm.Height - (oItem.Height + 48);
                        oItem.Left = (oItem.Width + 20) + 60;
                    }

                    if (pVal.EventType == BoEventTypes.et_CLICK && pVal.BeforeAction == true && pVal.ItemUID == "btnEntrega")
                    {
                        //Entrega
                        Documents oNE = oCompany.GetBusinessObject(BoObjectTypes.oDeliveryNotes);
                        Form oForm = SBO_Application.Forms.Item(FormUID);
                        Recordset oRecordSet = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        Item oItem;
                        EditText oText;

                        oItem = oForm.Items.Item("8");
                        oText = oItem.Specific;


                        string cmd = "Select T0.CardCode,T1.DocEntry,T1.LineNum,T1.Dscription  From ORDR T0 Inner Join RDR1 T1 on T0.DocEntry=T1.DocEntry" +
                                     " where T0.DocNum='" + oText.Value + "'";

                        oRecordSet.DoQuery(cmd);

                        if (oRecordSet.RecordCount > 0)
                        {
                            oNE.CardCode = oRecordSet.Fields.Item("CardCode").Value;
                            oNE.DocType = BoDocumentTypes.dDocument_Service;
                            oNE.DocDate = DateTime.Now;
                            oNE.DocDueDate = DateTime.Now;
                            oNE.Comments = "Entrega creada con DI API desde un boton generado con UI API";
                            while (!oRecordSet.EoF)
                            {
                                oNE.Lines.BaseEntry = oRecordSet.Fields.Item("DocEntry").Value;
                                oNE.Lines.BaseLine = oRecordSet.Fields.Item("LineNum").Value;
                                oNE.Lines.BaseType = Convert.ToInt32(BoObjectTypes.oOrders);
                                oNE.Lines.ItemDescription = oRecordSet.Fields.Item("Dscription").Value;
                                oNE.Lines.Add();
                                oRecordSet.MoveNext();
                            }
                        }

                        oCompany.StartTransaction();
                        if (oNE.Add() == 0)
                        {
                            SBO_Application.StatusBar.SetText("Entrega creada: " + oCompany.GetNewObjectKey(),BoMessageTime.bmt_Medium,BoStatusBarMessageType.smt_Success);
                            oCompany.EndTransaction(BoWfTransOpt.wf_Commit);
                            cmd = "Select T0.DocNum From ODLN T0 where T0.DocEntry='" + Convert.ToString(oCompany.GetNewObjectKey()) + "'";

                            oRecordSet.DoQuery(cmd);

                            SBO_Application.ActivateMenuItem("2051");
                            Form form = SBO_Application.Forms.ActiveForm;
                            form.Mode = BoFormMode.fm_FIND_MODE;
                            ((EditText)form.Items.Item("8").Specific).Value = Convert.ToString(oRecordSet.Fields.Item("DocNum").Value);
                            form.Items.Item("1").Click(BoCellClickType.ct_Regular);
                        }
                        else
                        {
                            SBO_Application.StatusBar.SetText("Entrega Error: " + oCompany.GetLastErrorDescription(), BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
                            oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("Error General - " + ex.ToString(), BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);

                if (oCompany.InTransaction)
                {
                    oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                }
            }
        }
    }
}
