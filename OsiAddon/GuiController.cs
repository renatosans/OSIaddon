using System;
using System.IO;
using System.Threading;
using System.ComponentModel;
using System.Data.SqlClient;
using System.Collections.Generic;
using SAPbouiCOM;
using ClassLibrary;
using DataAccessObjects;
using DataTransferObjects;


namespace OsiAddon
{
    /// <summary>
    /// Classe responsável pelo controle da interface de usuário (graphical user interface)
    /// </summary>
    public class GuiController
    {
        private const int ADDRESS_TAB = 7;           // Identificador da aba "Endereços"

        private const int INSTALLATION_TAB = 10;     // Identificador da aba "Instalação"

        private const int SPECIFICATIONS_TAB = 11;   // Identificador da aba "Especificações"

        private const int BILLING_TAB = 15;          // Identificador da aba "Faturamento"

        private SAPbouiCOM.Application sboApplication;

        private SAPbobsCOM.Company sboCompany;

        private DataConnector dataConnector;

        private Dictionary<String, AddressDTO> addressDictionary = new Dictionary<String, AddressDTO>();

        private String lastError;


        public GuiController()
        {
            this.sboApplication = null; // permanece nulo até InitializeGui ser chamado
            this.dataConnector = new DataConnector();
        }

        public void InitializeGui(SAPbouiCOM.Application sboApplication)
        {
            if (sboApplication == null) return;

            this.sboApplication = sboApplication;
            this.sboCompany = (SAPbobsCOM.Company)sboApplication.Company.GetDICompany();
        }

        public void MenuEvent(ref MenuEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
        }

        public void FormDataEvent(ref BusinessObjectInfo businessObjectInfo, out bool bubbleEvent)
        {
            bubbleEvent = true;
            if ((businessObjectInfo.FormTypeEx == "60150") && (businessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_LOAD))
            {
                SAPbouiCOM.Form targetForm = GetSBOForm(businessObjectInfo.FormUID);
                if (targetForm == null) return;

                ReloadAccessories(targetForm);
            }
            if ((businessObjectInfo.FormTypeEx == "60150") && (businessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_UPDATE))
            {
                String message = "Atualize também o contrato caso o equipamento mude de cliente";
                if (businessObjectInfo.BeforeAction) sboApplication.MessageBox(message, 0, "OK", null, null);
            }
            if ((businessObjectInfo.FormTypeEx == "170") && (businessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_LOAD) && (businessObjectInfo.BeforeAction == false))
            {
                SAPbouiCOM.Form targetForm = GetSBOForm(businessObjectInfo.FormUID);
                if (targetForm == null) return;

                bubbleEvent = false;
                FillComments(targetForm.UniqueID);
            }
            if ((businessObjectInfo.FormTypeEx == "133") || (businessObjectInfo.FormTypeEx == "139"))
            {
                SAPbouiCOM.Form targetForm = GetSBOForm(businessObjectInfo.FormUID);
                if (targetForm == null) return;

                SAPbouiCOM.Item contractRefField = targetForm.Items.Item("txtContrct");
                contractRefField.Enabled = false;
                SAPbouiCOM.Item memo = targetForm.Items.Item("txtLista");
                memo.Width = 200; memo.Height = 100; memo.Enabled = false;
            }
        }

        public void ItemEvent(String formUID, ref ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            Boolean formContainsAddress = (pVal.FormType == 134) || (pVal.FormType == 60072) || (pVal.FormType == 60150);
            if ((formContainsAddress) && (pVal.BeforeAction == false))
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                    CreateAddressPicker(formUID);
                if ((pVal.ItemUID == "cmbAddress") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK))
                    AddressComboClick(formUID);
                if ((pVal.ItemUID == "btnCopy") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK))
                    CopyButtonClick(formUID);
            }
            if ((pVal.FormType == 60126) && (pVal.BeforeAction == false))
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                {
                    CreateContractReportButton(formUID);
                }
                if ((pVal.ItemUID == "btnReport") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK))
                    ContractReportButtonClick(formUID);
            }
            /*
            if ((pVal.FormType == 191) && (pVal.BeforeAction == false))
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                    ReplaceGrid(formUID);
                if ((pVal.ItemUID == "boeGrid") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK))
                    SelectCell(pVal.ColUID, pVal.Row, formUID, pVal.Modifiers);
            }
            */
            if ((pVal.FormType == 133) && (pVal.BeforeAction == true))
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE)
                    ConfirmDocBind(formUID);
            }
            if (((pVal.FormType == 133) || (pVal.FormType == 139)) && (pVal.BeforeAction == true))
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                    CreateBillingTab(formUID);
                if ((pVal.ItemUID == "billingTab") && ((pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK) || (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)))
                    OpenBillingTab(formUID);
                if ((pVal.ItemUID == "cmbEqpment") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK))
                    EquipmentComboClick(formUID);
                if ((pVal.ItemUID == "btnGetCntr") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK))
                    GetContract(formUID);
                if ((pVal.ItemUID == "btnAddEqpt") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK))
                    AddEquipment(formUID);
                if ((pVal.ItemUID == "btnCheck") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK))
                    OpenBillingFilter(formUID);
            }
            if ((pVal.FormUID == "frmOptions") && (pVal.BeforeAction == false))
            {
                if ((pVal.ItemUID == "btnOk") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK))
                {
                    CheckBillings(formUID);
                }
            }
            if ((pVal.FormType == 150) && (pVal.BeforeAction == true))
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                    CreateItemUserFields(formUID);
            }
            if ((formUID == "Skill_frmExport") && (pVal.BeforeAction == false))
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                    SetEmailAddress(formUID);
            }
            Boolean isBillOfExchange = (pVal.FormType == 60053);
            if ((isBillOfExchange) && (pVal.BeforeAction == false))
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                {
                    DisplayOurNum(formUID);
                    CreatePaymentLoader(formUID);
                }
                if ((pVal.ItemUID == "btnLoad") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK))
                    LoadPayments(formUID);
            }
            Boolean isJournalEntry = (pVal.FormType == 392);
            if ((isJournalEntry) && (pVal.BeforeAction == false))
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                {
                    CreateReportButton(formUID);
                    CreateEditButton(formUID);
                }
                if ((pVal.ItemUID == "btnReport") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK))
                    OpenReportFilter(formUID);
                if ((pVal.ItemUID == "btnEdit") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK))
                    SelectReport(formUID);
            }
            if ((pVal.FormUID == "frmFilter") && (pVal.BeforeAction == false))
            {
                if ((pVal.ItemUID == "btnOk") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK))
                {
                    BuildAccountTreeReport(formUID);
                    BuildJournalReport(formUID);
                }
            }
            if ((pVal.FormUID == "frmEdit") && (pVal.BeforeAction == false))
            {
                if ((pVal.ItemUID == "btnOk") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK))
                    EditCrystalReport(formUID);
            }
            Boolean requiresStartPage = (pVal.FormType == 165) || (pVal.FormType == 166) || (pVal.FormType == 420) || (pVal.FormType == 604);
            if ((requiresStartPage) && (pVal.BeforeAction == false))
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                    CreatePaginationIndex(formUID);
            }
            if ((pVal.FormUID == "frmStartPage") && (pVal.BeforeAction == false))
            {
                if ((pVal.ItemUID == "btnOk") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK))
                    SavePaginationIndex(formUID);
            }
            if ((pVal.FormType == 60150) && (pVal.BeforeAction == false))
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                {
                    CreateInstallationTab(formUID);
                    CreateSpecificationsTab(formUID);
                }
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE)
                {
                    ResizeInstallationTab(formUID);
                    ResizeSpecificationsTab(formUID);
                }
                if ((pVal.ItemUID == "instTab") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK))
                    OpenInstallationTab(formUID);
                if ((pVal.ItemUID == "specTab") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK))
                    OpenSpecificationsTab(formUID);

                if ((pVal.ItemUID == "accessries") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK))
                    SelectAccessory(pVal.Row, formUID);
                if ((pVal.ItemUID == "btnAdd") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK))
                    AddAccessory(formUID);
                if ((pVal.ItemUID == "btnRemove") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK))
                    RemoveAccessory(formUID);
            }
            if ((formUID == "frmAccssry") && (pVal.BeforeAction == false))
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                    ChooseItem(pVal);
                if ((pVal.ItemUID == "btnAdd") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK))
                    SaveNewAccessory(formUID);
            }
            Boolean formContainsTax = (pVal.FormType == 133) || (pVal.FormType == 139) || (pVal.FormType == 140);
            formContainsTax = formContainsTax || (pVal.FormType == 149) || (pVal.FormType == 179) || (pVal.FormType == 180);
            formContainsTax = formContainsTax || (pVal.FormType == 60091) || (pVal.FormType == 65300);
            Boolean isTaxTabEvent = (pVal.ItemUID == "2013");
            if ((formContainsTax) && (isTaxTabEvent) && (pVal.BeforeAction == false))
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                    SetDefaultCarrier(formUID);
            }
        }

        private void CreateAddressPicker(String formUID)
        {
            SAPbouiCOM.Form activeForm = sboApplication.Forms.ActiveForm;
            SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            int targetPanel = ADDRESS_TAB;
            int leftPos = 20;
            int topPos = 420;
            if (targetForm.Type == 60072)
            {
                targetPanel = 0;
                leftPos = 280;
                topPos = 10;
            }
            if (targetForm.Type == 60150)
            {
                targetPanel = 1;
                leftPos = 300;
                topPos = 220;
            }

            if ((targetForm.Type == 60072) && ((activeForm.Type == 133) || (activeForm.Type == 139) || (activeForm.Type == 140) || (activeForm.Type == 180)))
            {
                SAPbouiCOM.Item textItem = activeForm.Items.Item("4");
                SAPbouiCOM.EditText txtCardCod = (SAPbouiCOM.EditText)textItem.Specific;
                UserDataSource formParent = targetForm.DataSources.UserDataSources.Add("formParent", BoDataType.dt_SHORT_TEXT, 30);
                formParent.Value = activeForm.UniqueID;
            }

            SAPbouiCOM.Item labelItem = targetForm.Items.Add("lblAdress", BoFormItemTypes.it_STATIC);
            labelItem.Left = leftPos;
            labelItem.Top = topPos;
            labelItem.FromPane = targetPanel;
            labelItem.ToPane = targetPanel;
            SAPbouiCOM.StaticText lblAdress = (SAPbouiCOM.StaticText)labelItem.Specific;
            lblAdress.Caption = "Endereço";

            SAPbouiCOM.Item comboboxItem = targetForm.Items.Add("cmbAddress", BoFormItemTypes.it_COMBO_BOX);
            comboboxItem.Left = leftPos + 60;
            comboboxItem.Top = topPos;
            comboboxItem.FromPane = targetPanel;
            comboboxItem.ToPane = targetPanel;
            SAPbouiCOM.ComboBox cmbAddress = (SAPbouiCOM.ComboBox)comboboxItem.Specific;

            SAPbouiCOM.Item buttonItem = targetForm.Items.Add("btnCopy", BoFormItemTypes.it_BUTTON);
            buttonItem.Left = leftPos + 150;
            buttonItem.Top = topPos - 3;
            buttonItem.FromPane = targetPanel;
            buttonItem.ToPane = targetPanel;
            SAPbouiCOM.Button btnCopy = (SAPbouiCOM.Button)buttonItem.Specific;
            btnCopy.Caption = "Copiar";
        }

        private void AddressComboClick(String formUID)
        {
            try
            {
                SAPbouiCOM.Form targetForm = GetSBOForm(formUID);

                String cardCodeItemUID = "5";
                if (targetForm.Type == 60150) cardCodeItemUID = "48";
                SAPbouiCOM.Form cardCodeOwner = targetForm;
                UserDataSource formParent = GetSBOUserDataSource(targetForm, "formParent");
                if (formParent != null)
                {
                    cardCodeItemUID = "4";
                    cardCodeOwner = GetSBOForm(formParent.Value);
                }

                SAPbouiCOM.Item textItem = cardCodeOwner.Items.Item(cardCodeItemUID);
                SAPbouiCOM.EditText txtCardCode = (SAPbouiCOM.EditText)textItem.Specific;
                String cardCode = txtCardCode.Value;
                if (String.IsNullOrEmpty(cardCode))
                {
                    sboApplication.MessageBox("Escolha o cliente primeiro!", 0, "OK", null, null);
                    return;
                }

                dataConnector.OpenConnection("sqlServer");
                dataConnector.SqlServerConnection.ChangeDatabase(sboApplication.Company.DatabaseName);
                AddressDAO addressDAO = new AddressDAO(dataConnector.SqlServerConnection);
                List<AddressDTO> addressList = addressDAO.GetPartnerAddresses(cardCode);
                EquipmentDAO equipmentDAO = new EquipmentDAO(dataConnector.SqlServerConnection);
                List<EquipmentDTO> equipmentList = equipmentDAO.GetCustomerEquipments(cardCode);
                dataConnector.CloseConnection();

                SAPbouiCOM.Item comboboxItem = targetForm.Items.Item("cmbAddress");
                SAPbouiCOM.ComboBox cmbAddress = (SAPbouiCOM.ComboBox)comboboxItem.Specific;
                // Remove os items do dicionário e do combo
                addressDictionary.Clear();
                while (cmbAddress.ValidValues.Count > 0)
                    cmbAddress.ValidValues.Remove(cmbAddress.ValidValues.Count - 1, BoSearchKey.psk_Index);
                // Adiciona os items recuperados da consulta ao banco
                foreach (AddressDTO addrDTO in addressList)
                {
                    String addressName = addrDTO.Address;
                    if (addressDictionary.ContainsKey(addrDTO.Address))
                        addressName = addrDTO.Address + " " + addrDTO.ZipCode;

                    if (addressName.Length < 48) // Só adiciona no combo os valores que não excedem o tamanho
                    {
                        addressDictionary.Add(addressName, addrDTO);
                        cmbAddress.ValidValues.Add(addressName, null);
                    }
                }
                foreach (EquipmentDTO equipmentDTO in equipmentList)
                {
                    AddressDTO addressDTO = new AddressDTO();
                    addressDTO.Address = equipmentDTO.ManufSN;
                    addressDTO.AddrType = equipmentDTO.AddrType;
                    addressDTO.Street = equipmentDTO.Street;
                    addressDTO.StreetNo = equipmentDTO.StreetNo;
                    addressDTO.Building = equipmentDTO.Building;
                    addressDTO.ZipCode = equipmentDTO.Zip;
                    addressDTO.Block = equipmentDTO.Block;
                    addressDTO.City = equipmentDTO.City;
                    addressDTO.State = equipmentDTO.State;
                    addressDTO.County = equipmentDTO.County;
                    addressDTO.Country = equipmentDTO.Country;

                    if (!addressDictionary.ContainsKey(addressDTO.Address)) // Verifica se já foi adicionado
                    {
                        addressDictionary.Add(addressDTO.Address, addressDTO);
                        cmbAddress.ValidValues.Add(addressDTO.Address, EquipmentDAO.GetStatusDescription(equipmentDTO.Status));
                    }
                }
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        private void CopyButtonClick(String formUID)
        {
            try
            {
                SAPbouiCOM.Form targetForm = GetSBOForm(formUID);

                SAPbouiCOM.Item comboboxItem = targetForm.Items.Item("cmbAddress");
                SAPbouiCOM.ComboBox cmbAddress = (SAPbouiCOM.ComboBox)comboboxItem.Specific;
                if (cmbAddress.Selected == null)
                {
                    sboApplication.MessageBox("Favor selecionar uma das opções ao lado!", 0, "OK", null, null);
                    return;
                }
                String selectedAddress = cmbAddress.Selected.Value;

                if (targetForm.Type == 60072)
                {
                    SAPbouiCOM.Item remarksItem = targetForm.Items.Item("7");
                    SAPbouiCOM.EditText remarks = (SAPbouiCOM.EditText)remarksItem.Specific;
                    remarks.Value = addressDictionary[selectedAddress].Address + "   " +
                                    "Endereço: " + addressDictionary[selectedAddress].AddrType + " " +
                                    addressDictionary[selectedAddress].Street + ", " +
                                    addressDictionary[selectedAddress].StreetNo + "   " +
                                    addressDictionary[selectedAddress].Building + "   " +
                                    "Secretaria: " + addressDictionary[selectedAddress].U_Secretaria + "   " +
                                    "CEP: " + addressDictionary[selectedAddress].ZipCode + "   " +
                                    "Bairro: " + addressDictionary[selectedAddress].Block + "   " +
                                    addressDictionary[selectedAddress].City + " " +
                                    addressDictionary[selectedAddress].State + " " +
                                    addressDictionary[selectedAddress].Country;
                    return;
                }

                if (targetForm.Type == 60150)
                {
                    SAPbouiCOM.EditText addressType = GetSBOEditText(targetForm, "2004");
                    addressType.Value = addressDictionary[selectedAddress].AddrType;
                    SAPbouiCOM.EditText street = GetSBOEditText(targetForm, "63");
                    street.Value = addressDictionary[selectedAddress].Street;
                    SAPbouiCOM.EditText streetNo = GetSBOEditText(targetForm, "2006");
                    streetNo.Value = addressDictionary[selectedAddress].StreetNo;
                    SAPbouiCOM.EditText building = GetSBOEditText(targetForm, "2001");
                    building.Value = addressDictionary[selectedAddress].Building;
                    SAPbouiCOM.EditText zip = GetSBOEditText(targetForm, "69");
                    zip.Value = addressDictionary[selectedAddress].ZipCode;
                    SAPbouiCOM.EditText block = GetSBOEditText(targetForm, "66");
                    block.Value = addressDictionary[selectedAddress].Block;
                    SAPbouiCOM.EditText city = GetSBOEditText(targetForm, "67");
                    city.Value = addressDictionary[selectedAddress].City;
                    SAPbouiCOM.ComboBox country = GetSBOComboBox(targetForm, "76");
                    country.Select(addressDictionary[selectedAddress].Country, BoSearchKey.psk_ByValue);
                    SAPbouiCOM.ComboBox state = GetSBOComboBox(targetForm, "75");
                    state.Select(addressDictionary[selectedAddress].State, BoSearchKey.psk_ByValue);
                    SAPbouiCOM.ComboBox county = GetSBOComboBox(targetForm, "2002");
                    county.Select(addressDictionary[selectedAddress].County, BoSearchKey.psk_ByValue);
                    return;
                }

                int dialogResult = sboApplication.MessageBox("Deseja copiar o endereço para etiquetas ?", 0, "Sim", "Não", null);
                if (dialogResult == 1)
                {
                    BuildTagSheet(addressDictionary[selectedAddress]);
                    return;
                }

                SAPbouiCOM.Item matrixItem = targetForm.Items.Item("178");
                SAPbouiCOM.Matrix addressMatrix = (SAPbouiCOM.Matrix)matrixItem.Specific;
                int columnCount = addressMatrix.Columns.Count;
                if (columnCount == 0) return;

                Column column1 = addressMatrix.Columns.Item("1");
                // Mantem o valor do campo // SetColumnData(column1, "Definir Novo");
                Column column2 = addressMatrix.Columns.Item("2002");
                SetColumnData(column2, addressDictionary[selectedAddress].AddrType);
                Column column3 = addressMatrix.Columns.Item("2");
                SetColumnData(column3, addressDictionary[selectedAddress].Street);
                Column column4 = addressMatrix.Columns.Item("2003");
                SetColumnData(column4, addressDictionary[selectedAddress].StreetNo);
                Column column5 = addressMatrix.Columns.Item("2000");
                SetColumnData(column5, addressDictionary[selectedAddress].Building);
                Column column6 = addressMatrix.Columns.Item("5");
                SetColumnData(column6, addressDictionary[selectedAddress].ZipCode);
                Column column7 = addressMatrix.Columns.Item("3");
                SetColumnData(column7, addressDictionary[selectedAddress].Block);
                Column column8 = addressMatrix.Columns.Item("4");
                SetColumnData(column8, addressDictionary[selectedAddress].City);
                Column column9 = addressMatrix.Columns.Item("7");
                SetColumnData(column9, addressDictionary[selectedAddress].State);
                Column column10 = addressMatrix.Columns.Item("6");
                SetColumnData(column10, addressDictionary[selectedAddress].County);
                Column column11 = addressMatrix.Columns.Item("8");
                SetColumnData(column11, addressDictionary[selectedAddress].Country);
                Column column12 = addressMatrix.Columns.Item("U_Secretaria");
                SetColumnData(column12, addressDictionary[selectedAddress].U_Secretaria);
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        private void SetColumnData(Column column, Object data)
        {
            switch (column.Type)
            {
                case BoFormItemTypes.it_COMBO_BOX:
                    foreach (Cell cell in column.Cells)
                    {
                        SAPbouiCOM.ComboBox comboboxCell = (SAPbouiCOM.ComboBox)cell.Specific;
                        comboboxCell.Select((String)data, BoSearchKey.psk_ByValue);
                    }
                    break;
                case BoFormItemTypes.it_EDIT:
                    foreach (Cell cell in column.Cells)
                    {
                        SAPbouiCOM.EditText textCell = (SAPbouiCOM.EditText)cell.Specific;
                        textCell.Value = (String)data;
                    }
                    break;
                default:
                    break;
            }
        }

        private void ReplaceGrid(String formUID)
        {
            SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            if (targetForm == null) return;

            try
            {
                DBDataSource dbDataSource = targetForm.DataSources.DBDataSources.Item(0);
                String dbTableName = dbDataSource.TableName;
                if (dbTableName != "OBOE") return;

                Item matrixItem = targetForm.Items.Item("3");
                Matrix matrixSpecific = (Matrix)matrixItem.Specific;
                List<int> boeNumbers = new List<int>();
                foreach (Cell cell in matrixSpecific.Columns.Item(2).Cells)
                {
                    EditText cellSpecific = (EditText)cell.Specific;
                    int boeNum = int.Parse(cellSpecific.Value);
                    boeNumbers.Add(boeNum);
                }

                dataConnector.OpenConnection("sqlServer");
                dataConnector.SqlServerConnection.ChangeDatabase(sboApplication.Company.DatabaseName);
                BillOfExchangeDAO boeDAO = new BillOfExchangeDAO(dataConnector.SqlServerConnection);
                List<BillOfExchangeDTO> boeList = boeDAO.GetBillsOfExchange(boeNumbers.ToArray());
                dataConnector.CloseConnection();

                Dictionary<int, Object> boeDictionary = new Dictionary<int, Object>();
                foreach (BillOfExchangeDTO boe in boeList)
                    boeDictionary.Add(boe.BoeNum, boe);

                Item gridItem = targetForm.Items.Add("boeGrid", BoFormItemTypes.it_GRID);
                gridItem.Top = matrixItem.Top;
                gridItem.Left = matrixItem.Left;
                gridItem.Width = matrixItem.Width;
                gridItem.Height = matrixItem.Height;
                Grid gridSpecific = (Grid)gridItem.Specific;
                gridSpecific.SelectionMode = BoMatrixSelect.ms_Auto;

                SAPbouiCOM.DataTable gridTable = targetForm.DataSources.DataTables.Add("boeTable");
                gridSpecific.DataTable = gridTable;
                gridSpecific.DataTable.Clear();
                gridSpecific.DataTable.Columns.Add("boeNum", BoFieldsType.ft_Integer, 10);
                gridSpecific.DataTable.Columns.Add("dueDate", BoFieldsType.ft_Date, 20);
                gridSpecific.DataTable.Columns.Add("boeSum", BoFieldsType.ft_Text, 50);
                gridSpecific.DataTable.Columns.Add("refNF", BoFieldsType.ft_Text, 50);
                gridSpecific.DataTable.Columns.Add("cardCode", BoFieldsType.ft_Text, 50);
                gridSpecific.DataTable.Columns.Add("cardName", BoFieldsType.ft_Text, 250);
                gridSpecific.RowHeaders.Width = 30;
                gridSpecific.Columns.Item(0).TitleObject.Caption = "Boleto";
                gridSpecific.Columns.Item(0).Width = 45;
                gridSpecific.Columns.Item(0).Editable = false;
                gridSpecific.Columns.Item(1).TitleObject.Caption = "Vencimento";
                gridSpecific.Columns.Item(1).Width = 80;
                gridSpecific.Columns.Item(1).Editable = false;
                gridSpecific.Columns.Item(2).TitleObject.Caption = "Total";
                gridSpecific.Columns.Item(2).Width = 80;
                gridSpecific.Columns.Item(2).Editable = false;
                gridSpecific.Columns.Item(3).TitleObject.Caption = "Ref. NF";
                gridSpecific.Columns.Item(3).Width = 50;
                gridSpecific.Columns.Item(3).Editable = false;
                gridSpecific.Columns.Item(4).TitleObject.Caption = "Cod. Cliente";
                gridSpecific.Columns.Item(4).Width = 80;
                gridSpecific.Columns.Item(4).Editable = false;
                gridSpecific.Columns.Item(5).TitleObject.Caption = "Cliente";
                gridSpecific.Columns.Item(5).Width = 200;
                gridSpecific.Columns.Item(5).Editable = false;
                int rowCount = matrixSpecific.Columns.Item(0).Cells.Count;
                gridSpecific.DataTable.Rows.Add(rowCount);
                for (int index = 1; index <= rowCount; index++)
                {
                    // Realiza a cópia das linhas da matrix para o grid
                    Cell cell = matrixSpecific.Columns.Item(2).Cells.Item(index);
                    EditText cellSpecific = (EditText)cell.Specific;
                    int boeNum = int.Parse(cellSpecific.Value);
                    BillOfExchangeDTO boeDTO = (BillOfExchangeDTO)boeDictionary[boeNum];
                    String boeTotal = String.Format("{0:0.00}", boeDTO.BoeSum);
                    gridSpecific.DataTable.SetValue(0, index - 1, boeDTO.BoeNum);
                    gridSpecific.DataTable.SetValue(1, index - 1, boeDTO.DueDate);
                    gridSpecific.DataTable.SetValue(2, index - 1, boeTotal);
                    gridSpecific.DataTable.SetValue(3, index - 1, "" + boeDTO.RefNum);
                    gridSpecific.DataTable.SetValue(4, index - 1, boeDTO.CardCode);
                    gridSpecific.DataTable.SetValue(5, index - 1, boeDTO.CardName);
                }
                matrixItem.Left = 1000; // oculta a matrix, jogando ela para fora do form
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        private void SelectCell(String columnName, int row, String formUID, BoModifiersEnum modifiers)
        {
            SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            if (targetForm == null) return;

            Item gridItem = targetForm.Items.Item("boeGrid");
            Grid gridSpecific = (Grid)gridItem.Specific;

            Item matrixItem = targetForm.Items.Item("3");
            Matrix matrixSpecific = (Matrix)matrixItem.Specific;

            try
            {
                // Limpa a seleção previa
                for (int rowIndex = 1; rowIndex <= matrixSpecific.Columns.Item(0).Cells.Count; rowIndex++)
                {
                    Cell targetCell = matrixSpecific.Columns.Item(0).Cells.Item(rowIndex);
                    if (matrixSpecific.IsRowSelected(rowIndex))
                        targetCell.Click(BoCellClickType.ct_Regular, (int)BoModifiersEnum.mt_CTRL);
                }
                // Seleciona as mesmas colunas na matrix
                foreach (int rowIndex in gridSpecific.Rows.SelectedRows)
                {
                    Cell targetCell = matrixSpecific.Columns.Item(0).Cells.Item(rowIndex + 1);
                    if (!matrixSpecific.IsRowSelected(rowIndex + 1))
                        targetCell.Click(BoCellClickType.ct_Regular, (int)BoModifiersEnum.mt_CTRL);
                }
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        private void OpenInvoicePayment()
        {
            //SAPbouiCOM.Form activeForm = sboApplication.Forms.ActiveForm;
            //if (activeForm.Type != 133) return;

            //SAPbouiCOM.Item invoiceNumber = activeForm.Items.Item("8");
            //SAPbouiCOM.EditText invoiceNumSpecific = (SAPbouiCOM.EditText)invoiceNumber.Specific;
            //int invoiceNum = int.Parse(invoiceNumSpecific.Value);

            //sboApplication.ActivateMenuItem("2817");
            //Thread.Sleep(500); // Aguarda meio segundo
            //SAPbouiCOM.Form openingForm = sboApplication.Forms.ActiveForm;
            //if (openingForm.Type != 170) return;

            //dataConnector.OpenConnection("sqlServer");
            //dataConnector.SqlServerConnection.ChangeDatabase(sboApplication.Company.DatabaseName);
            //InvoicePaymentDAO invoicePaymentDAO = new InvoicePaymentDAO(dataConnector.SqlServerConnection);
            //InvoicePaymentDTO payment = invoicePaymentDAO.GetPayment(invoiceNum);
            //int paymentNum = 0; if (payment != null) paymentNum = payment.docNum;
            //dataConnector.CloseConnection();

            //SAPbouiCOM.Item paymentNumber = openingForm.Items.Item("3");
            //SAPbouiCOM.EditText paymentNumSpecific = (SAPbouiCOM.EditText)paymentNumber.Specific;
            //paymentNumSpecific.Value = paymentNum.ToString();
        }

        private void FillComments(String formUID)
        {
            //SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            //if (targetForm == null) return;

            //SAPbouiCOM.Item jrnlMemoItem = targetForm.Items.Item("59");
            //SAPbouiCOM.EditText jrnlMemo = (SAPbouiCOM.EditText)jrnlMemoItem.Specific;
            //jrnlMemo.ClickPicker();

            //SAPbouiCOM.Item commentsItem = targetForm.Items.Item("26");
            //SAPbouiCOM.EditText comments = (SAPbouiCOM.EditText)commentsItem.Specific;
            //comments.ClickPicker();

            //sboApplication.MessageBox("Verifique o campo de observações", 0, "OK", null, null);
        }

        private void CreateBillingTab(String formUID)
        {
            SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            if (targetForm == null) return;

            try
            {
                // Cria a aba "Faturamento" no form
                SAPbouiCOM.Item lastFolder = targetForm.Items.Item("2013");
                SAPbouiCOM.Item newFolder = targetForm.Items.Add("billingTab", BoFormItemTypes.it_FOLDER);
                newFolder.Width = lastFolder.Width;
                newFolder.Height = lastFolder.Height;
                newFolder.Top = lastFolder.Top;
                newFolder.Left = lastFolder.Left + lastFolder.Width;
                newFolder.Visible = true;
                SAPbouiCOM.Folder billingTab = ((SAPbouiCOM.Folder)(newFolder.Specific));
                billingTab.Caption = "Faturamento";
                billingTab.GroupWith("2013");

                // Adiciona os campos de usuário a aba de "Faturamento"
                String tableName = "OINV"; if (targetForm.Type == 139) tableName = "ORDR";
                SAPbouiCOM.EditText billingRef = AddSBOTextField(targetForm, "Billing", "Nº do Faturamento", 25, newFolder.Top + 35, BILLING_TAB);
                billingRef.DataBind.SetBound(true, tableName, "U_demFaturamento");
                SAPbouiCOM.EditText contractRef = AddSBOTextField(targetForm, "Contrct", "Contrato(SYS ID)", 25, newFolder.Top + 60, BILLING_TAB);
                contractRef.DataBind.SetBound(true, tableName, "U_CONTRATO");
                SAPbouiCOM.Item contractRefField = targetForm.Items.Item("txtContrct");
                contractRefField.Enabled = false;

                SAPbouiCOM.ComboBox equipamento = AddSBOCombobox(targetForm, "Eqpment", "Equipamento", 25, newFolder.Top + 125, BILLING_TAB);
                SAPbouiCOM.EditText listaEquipamentos = AddSBOTextField(targetForm, "Lista", "Lista Equipamentos", 25, newFolder.Top + 150, BILLING_TAB, true);
                listaEquipamentos.ScrollBars = BoScrollBars.sb_Vertical;
                listaEquipamentos.TextStyle = (int)SAPbouiCOM.BoTextStyle.ts_BOLD;
                listaEquipamentos.DataBind.SetBound(true, tableName, "U_ListaEquipamentos");
                SAPbouiCOM.Item memo = targetForm.Items.Item("txtLista");
                memo.Width = 200; memo.Height = 100; memo.Enabled = false;
                SAPbouiCOM.Button getContractButton = AddSBOButton(targetForm, "GetCntr", "Obter Contrato", 225, newFolder.Top + 57, BILLING_TAB);
                SAPbouiCOM.Button addEquipmentButton = AddSBOButton(targetForm, "AddEqpt", "Incluir", 225, newFolder.Top + 122, BILLING_TAB);
                SAPbouiCOM.Button checkButton = AddSBOButton(targetForm, "Check", "Verificar Lançamentos", 25, newFolder.Top + 200, BILLING_TAB);

                // Recoloca o combo de equipamentos na aba de faturamento ( correção de bug )
                SAPbouiCOM.Item cmbEquipamento = targetForm.Items.Item("cmbEqpment");
                cmbEquipamento.FromPane = BILLING_TAB;
                cmbEquipamento.ToPane = BILLING_TAB;
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }

        }

        private void OpenBillingTab(String formUID)
        {
            // Altera o pane level do form para BILLING_TAB, fazendo com que o SAP exiba apenas controles desta aba
            SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            if (targetForm != null) targetForm.PaneLevel = BILLING_TAB;

            SAPbouiCOM.Item contractRefField = targetForm.Items.Item("txtContrct");
            contractRefField.Enabled = false;
            SAPbouiCOM.Item memo = targetForm.Items.Item("txtLista");
            memo.Width = 200; memo.Height = 100; memo.Enabled = false;
        }

        private void ConfirmDocBind(String formUID)
        {
            SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            if (targetForm == null) return;

            // Verifica se a natureza da operação(campo usage) é locação de equipamentos ou assistência técnica, aborta
            // caso não seja nenhum dos dois
            DBDataSource invoiceItems = targetForm.DataSources.DBDataSources.Item("INV1");
            String usage = invoiceItems.GetValue("Usage", 0);
            if ((usage != "18") && (usage != "17")) return;

            SAPbouiCOM.EditText billingRef = GetSBOEditText(targetForm, "txtBilling");
            int billingId = 0;
            if (String.IsNullOrEmpty(billingRef.Value) || (!int.TryParse(billingRef.Value, out billingId)))
            {
                sboApplication.MessageBox("Informe corretamente o número do faturamento.", 0, "OK", null, null);
                return;
            }

            SAPbouiCOM.EditText contractRef = GetSBOEditText(targetForm, "txtContrct");
            if (String.IsNullOrEmpty(contractRef.Value))
            {
                sboApplication.MessageBox("Favor preencher o número do contrato na aba Faturamento!", 0, "OK", null, null);
                return;
            }

            dataConnector.OpenConnection("mySql");
            BillingDAO billingDAO = new BillingDAO(dataConnector.MySqlConnection);
            BillingDTO billing = billingDAO.GetBilling(billingId);
            // billing.date = 0;
            // billing.total = 0;
            // billingDAO.SetBilling(billing);
            dataConnector.CloseConnection();
        }

        private void GetContract(String formUID)
        {
            SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            if (targetForm == null) return;

            SAPbouiCOM.Item textItem = targetForm.Items.Item("4");
            SAPbouiCOM.EditText textSpecific = (SAPbouiCOM.EditText)textItem.Specific;
            String cardCode = textSpecific.Value;
            if (String.IsNullOrEmpty(cardCode))
            {
                sboApplication.MessageBox("Escolha o cliente primeiro!", 0, "OK", null, null);
                return;
            }

            SAPbouiCOM.EditText billingRef = GetSBOEditText(targetForm, "txtBilling");
            int billingId = 0;
            if (String.IsNullOrEmpty(billingRef.Value) || (!int.TryParse(billingRef.Value, out billingId)))
            {
                sboApplication.MessageBox("Informe corretamente o número do faturamento.", 0, "OK", null, null);
                return;
            }

            try
            {
                dataConnector.OpenConnection("both");
                BillingDAO billingDAO = new BillingDAO(dataConnector.MySqlConnection);
                BillingDTO billing = billingDAO.GetBilling(billingId);
                if (billing == null)
                {
                    sboApplication.MessageBox("Número do faturamento incorreto!", 0, "OK", null, null);
                    dataConnector.CloseConnection();
                    return;
                }
                MailingDAO mailingDAO = new MailingDAO(dataConnector.MySqlConnection);
                MailingDTO mailing = mailingDAO.GetMailing(billing.mailing_id);
                dataConnector.CloseConnection();

                if (billing.businessPartnerCode != cardCode) // Verifica se o cliente do demonstrativo é o mesmo da nota fiscal
                {
                    String warningMessage = "O cliente não é o mesmo. Este faturamento é referente ao cliente " + billing.businessPartnerName + Environment.NewLine + Environment.NewLine +
                                            "Verifique o campo cliente no Contrato,  no Demonstrativo, e nos Cartões de Equipamento";
                    sboApplication.MessageBox(warningMessage, 0, "OK", null, null);
                    return;
                }

                SAPbouiCOM.EditText contractRef = GetSBOEditText(targetForm, "txtContrct");
                contractRef.Value = mailing.codigoContrato.ToString();

                String appDataFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                String reportFolder = Path.Combine(appDataFolder, "OsiReports");
                Directory.CreateDirectory(reportFolder);
                String filename = Path.Combine(reportFolder, "ContratosCliente.htm");

                FileStream fileStream = new FileStream(filename, FileMode.Create);
                StreamWriter streamWriter = new StreamWriter(fileStream);

                String contractInfo = "";
                if (mailing.codigoContrato != 0) contractInfo = MountContractInfo(mailing.codigoContrato, true); else contractInfo = MountClientContractsInfo(cardCode, true);
                String billingInfo = "";
                if (billing != null) billingInfo = MountBillingInfo(billing.id);
                streamWriter.WriteLine("<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Strict//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd'>");
                streamWriter.WriteLine("<html xmlns='http://www.w3.org/1999/xhtml'>");
                streamWriter.WriteLine("<head>");
                streamWriter.WriteLine("    <meta http-equiv='Content-type' content='text/html; charset=UTF-8' />");
                streamWriter.WriteLine("    <meta http-equiv='Content-Language' content='pt-br' />");
                streamWriter.WriteLine("    <title>Contratos do Cliente</title>");
                streamWriter.WriteLine("</head>");
                streamWriter.WriteLine("<body style='text-align:center;font-size:15px;color:#0073EA;'>");
                streamWriter.WriteLine(contractInfo);
                streamWriter.WriteLine("<div style='clear:both;'><br/><br/></div>");
                streamWriter.WriteLine(billingInfo);
                streamWriter.WriteLine("</body>");
                streamWriter.WriteLine("</html>");
                streamWriter.Close();
                System.Diagnostics.Process.Start(filename);
            }
            catch (Exception exc)
            {
                sboApplication.MessageBox(exc.Message, 0, "OK", null, null);
            }
        }

        // Utilizando linked Server dentro do SQL SERVER 2008
        //
        // String query = "SELECT EQP.itemName AS equipamento, MDL.modelo, FAB.FirmName, CTR.nome AS contador, ITM.total AS valorAluguel, ITM.acrescimoDesconto" + Environment.NewLine +
        //               "FROM MYSQL...itemFaturamento ITM" + Environment.NewLine +
        //               "JOIN MYSQL...contador CTR ON ITM.counterId = CTR.id" + Environment.NewLine +
        //               "JOIN OINS EQP ON ITM.codigoCartaoEquipamento = EQP.insID" + Environment.NewLine +
        //               "JOIN MYSQL...modeloEquipamento MDL ON EQP.U_Model = MDL.id" + Environment.NewLine +
        //               "JOIN OMRC FAB ON MDL.fabricante = FAB.FirmCode" + Environment.NewLine +
        //               "WHERE codigoFaturamento = @billingId";

        private String MountBillingInfo(int billingId)
        {
            String content = "";

            dataConnector.OpenConnection("both");
            BillingDAO billingDAO = new BillingDAO(dataConnector.MySqlConnection);
            BillingDTO billing = billingDAO.GetBilling(billingId);
            content += "<head>";
            content += "<style type='text/css'>";
            content += "    table{  border-left:1px solid black; border-top:1px solid black; width:98%; margin-left:auto; margin-right:auto; border-spacing:0; font-size: 11px; }";
            content += "    td{  border-right:1px solid black; border-bottom:1px solid black; margin:0; padding:0; text-align:center;  }";
            content += "    th{  border-right:1px solid black; border-bottom:1px solid black; margin:0; padding:0; text-align:center;  }";
            content += "</style>";
            content += "</head>";
            content += "<h3 style='border:0; margin:0;' >DEMONSTRATIVO DE FATURAMENTO (Nº " + billing.id.ToString().PadLeft(5, '0') + ")</h3><br/>";
            content += "Data Inicial: " + billing.dataInicial.ToString("dd/MM/yyyy") + "&nbsp;&nbsp; Data Final: " + billing.dataFinal.ToString("dd/MM/yyyy") + "<br/>";
            content += "<div style='clear:both;'><br/><br/></div>";
            content += "<table>";
            content += "<tr bgcolor='YELLOW' style='height:30px;' ><td>Tipo do Contador</td><td>Data de Leitura</td><td>Medição Final</td><td>Medição Inicial</td><td>Consumo</td><td>Franquia</td><td>Excedente (Págs.)</td><td>Tarifa sobre exced.</td><td>Valor Fixo (R$)</td><td>Valor Variável (R$)</td><td>Valor Total (R$)</td></tr>";
            CounterDAO counterDAO = new CounterDAO(dataConnector.MySqlConnection);
            Dictionary<int, String> counters = counterDAO.GetAllCounter();
            Dictionary<int, BillingSummary> summaries = new Dictionary<int, BillingSummary>();
            foreach(KeyValuePair<int, String> counter in counters)
                summaries.Add(counter.Key, new BillingSummary(counter.Key, counter.Value));

            BillingItemDAO billingItemDAO = new BillingItemDAO(dataConnector.MySqlConnection);
            EquipmentDAO equipmentDAO = new EquipmentDAO(dataConnector.SqlServerConnection);
            List<BillingItemDTO> billingItems = billingItemDAO.GetBillingItems(billingId);
            foreach (BillingItemDTO billingItem in billingItems)
            {
                int counterId = billingItem.counterId;
                String dataLeitura = billingItem.dataLeitura.ToString("dd/MM/yyyy");
                if (billingItem.dataLeitura.Equals(DateTime.MinValue)) dataLeitura = "Sem Leitura";
                EquipmentDTO equipment = equipmentDAO.GetEquipment(billingItem.codigoCartaoEquipamento);
                String equipmentInfo = "Cartão Equipamento: " + equipment.InsID + "".PadLeft(3, ' ') + "Modelo: " + equipment.ItemName + "".PadLeft(3, ' ');
                equipmentInfo += "Série: " + equipment.ManufSN + " (" + equipment.InternalSN + ") " + "".PadLeft(3, ' ') + "Departamento: " + equipment.InstLocation + "".PadLeft(3, ' ');
                equipmentInfo += "Tipo: " + billingItem.tipoLocacao;

                content += "<tr bgcolor='LIGHTGRAY' ><td colspan='11' >" + equipmentInfo  + "</td></tr>";
                content += "<tr>";
                content += "<td>" + billingItem.counterName + "</td><td>" + dataLeitura + "</td>";
                content += "<td>" + billingItem.medicaoFinal + "</td><td>" + billingItem.medicaoInicial + "</td>";
                content += "<td>" + billingItem.consumo + "<br/>(Acrésc/Desc = " + billingItem.ajuste + ")</td>";
                content += "<td>" + billingItem.franquia + "</td><td>" + billingItem.excedente + "</td>";
                content += "<td>" + String.Format("{0:0.00000}", billingItem.tarifaSobreExcedente) + "</td><td>" + String.Format("{0:0.00}", billingItem.fixo) + "</td>";
                content += "<td>" + String.Format("{0:0.00}", billingItem.variavel) + "</td><td>" + String.Format("{0:0.00}", billingItem.total) + "</td>";
                content += "</tr>";

                summaries[counterId].consumo += billingItem.consumo;
                summaries[counterId].franquia += billingItem.franquia;
                summaries[counterId].excedente += billingItem.excedente;
                summaries[counterId].fixo += billingItem.fixo;
                summaries[counterId].variavel += billingItem.variavel;
                summaries[counterId].total += billingItem.total;
            }
            content += "</table>";
            content += "<div style='clear:both;'><br/></div>";

            content += "<h3 style='border:0; margin:0;' >&nbsp;&nbsp;QUADRO RESUMO</h3>";
            content += "<table>";
            content += "<tr bgcolor='LIGHTGRAY' ><td>Tipo do Contador</td><td>Consumo</td><td>Franquia</td><td>Excedente</td><td>Valor Fixo (R$)</td><td>Valor Variável (R$)</td><td>Valor Total (R$)</td></tr>";
            foreach (KeyValuePair<int, BillingSummary> billingSummary in summaries)
            {
                BillingSummary summary = billingSummary.Value;
                if (summary.total > 0) content += "<tr bgcolor='WHITE' ><td>" + summary.counterName + "</td><td>" + summary.consumo + "</td><td>" + summary.franquia + "</td><td>" + summary.excedente + "</td><td>" + String.Format("{0:0.00}", summary.fixo) + "</td><td>" + String.Format("{0:0.00}", summary.variavel) + "</td><td>" + String.Format("{0:0.00}", summary.total) + "</td></tr>";
            }
            content += "</table>";

            content += "<div style='width:70%; margin-left:auto; margin-right:auto;' ><h4>Acrescimo/Desconto: " + String.Format("{0:0.00}", billing.acrescimoDesconto) + "</h4></div>";
            content += "<div style='width:70%; margin-left:auto; margin-right:auto;' ><h4>Observações: " + billing.obs + "</h4></div>";
            dataConnector.CloseConnection();

            return content;
        }

        private String MountContractInfo(int contractId, Boolean alertWhenCompleted)
        {
            String content = "";

            dataConnector.OpenConnection("both");
            // Obtem os dados do contrato
            ContractDAO contractDAO = new ContractDAO(dataConnector.MySqlConnection);
            ContractDTO contract = contractDAO.GetContract(contractId);
            // Obtem os subcontratos pertencentes ao contrato
            SubContractDAO subContractDAO = new SubContractDAO(dataConnector.MySqlConnection);
            List<SubContractDTO> subContractList = subContractDAO.GetSubContracts("contrato_id=" + contractId);
            // Obtem os dados do cliente
            BusinessPartnerDAO businessPartnerDAO = new BusinessPartnerDAO(dataConnector.SqlServerConnection);
            BusinessPartnerDTO businessPartner = businessPartnerDAO.GetBusinessPartner(contract.pn);
            // Cria os objetos auxiliares de ORM
            ContractItemDAO contractItemDAO = new ContractItemDAO(dataConnector.MySqlConnection);
            EquipmentDAO equipmentDAO = new EquipmentDAO(dataConnector.SqlServerConnection);

            content += "<h3>Contrato: " + contract.numero + "</h3>";
            content += "<h3>Cliente: " + businessPartner.CardName + "</h3>";
            content += "<h3>Parcela Atual: " + contract.parcelaAtual + " Quant. Parcelas: " + contract.quantidadeParcelas + "</h3>";
            foreach (SubContractDTO subContract in subContractList)
            {
                // Obtem os equipamentos pertencentes ao subcontrato
                List<ContractItemDTO> itemList = contractItemDAO.GetItems("subcontrato_id = " + subContract.id);
                String equipmentEnumeration = "";
                foreach (ContractItemDTO contractItem in itemList)
                {
                    if (!String.IsNullOrEmpty(equipmentEnumeration)) equipmentEnumeration += ", ";
                    equipmentEnumeration += contractItem.codigoCartaoEquipamento;
                }
                // Obtem os respectivos números de série
                List<EquipmentDTO> equipamentList = equipmentDAO.GetEquipments(equipmentEnumeration);
                String serialNumbers = "";
                foreach (EquipmentDTO equipment in equipamentList)
                {
                    if (!String.IsNullOrEmpty(serialNumbers)) serialNumbers += ", ";
                    serialNumbers += equipment.ManufSN;
                }
                if (String.IsNullOrEmpty(serialNumbers)) serialNumbers = "Nenhum item encontrado";

                content += "<h3>" + subContract.siglaTipo + " - " + serialNumbers + "</h3>";
            }
            dataConnector.CloseConnection();

            if (alertWhenCompleted) sboApplication.MessageBox("Contrato: " + contract.numero.PadRight(10, ' ') + "(SYS ID: " + contract.id + ")", 0, "OK", null, null);
            return content;
        }

        private String MountClientContractsInfo(String cardCode, Boolean alertWhenCompleted)
        {
            String content = "";

            dataConnector.OpenConnection("both");
            // Obtem os dados do cliente
            BusinessPartnerDAO businessPartnerDAO = new BusinessPartnerDAO(dataConnector.SqlServerConnection);
            BusinessPartnerDTO businessPartner = businessPartnerDAO.GetBusinessPartner(cardCode);
            // Cria os objetos auxiliares de ORM
            ContractDAO contractDAO = new ContractDAO(dataConnector.MySqlConnection);
            SubContractDAO subContractDAO = new SubContractDAO(dataConnector.MySqlConnection);
            EquipmentDAO equipmentDAO = new EquipmentDAO(dataConnector.SqlServerConnection);
            // Obtem os itens de contrato pertencentes ao cliente, agrupa-os por subcontrato
            Dictionary<int, List<ContractItemDTO>> itemGroups = GetContractItemGroups(dataConnector, cardCode);

            content += "<h3>Cliente: " + businessPartner.CardName + "</h3>";
            // Para cada grupo monta as informações de contrato
            foreach (int subContractId in itemGroups.Keys)
            {
                List<ContractItemDTO> group = itemGroups[subContractId];
                String itemEnumeration = "";
                foreach (ContractItemDTO contractItem in group)
                {
                    if (!String.IsNullOrEmpty(itemEnumeration)) itemEnumeration += ", ";
                    itemEnumeration += contractItem.codigoCartaoEquipamento;
                }
                List<EquipmentDTO> equipamentList = equipmentDAO.GetEquipments(itemEnumeration);
                String equipmentEnumeration = "";
                foreach (EquipmentDTO equipment in equipamentList)
                {
                    if (!String.IsNullOrEmpty(equipmentEnumeration)) equipmentEnumeration += ", ";
                    equipmentEnumeration += equipment.ManufSN;
                }
                SubContractDTO subContract = subContractDAO.GetSubContract(subContractId);
                ContractDTO contract = contractDAO.GetContract(subContract.contrato_id);
                content += "<br/><h3>Contrato: " + contract.numero + "<br/>" + "Parcela Atual: " + contract.parcelaAtual + " Quant. Parcelas: " + contract.quantidadeParcelas + "</h3>";
                content += "<h3>" + subContract.siglaTipo + " - " + equipmentEnumeration + "</h3>";
            }
            dataConnector.CloseConnection();

            if (alertWhenCompleted) sboApplication.MessageBox("Este faturamento agrupa todos os equipamentos do cliente independente do contrato.", 0, "OK", null, null);
            return content;
        }

        private void AddParcel(String formUID)
        {
            SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            if (targetForm == null) return;

            SAPbouiCOM.Item textItem = targetForm.Items.Item("4");
            SAPbouiCOM.EditText textSpecific = (SAPbouiCOM.EditText)textItem.Specific;
            String cardCode = textSpecific.Value;
            if (String.IsNullOrEmpty(cardCode))
            {
                sboApplication.MessageBox("Escolha o cliente primeiro!", 0, "OK", null, null);
                return;
            }

            SAPbouiCOM.ComboBox contractRef = GetSBOComboBox(targetForm, "cmbContrct");
            if (String.IsNullOrEmpty(contractRef.Value))
            {
                sboApplication.MessageBox("Escolha o contrato!", 0, "OK", null, null);
                return;
            }
            int contractId = int.Parse(contractRef.Value);

            // Caso seja apenas um contrato, executa a alteração e retorna
            if (contractId != 0)
            {
                dataConnector.OpenConnection("mySql");
                ContractDAO singleContractDAO = new ContractDAO(dataConnector.MySqlConnection);
                ContractDTO contract = singleContractDAO.GetContract(contractId);
                singleContractDAO.SetContractParcell(contract.id, contract.parcelaAtual + 1);
                dataConnector.CloseConnection();
                sboApplication.MessageBox("Parcela adicionada ao contrato nº " + contract.numero, 0, "OK", null, null);
                return;
            }

            // Caso contrário são todos os contratos do cliente
            dataConnector.OpenConnection("mySql");
            List<int> contractList = new List<int>();
            ContractDAO contractDAO = new ContractDAO(dataConnector.MySqlConnection);
            SubContractDAO subContractDAO = new SubContractDAO(dataConnector.MySqlConnection);
            Dictionary<int, List<ContractItemDTO>> itemGroups = GetContractItemGroups(dataConnector, cardCode);
            foreach (int subContractId in itemGroups.Keys)
            {
                SubContractDTO subContract = subContractDAO.GetSubContract(subContractId);
                if (!contractList.Contains(subContract.contrato_id))
                    contractList.Add(subContract.contrato_id);
            }
            String contractNumbers = "";
            foreach (int id in contractList)
            {
                ContractDTO contract = contractDAO.GetContract(id);
                contractDAO.SetContractParcell(contract.id, contract.parcelaAtual + 1);

                if (!String.IsNullOrEmpty(contractNumbers)) contractNumbers += ", ";
                contractNumbers += "nº " + contract.numero;
            }
            dataConnector.CloseConnection();
            sboApplication.MessageBox("Parcela adicionada aos contratos " + contractNumbers, 0, "OK", null, null);
        }

        // Obtem os itens de contrato pertencentes ao cliente, agrupa-os por subcontrato
        private Dictionary<int, List<ContractItemDTO>> GetContractItemGroups(DataConnector dataConnector, String cardCode)
        {
            Dictionary<int, List<ContractItemDTO>> itemGroups = new Dictionary<int, List<ContractItemDTO>>();

            ContractItemDAO contractItemDAO = new ContractItemDAO(dataConnector.MySqlConnection);
            List<ContractItemDTO> itemList = contractItemDAO.GetItems("businessPartnerCode = '" + cardCode + "'");
            foreach (ContractItemDTO item in itemList)
            {
                // Cria um novo grupo caso não encontre um grupo para este subcontrato
                if (!itemGroups.ContainsKey(item.subContrato_id))
                    itemGroups.Add(item.subContrato_id, new List<ContractItemDTO>());

                // Adiciona o item ao grupo do subcontrato
                List<ContractItemDTO> group = itemGroups[item.subContrato_id];
                group.Add(item);
            }

            return itemGroups;
        }

        private void EquipmentComboClick(String formUID)
        {
            SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            if (targetForm == null) return;

            SAPbouiCOM.Item textItem = targetForm.Items.Item("4");
            SAPbouiCOM.EditText textSpecific = (SAPbouiCOM.EditText)textItem.Specific;
            String cardCode = textSpecific.Value;
            if (String.IsNullOrEmpty(cardCode))
            {
                sboApplication.MessageBox("Escolha o cliente primeiro!", 0, "OK", null, null);
                return;
            }

            // Verifica se o combobox já foi populado
            String lastCardCode = "-";
            UserDataSource lastCardCodeDs = GetSBOUserDataSource(targetForm, "bpCardCode");
            if (lastCardCodeDs != null) lastCardCode = lastCardCodeDs.Value;
            if (lastCardCode == cardCode) return; // O combo já foi populado, termina a execução

            UserDataSource cardCodeDs = GetSBOUserDataSource(targetForm, "bpCardCode");
            if (cardCodeDs == null) cardCodeDs = targetForm.DataSources.UserDataSources.Add("bpCardCode", BoDataType.dt_SHORT_TEXT, 30);
            cardCodeDs.Value = cardCode;
            SAPbouiCOM.ComboBox equipamento = GetSBOComboBox(targetForm, "cmbEqpment");

            // Preenche o combo de escolha de equipamento
            dataConnector.OpenConnection("sqlServer");
            dataConnector.SqlServerConnection.ChangeDatabase(sboApplication.Company.DatabaseName);
            EquipmentDAO equipmentDAO = new EquipmentDAO(dataConnector.SqlServerConnection);
            List<EquipmentDTO> equipmentList = equipmentDAO.GetCustomerEquipments(cardCode);
            dataConnector.CloseConnection();

            // Remove todos os items do combo
            while (equipamento.ValidValues.Count > 0)
                equipamento.ValidValues.Remove(equipamento.ValidValues.Count - 1, BoSearchKey.psk_Index);

            List<String> serialNumbers = new List<String>();
            foreach (EquipmentDTO equipment in equipmentList) // Adiciona os items recuperados da consulta ao banco
            {
                if (!serialNumbers.Contains(equipment.ManufSN))
                {
                    serialNumbers.Add(equipment.ManufSN);
                    equipamento.ValidValues.Add(equipment.ManufSN, "(" + equipment.InternalSN + ") " + EquipmentDAO.GetStatusDescription(equipment.Status));
                }
            }
        }

        private void AddEquipment(String formUID)
        {
            SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            if (targetForm == null) return;

            SAPbouiCOM.ComboBox equipamento = GetSBOComboBox(targetForm, "cmbEqpment");
            SAPbouiCOM.EditText listaEquipamentos = GetSBOEditText(targetForm, "txtLista");

            String[] listNumbers = listaEquipamentos.Value.Split(new Char[] {','} );
            Boolean alreadyThere = false;
            foreach (String number in listNumbers)
                if (number.Trim() == equipamento.Selected.Value) alreadyThere = true;
            if (alreadyThere)
            {
                sboApplication.MessageBox("Este número já está na lista!", 0, "OK", null, null);
                return;
            }

            if (!String.IsNullOrEmpty(listaEquipamentos.Value)) listaEquipamentos.Value += ", ";
            listaEquipamentos.Value += " " + equipamento.Selected.Value;
        }

        private void OpenBillingFilter(String formUID)
        {
            SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            if (targetForm == null) return;

            try
            {
                SAPbouiCOM.Form frmOptions = CreateSBOForm("frmOptions", "Opções de Verificação", 260, 200);
                frmOptions.DataSources.UserDataSources.Add("dsStartDt", BoDataType.dt_DATE, 8);
                frmOptions.DataSources.UserDataSources.Add("dsEndDt", BoDataType.dt_DATE, 8);
                SAPbouiCOM.EditText txtStartDate = AddSBOTextField(frmOptions, "StartDt", "Data Inicial (NFs)", 25, 25, 0);
                txtStartDate.DataBind.SetBound(true, "", "dsStartDt");
                SAPbouiCOM.EditText txtEndDate = AddSBOTextField(frmOptions, "EndDt", "Data Final (NFs)", 25, 50, 0);
                txtEndDate.DataBind.SetBound(true, "", "dsEndDt");
                SAPbouiCOM.EditText txtMonth = AddSBOTextField(frmOptions, "Month", "Mês de Faturamento", 25, 75, 0);
                SAPbouiCOM.EditText txtYear = AddSBOTextField(frmOptions, "Year", "Ano de Faturamento", 25, 100, 0);
                SAPbouiCOM.Button btnOK = AddSBOButton(frmOptions, "Ok", "Ok", 80, 140, 0);
                frmOptions.Items.Item("txtStartDt").Click(BoCellClickType.ct_Regular); // posiciona o cursor no primeiro campo
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        private void CheckBillings(String formUID)
        {
            SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            if (targetForm == null) return;

            try
            {
                SAPbouiCOM.EditText txtStartDate = GetSBOEditText(targetForm, "txtStartDt");
                SAPbouiCOM.EditText txtEndDate = GetSBOEditText(targetForm, "txtEndDt");
                SAPbouiCOM.EditText txtMonth = GetSBOEditText(targetForm, "txtMonth");
                SAPbouiCOM.EditText txtYear = GetSBOEditText(targetForm, "txtYear");

                System.Globalization.CultureInfo provider = System.Globalization.CultureInfo.InvariantCulture;
                DateTime dtStartDate = DateTime.ParseExact(txtStartDate.Value, "yyyyMMdd", provider);
                DateTime dtEndDate = DateTime.ParseExact(txtEndDate.Value, "yyyyMMdd", provider);
                String startDate = String.Format("{0:dd/MM/yyyy}", dtStartDate);
                String endDate = String.Format("{0:dd/MM/yyyy}", dtEndDate);
                String month = txtMonth.Value;
                String year = txtYear.Value;
                targetForm.Close();

                int num;
                Boolean monthIsNumeric = int.TryParse(month, out num);
                Boolean yearIsNumeric = int.TryParse(year, out num);
                if (!monthIsNumeric)
                {
                    sboApplication.MessageBox("O mês deve ser um número", 0, "OK", "", "");
                    return;
                }
                if (!yearIsNumeric)
                {
                    sboApplication.MessageBox("O ano deve ser um número", 0, "OK", "", "");
                    return;
                }

                dataConnector.OpenConnection("both");
                InvoiceDAO invoiceDAO = new InvoiceDAO(dataConnector.SqlServerConnection);
                List<InvoiceDTO> startedInvoiceList = invoiceDAO.GetAllInvoices("OINV.DocDate BETWEEN '" + startDate + "' AND '" + endDate + "' AND INV1.ItemCode = 'S001'");
                List<InvoiceDTO> returnedInvoiceList = invoiceDAO.GetReturnedInvoices("(T0.DocDate BETWEEN '" + startDate + "' AND '" + endDate + "' OR T3.DocDate BETWEEN '" + startDate + "' AND '" + endDate + "') AND T3.DocTotal > 0");
                decimal grandTotal = 0;
                List<InvoiceDTO> invoiceList = new List<InvoiceDTO>();
                foreach (InvoiceDTO startedInvoice in startedInvoiceList)
                {
                    if (!returnedInvoiceList.Contains(startedInvoice))
                    {
                        grandTotal += startedInvoice.docTotal;
                        invoiceList.Add(startedInvoice);
                    }
                }
                // Cuidado ao exibir valores de faturamento no sistema, pode ser exposto a funcionários não autorizados
                // sboApplication.MessageBox("Total Faturamento: " + String.Format("{0:0.00}", grandTotal), 0, "OK", "", "");

                List<BillingDTO> billingFaults = new List<BillingDTO>();
                List<InvoiceDTO> invoiceFaults = new List<InvoiceDTO>();
                BillingDAO billingDAO = new BillingDAO(dataConnector.MySqlConnection);
                List<BillingDTO> billingList = billingDAO.GetAllBillings("mesReferencia = '" + month + "' AND anoReferencia = '" + year + "'");
                // Varre os demonstrativos de faturamento procurando correspondências nas faturas
                foreach (BillingDTO billing in billingList)
                {
                    decimal billingTotal = (decimal)billing.total + (decimal)billing.acrescimoDesconto;
                    Boolean matching = false;
                    foreach (InvoiceDTO invoice in invoiceList)
                    {
                        if (invoice.demFaturamento != null)
                        {
                            if (invoice.demFaturamento.Value == billing.id)
                            {
                                decimal diff = Math.Abs(invoice.docTotal - billingTotal);
                                if (diff < 1) matching = true;
                            }
                        }
                    }
                    if (!matching) billingFaults.Add(billing);
                }
                // Varre as faturas procurando correspondências nos demonstrativos de faturamento (caminho inverso)
                foreach (InvoiceDTO invoice in invoiceList)
                {
                    Boolean matching = false;
                    foreach (BillingDTO billing in billingList)
                    {
                        decimal billingTotal = (decimal)billing.total + (decimal)billing.acrescimoDesconto;
                        if (invoice.demFaturamento != null)
                        {
                            if (invoice.demFaturamento.Value == billing.id)
                            {
                                decimal diff = Math.Abs(invoice.docTotal - billingTotal);
                                if (diff < 1) matching = true;
                            }
                        }
                    }
                    if (!matching) invoiceFaults.Add(invoice);
                }
                // Faz uma varredura complementar para os casos onde existem correspondências 1 para n
                foreach (BillingDTO billing in billingFaults.ToArray())
                {
                    List<InvoiceDTO> invoiceMatches = invoiceFaults.FindAll(delegate(InvoiceDTO invoice) { return invoice.demFaturamento == billing.id; });
                    decimal billingTotal = (decimal)billing.total + (decimal)billing.acrescimoDesconto;
                    decimal invoiceTotal = 0;
                    foreach (InvoiceDTO invoice in invoiceMatches)
                    {
                        invoiceTotal += invoice.docTotal;
                    }
                    decimal diff = Math.Abs(billingTotal - invoiceTotal);
                    if (diff < 1)
                    {
                        // remove da lista as correspondências 1 para n
                        billingFaults.Remove(billing);
                        invoiceFaults.RemoveAll(delegate(InvoiceDTO invoice) { return invoice.demFaturamento == billing.id; });
                    }
                }

                foreach (BillingDTO billing in billingFaults)
                    sboApplication.MessageBox("Verificar demonstrativo nº " + billing.id, 0, "OK", "", "");
                foreach (InvoiceDTO invoice in invoiceFaults)
                    sboApplication.MessageBox("Verificar NF " + invoice.docNum + " (" + invoice.comments + ")", 0, "OK", "", "");

                dataConnector.CloseConnection();

                sboApplication.MessageBox("Verificação concluida.", 0, "OK", "", "");
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }


        private void CreateItemUserFields(String formUID)
        {
            SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            if (targetForm == null) return;

            SAPbouiCOM.EditText durability = AddSBOTextField(targetForm, "durblty", "Durabilidade/Vida Útil", 400, 85, 0);
            durability.DataBind.SetBound(true, "OITM", "U_Durability");
        }

        private void SetEmailAddress(String formUID)
        {
            SAPbouiCOM.Form activeForm = sboApplication.Forms.ActiveForm;
            if (activeForm.Type != 133) return;

            SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            if (targetForm == null) return;

            SAPbouiCOM.Item textItem = activeForm.Items.Item("4");
            SAPbouiCOM.EditText textSpecific = (SAPbouiCOM.EditText)textItem.Specific;
            String cardCode = textSpecific.Value;
            UserDataSource cardCodeDs = targetForm.DataSources.UserDataSources.Add("cardCodeDs", BoDataType.dt_SHORT_TEXT, 20);
            cardCodeDs.Value = cardCode;

            // Realiza a operação em uma thread separada para evitar o congelamento da interface
            BackgroundWorker backgroundWorker = new BackgroundWorker();
            backgroundWorker.DoWork += new DoWorkEventHandler(UpdateEmailAddressField);
            backgroundWorker.RunWorkerAsync(targetForm);

            System.Windows.Forms.Application.DoEvents();
        }

        private void UpdateEmailAddressField(Object sender, DoWorkEventArgs e)
        {
            SAPbouiCOM.Form targetForm = (SAPbouiCOM.Form)e.Argument;
            Thread.Sleep(1500); // aguarda pouco mais de um segundo até o carregamento completo do form
            try
            {
                UserDataSource cardCodeDs = targetForm.DataSources.UserDataSources.Item("cardCodeDs");
                dataConnector.OpenConnection("sqlServer");
                dataConnector.SqlServerConnection.ChangeDatabase(sboApplication.Company.DatabaseName);
                BusinessPartnerDAO bpDAO = new BusinessPartnerDAO(dataConnector.SqlServerConnection);
                BusinessPartnerDTO businessPartner = bpDAO.GetBusinessPartner(cardCodeDs.Value);
                PartnerContactDAO contactDAO = new PartnerContactDAO(dataConnector.SqlServerConnection);
                PartnerContactDTO contact = contactDAO.GetContact(businessPartner.CardCode, businessPartner.CntctPrsn);
                dataConnector.CloseConnection();

                SAPbouiCOM.Item textItem = targetForm.Items.Item("txtPara");
                SAPbouiCOM.EditText textSpecific = (SAPbouiCOM.EditText)textItem.Specific;
                textSpecific.Value = contact.Email;
                targetForm.DataSources.UserDataSources.Item(6).Value = contact.Email;
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        private void CreatePaymentLoader(String formUID)
        {
            SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            if (targetForm == null) return;

            try
            {
                SAPbouiCOM.Item btnLoad = targetForm.Items.Add("btnLoad", BoFormItemTypes.it_BUTTON);
                btnLoad.Top = 100;
                btnLoad.Left = 270;
                btnLoad.Width = 120;
                btnLoad.Height = 30;
                ((SAPbouiCOM.Button)btnLoad.Specific).Caption = "Carregar Pagamentos";
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        private void LoadPayments(String formUID)
        {
            SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            if (targetForm == null) return;

            try
            {
                dataConnector.OpenConnection("sqlServer");
                dataConnector.SqlServerConnection.ChangeDatabase(sboApplication.Company.DatabaseName);
                BoeTransactionDAO transactionDAO = new BoeTransactionDAO(dataConnector.SqlServerConnection);
                List<BoeTransactionDTO> paymentList = transactionDAO.GetBoePayments();
                JournalEntryDAO journalEntryDAO = new JournalEntryDAO(dataConnector.SqlServerConnection);

                foreach (BoeTransactionDTO payment in paymentList)
                {
                    // Obtem os créditos para o boleto pago (valores, juros e taxas)
                    Decimal boeOverdueFine = 0;
                    Decimal boeBankingFee = 0;
                    List<JournalEntryDTO> boeCredits = journalEntryDAO.GetBoeCredits(payment.boeNumber);
                    foreach (JournalEntryDTO credit in boeCredits)
                    {
                        if (credit.memo.ToUpper().Contains("JUROS")) boeOverdueFine += credit.SysTotal;
                        if (credit.memo.ToUpper().Contains("TAXA")) boeBankingFee += credit.SysTotal;
                    }

                    String query = "UPDATE OBOE SET U_PaymentDate = @PaymentDate, U_OverdueFine = @OverdueFine, U_ReceivedAmount = @ReceivedAmount WHERE boeNum = " + payment.boeNumber;
                    SqlParameter param1 = new SqlParameter("@PaymentDate", System.Data.SqlDbType.DateTime);
                    param1.Value = payment.paymentDate;
                    SqlParameter param2 = new SqlParameter("@OverdueFine", System.Data.SqlDbType.Decimal);
                    param2.Value = boeOverdueFine;
                    SqlParameter param3 = new SqlParameter("@ReceivedAmount", System.Data.SqlDbType.Decimal);
                    param3.Value = payment.boeSum + boeOverdueFine;

                    SqlCommand command = new SqlCommand(query, dataConnector.SqlServerConnection);
                    command.Parameters.Add(param1);
                    command.Parameters.Add(param2);
                    command.Parameters.Add(param3);
                    command.ExecuteNonQuery();
                }
                dataConnector.CloseConnection();

                sboApplication.MessageBox("Pagamentos carregados com exito", 0, "", "", "");
            }
            catch (Exception exc)
            {
                lastError = "Erro ao carregar pagamentos" + exc.Message;
            }
        }

        private void CreateReportButton(String formUID)
        {
            SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            if (targetForm == null) return;

            try
            {
                SAPbouiCOM.Item btnReport = targetForm.Items.Add("btnReport", BoFormItemTypes.it_BUTTON);
                btnReport.Top = 100;
                btnReport.Left = 450;
                btnReport.Width = 120;
                btnReport.Height = 30;
                ((SAPbouiCOM.Button)btnReport.Specific).Caption = "Relatório ( Diário )";
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        private void CreateEditButton(String formUID)
        {
            SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            if (targetForm == null) return;

            try
            {
                SAPbouiCOM.Item btnEdit = targetForm.Items.Add("btnEdit", BoFormItemTypes.it_BUTTON);
                btnEdit.Top = 100;
                btnEdit.Left = 600;
                btnEdit.Width = 120;
                btnEdit.Height = 30;
                ((SAPbouiCOM.Button)btnEdit.Specific).Caption = "Editar Relatório";
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        private void SelectReport(String formUID)
        {
            SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            if (targetForm == null) return;

            try
            {
                // Recupera os relatórios em formato Crystal Reports ( Categoria C )
                dataConnector.OpenConnection("sqlServer");
                dataConnector.SqlServerConnection.ChangeDatabase(sboApplication.Company.DatabaseName);
                ReportingDocumentDAO reportingDocumentDAO = new ReportingDocumentDAO(dataConnector.SqlServerConnection);
                List<ReportingDocumentDTO> reportList = reportingDocumentDAO.GetReports("Category = 'C'");
                dataConnector.CloseConnection();

                // Cria um formulário de escolha de relatório
                SAPbouiCOM.Form frmEdit = CreateSBOForm("frmEdit", "Editar relatório", 260, 200);
                SAPbouiCOM.ComboBox cmbReport = AddSBOCombobox(frmEdit, "Report", "Relatório", 25, 25, 0);
                foreach (ReportingDocumentDTO report in reportList)
                {
                    cmbReport.ValidValues.Add(report.DocCode, report.DocName);
                }
                SAPbouiCOM.Button btnOK = AddSBOButton(frmEdit, "Ok", "Ok", 80, 120, 0);
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        private void EditCrystalReport(String formUID)
        {
            SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            if (targetForm == null) return;

            try
            {
                // Recupera o relatório selecionado no combo
                SAPbouiCOM.ComboBox cmbReport = GetSBOComboBox(targetForm, "cmbReport");
                String selectedReport = cmbReport.Selected.Value;
                String fileName = cmbReport.Selected.Description + ".rpt";

                // Define o diretório de gravação
                String appDataFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                String reportFolder = Path.Combine(appDataFolder, "OsiReports");
                Directory.CreateDirectory(reportFolder);

                // Exporta o relatório e abre o arquivo exportado
                String filePath = Path.Combine(reportFolder, fileName);
                ExportCrystalReport(selectedReport, filePath);
                targetForm.Close();
                System.Diagnostics.Process.Start(filePath);
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        private void OpenReportFilter(String formUID)
        {
            SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            if (targetForm == null) return;

            try
            {
                SAPbouiCOM.Form frmFilter = CreateSBOForm("frmFilter", "Filtro de relatório", 260, 200);
                frmFilter.DataSources.UserDataSources.Add("dsStartDt", BoDataType.dt_DATE, 8);
                frmFilter.DataSources.UserDataSources.Add("dsEndDt", BoDataType.dt_DATE, 8);
                SAPbouiCOM.EditText txtStartDate = AddSBOTextField(frmFilter, "StartDt", "Data Inicial", 25, 25, 0);
                txtStartDate.DataBind.SetBound(true, "", "dsStartDt");
                SAPbouiCOM.EditText txtEndDate = AddSBOTextField(frmFilter, "EndDt", "Data Final", 25, 50, 0);
                txtEndDate.DataBind.SetBound(true, "", "dsEndDt");
                SAPbouiCOM.EditText txtStartPage = AddSBOTextField(frmFilter, "Page", "Página Inicial", 25, 75, 0);
                SAPbouiCOM.EditText txtSkipped = AddSBOTextField(frmFilter, "Skipped", "Páginas Puladas", 25, 100, 0);
                SAPbouiCOM.Button btnOK = AddSBOButton(frmFilter, "Ok", "Ok", 80, 140, 0);
                frmFilter.Items.Item("txtStartDt").Click(BoCellClickType.ct_Regular);
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        private void CreateContractReportButton(String formUID)
        {
            SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            if (targetForm == null) return;

            try
            {
                SAPbouiCOM.Item btnLoad = targetForm.Items.Add("btnReport", BoFormItemTypes.it_BUTTON);
                btnLoad.Top = 25;
                btnLoad.Left = 360;
                btnLoad.Width = 120;
                btnLoad.Height = 30;
                ((SAPbouiCOM.Button)btnLoad.Specific).Caption = "Gerar Relatório";
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        private void ContractReportButtonClick(String formUID)
        {
            SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            if (targetForm == null) return;

            try
            {
                String appDataFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                String reportFolder = Path.Combine(appDataFolder, "OsiReports");
                Directory.CreateDirectory(reportFolder);
                String filename = Path.Combine(reportFolder, "Contratos.htm");

                FileStream fileStream = new FileStream(filename, FileMode.Create);
                StreamWriter streamWriter = new StreamWriter(fileStream);

                streamWriter.WriteLine("<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Strict//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd'>");
                streamWriter.WriteLine("<html xmlns='http://www.w3.org/1999/xhtml'>");
                streamWriter.WriteLine("<head>");
                streamWriter.WriteLine("    <meta http-equiv='Content-type' content='text/html; charset=UTF-8' />");
                streamWriter.WriteLine("    <meta http-equiv='Content-Language' content='pt-br' />");
                streamWriter.WriteLine("    <style type='text/css'>");
                streamWriter.WriteLine("        table{  border-left:1px solid black; border-top:1px solid black; width:100%; margin:0; padding:0; border-spacing:0;  }");
                streamWriter.WriteLine("        td{  border-right:1px solid black; border-bottom:1px solid black; margin:0; padding:0; text-align:center;  }");
                streamWriter.WriteLine("    </style>");
                streamWriter.WriteLine("    <title>Relatório</title>");
                streamWriter.WriteLine("</head>");
                streamWriter.WriteLine("<body style='font-size: 11px;'>");
                streamWriter.WriteLine("<div width='80%'><h1>SITUAÇÃO DOS CONTRATOS</h1>");
                streamWriter.WriteLine("<table>");
                String tableHeader = "<tr><td>Contrato</td><td>Status</td><td>Cliente</td><td>Parcela</td></tr>";
                streamWriter.WriteLine(tableHeader);

                int itemsOnPage = 0;
                int pageCount = 1;
                dataConnector.OpenConnection("both");
                ContractDAO contractDAO = new ContractDAO(dataConnector.MySqlConnection);
                BusinessPartnerDAO businessPartnerDAO = new BusinessPartnerDAO(dataConnector.SqlServerConnection);

                List<ContractDTO> contractList = contractDAO.GetAllContracts(null);
                foreach(ContractDTO contract in contractList) {
                    BusinessPartnerDTO businessPartner = businessPartnerDAO.GetBusinessPartner(contract.pn);
                    String clientName = contract.pn + " - " + businessPartner.CardName;
                    if (!String.IsNullOrEmpty(contract.divisao)) clientName = clientName + " - " + contract.divisao;
                    String parcela = contract.parcelaAtual + "/" + contract.quantidadeParcelas;

                    streamWriter.WriteLine("<tr><td>" + contract.numero.PadLeft(5, '0') + "</td><td>" + ContractDAO.GetStatusAsText(contract.status) + "</td><td>" + clientName + "</td><td>" + parcela + "</td></tr>");
                    InsertNewPageIfNeeded("SITUAÇÃO DOS CONTRATOS", tableHeader, ref itemsOnPage, ref pageCount, 0, 0, streamWriter);
                }

                dataConnector.CloseConnection();

                streamWriter.WriteLine("</table></div><br/><hr/>");
                streamWriter.WriteLine("<div style='width: 100%; text-align: center;'><span style='font-weight: bold; margin:0px auto;' >" + pageCount + "</span></div>");
                streamWriter.WriteLine("</body>");
                streamWriter.WriteLine("</html>");
                streamWriter.Close();
                System.Diagnostics.Process.Start(filename);
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        private void BuildTagSheet(AddressDTO address)
        {
            String appDataFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            String reportFolder = Path.Combine(appDataFolder, "OsiReports");
            Directory.CreateDirectory(reportFolder);
            String filename = Path.Combine(reportFolder, "Etiquetas.htm");

            FileStream fileStream = new FileStream(filename, FileMode.Create);
            StreamWriter streamWriter = new StreamWriter(fileStream);

            try
            {
                streamWriter.WriteLine("<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Strict//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd'>");
                streamWriter.WriteLine("<html xmlns='http://www.w3.org/1999/xhtml'>");
                streamWriter.WriteLine("<head>");
                streamWriter.WriteLine("    <meta http-equiv='Content-type' content='text/html; charset=UTF-8' />");
                streamWriter.WriteLine("    <meta http-equiv='Content-Language' content='pt-br' />");
                streamWriter.WriteLine("    <style type='text/css'>");
                streamWriter.WriteLine("        table{  border-left:1px solid LightGray; border-top:1px solid LightGray; width:100%; margin:0; padding:0; border-spacing:0;  }");
                streamWriter.WriteLine("        td{  border-right:1px solid LightGray; border-bottom:1px solid LightGray; margin:0; padding:0; text-align:left;  }");
                streamWriter.WriteLine("    </style>");
                streamWriter.WriteLine("    <title>Etiquetas</title>");
                streamWriter.WriteLine("</head>");
                streamWriter.WriteLine("<body style='font-size: 11px;'>");
                streamWriter.WriteLine("<div style='width:95%; margin-left:auto; margin-right:auto;' >");
                streamWriter.WriteLine("<table>");

                dataConnector.OpenConnection("sqlServer");
                dataConnector.SqlServerConnection.ChangeDatabase(sboApplication.Company.DatabaseName);
                BusinessPartnerDAO businessPartnerDAO = new BusinessPartnerDAO(dataConnector.SqlServerConnection);
                BusinessPartnerDTO businessPartner = businessPartnerDAO.GetBusinessPartner(address.CardCode);
                String contact = "";
                if (!String.IsNullOrEmpty(businessPartner.CntctPrsn))
                {
                    PartnerContactDAO partnerContactDAO = new PartnerContactDAO(dataConnector.SqlServerConnection);
                    PartnerContactDTO contactPerson = partnerContactDAO.GetContact(businessPartner.CardCode, businessPartner.CntctPrsn);
                    contact = "Contato: " + contactPerson.Name;
                }

                int columnCount = 2;
                int rowCount = 8;
                for(int y = 1; y <= rowCount; y++)
                {
                    String tag = "<tr>";
                    for(int x = 1; x <= columnCount; x ++)
                    {
                        String customer = "<b>" + businessPartner.CardName + "</b>";
                        String locale1 = "Endereço: " + address.AddrType + " " + address.Street + ", " + address.StreetNo + "   " + address.Building + "   " + contact;
                        String locale2 = "CEP: " + address.ZipCode + "   " + "Bairro: " + address.Block + "   " + address.City + " " + address.State + " " + address.Country;
                        tag += "<td style='height:120px;' ><div style='width:90%; margin-left:auto; margin-right:auto;' >" + customer + "<br/><br/>" + locale1 + "<br/><br/>" + locale2 + "</div></td>";
                    }
                    tag += "</tr>";
                    streamWriter.WriteLine(tag);
                }

                dataConnector.CloseConnection();

                streamWriter.WriteLine("</table>");
                streamWriter.WriteLine("</div>");
                streamWriter.WriteLine("</body>");
                streamWriter.WriteLine("</html>");
                streamWriter.Close();
                System.Diagnostics.Process.Start(filename);
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        private void BuildAccountTreeReport(String formUID)
        {
            SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            if (targetForm == null) return;

            try
            {
                SAPbouiCOM.EditText txtPage = GetSBOEditText(targetForm, "txtPage");
                String startPage = txtPage.Value;

                String appDataFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                String reportFolder = Path.Combine(appDataFolder, "OsiReports");
                Directory.CreateDirectory(reportFolder);
                String filename = Path.Combine(reportFolder, "PlanoContas.htm");

                FileStream fileStream = new FileStream(filename, FileMode.Create);
                StreamWriter streamWriter = new StreamWriter(fileStream);

                streamWriter.WriteLine("<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Strict//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd'>");
                streamWriter.WriteLine("<html xmlns='http://www.w3.org/1999/xhtml'>");
                streamWriter.WriteLine("<head>");
                streamWriter.WriteLine("    <meta http-equiv='Content-type' content='text/html; charset=UTF-8' />");
                streamWriter.WriteLine("    <meta http-equiv='Content-Language' content='pt-br' />");
                streamWriter.WriteLine("    <style type='text/css'>");
                streamWriter.WriteLine("        table{  border-left:1px solid LightGray; border-top:1px solid LightGray; width:100%; margin:0; padding:0; border-spacing:0;  }");
                streamWriter.WriteLine("        td{  border-right:1px solid LightGray; border-bottom:1px solid LightGray; margin:0; padding:0; text-align:left;  }");
                streamWriter.WriteLine("    </style>");
                streamWriter.WriteLine("    <title>Plano de Contas</title>");
                streamWriter.WriteLine("</head>");
                streamWriter.WriteLine("<body style='font-size: 11px;'>");
                streamWriter.WriteLine("<div width='80%'><h1>PLANO DE CONTAS</h1>");
                streamWriter.WriteLine("<table>");

                dataConnector.OpenConnection("sqlServer");
                dataConnector.SqlServerConnection.ChangeDatabase(sboApplication.Company.DatabaseName);
                AccountDAO accountDAO = new AccountDAO(dataConnector.SqlServerConnection);

                List<AccountDTO> accountList = accountDAO.GetLeafAccounts();
                String[] subTitles = new String[] { "Ativo", "Passivo", "Receita", "Despesa", "Custo", "Contas De Compensação" };
                int previousPrefix = 0;
                int itemsOnPage = 0;
                int pageCount = 0;
                int.TryParse(startPage, out pageCount);
                foreach (AccountDTO account in accountList)
                {
                    int accountPrefix = int.Parse(account.acctCode[0].ToString());
                    if (accountPrefix != previousPrefix)
                    {
                        String subTitle = "<tr><td>" + subTitles[accountPrefix - 1] + "</td></tr>";
                        streamWriter.WriteLine(subTitle);
                        previousPrefix = accountPrefix;
                    }
                    String indentation = new System.Text.StringBuilder().Insert(0, "&nbsp", account.level + 1).ToString();
                    String docEntry = "<tr><td>" + indentation + account.acctCode + " - " + account.acctName + "</td></tr>";
                    streamWriter.WriteLine(docEntry);
                    InsertNewPageIfNeeded("PLANO DE CONTAS", null, ref itemsOnPage, ref pageCount, 0, 0, streamWriter);
                }
                dataConnector.CloseConnection();

                streamWriter.WriteLine("</table></div><br/><hr/>");
                streamWriter.WriteLine("<div style='width: 100%; text-align: center;'><span style='font-weight: bold; margin:0px auto;' >" + pageCount + "</span></div>");
                streamWriter.WriteLine("</body>");
                streamWriter.WriteLine("</html>");
                streamWriter.Close();
                System.Diagnostics.Process.Start(filename);
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        private void BuildJournalReport(String formUID)
        {
            SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            if (targetForm == null) return;

            try
            {
                SAPbouiCOM.EditText txtStartDate = GetSBOEditText(targetForm, "txtStartDt");
                SAPbouiCOM.EditText txtEndDate = GetSBOEditText(targetForm, "txtEndDt");
                SAPbouiCOM.EditText txtPage = GetSBOEditText(targetForm, "txtPage");
                SAPbouiCOM.EditText txtSkipped = GetSBOEditText(targetForm, "txtSkipped");
                System.Globalization.CultureInfo provider = System.Globalization.CultureInfo.InvariantCulture;
                DateTime startDate = DateTime.ParseExact(txtStartDate.Value, "yyyyMMdd", provider);
                DateTime endDate = DateTime.ParseExact(txtEndDate.Value, "yyyyMMdd", provider);
                String startPage = txtPage.Value;
                String skippedPages = txtSkipped.Value;
                targetForm.Close();

                String appDataFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                String reportFolder = Path.Combine(appDataFolder, "OsiReports");
                Directory.CreateDirectory(reportFolder);
                String filename = Path.Combine(reportFolder, "Diario.htm");

                FileStream fileStream = new FileStream(filename, FileMode.Create);
                StreamWriter streamWriter = new StreamWriter(fileStream);

                streamWriter.WriteLine("<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Strict//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd'>");
                streamWriter.WriteLine("<html xmlns='http://www.w3.org/1999/xhtml'>");
                streamWriter.WriteLine("<head>");
                streamWriter.WriteLine("    <meta http-equiv='Content-type' content='text/html; charset=UTF-8' />");
                streamWriter.WriteLine("    <meta http-equiv='Content-Language' content='pt-br' />");
                streamWriter.WriteLine("    <style type='text/css'>");
                streamWriter.WriteLine("        table{  border-left:1px solid black; border-top:1px solid black; width:100%; margin:0; padding:0; border-spacing:0;  }");
                streamWriter.WriteLine("        td{  border-right:1px solid black; border-bottom:1px solid black; margin:0; padding:0; text-align:center;  }");
                streamWriter.WriteLine("    </style>");
                streamWriter.WriteLine("    <title>Diário Geral</title>");
                streamWriter.WriteLine("</head>");
                streamWriter.WriteLine("<body style='font-size: 11px;'>");
                streamWriter.WriteLine("<div width='80%'><h1>LIVRO DIÁRIO GERAL</h1>");
                streamWriter.WriteLine("<table>");
                String tableHeader = "<tr><td width='10%'>Data</td><td width='10%'>No Lançamento</td><td width='15%'>Débito</td><td width='15%'>Crédito</td><td width='40%'>Descrição</td><td width='10%'>Valor (R$)</td></tr>";
                streamWriter.WriteLine(tableHeader);

                dataConnector.OpenConnection("sqlServer");
                dataConnector.SqlServerConnection.ChangeDatabase(sboApplication.Company.DatabaseName);
                JournalEntryDAO journalEntryDAO = new JournalEntryDAO(dataConnector.SqlServerConnection);

                List<JournalEntryDTO> entryList = journalEntryDAO.GetEntriesByPeriod(startDate, endDate);
                int itemsOnPage = 0;
                int pageCount = 0;
                int firstPageNum = 0;
                int pagesToSkip = 0;
                int.TryParse(startPage, out firstPageNum);
                int.TryParse(skippedPages, out pagesToSkip);
                foreach (JournalEntryDTO journalEntry in entryList)
                {
                    String docDate = String.Format("{0:dd/MM/yyyy}", journalEntry.refDate);
                    int docNumber = journalEntry.number;
                    String creditAccount = "";
                    String debitAccount = "";
                    Decimal debit = 0;
                    Decimal credit = 0;
                    String description = "";
                    List<JournalEntryItemDTO> itemList = journalEntryDAO.GetItems(journalEntry.transId);
                    Boolean showDetails = false;
                    List<String> details = new List<String>();
                    if (itemList.Count > 2) showDetails = true;
                    foreach (JournalEntryItemDTO item in itemList)
                    {
                        if ((!showDetails) && (item.credit == 0))
                        {
                            debitAccount = item.account;
                            description = item.lineMemo;
                            debit = item.debit;
                        }
                        if ((!showDetails) && (item.debit == 0))
                        {
                            creditAccount = item.account;
                            description = item.lineMemo;
                            credit = item.credit;
                        }

                        if ((showDetails) && (item.credit == 0))
                        {
                            String debitedValue = String.Format("{0:0.00}", item.debit);
                            details.Add("<tr><td></td><td></td><td>" + item.account + "</td><td></td><td>" + item.lineMemo + "</td><td>" + debitedValue + "</td></tr>");
                        }
                        if ((showDetails) && (item.debit == 0))
                        {
                            String creditedValue = String.Format("{0:0.00}", item.credit);
                            details.Add("<tr><td></td><td></td><td></td><td>" + item.account + "</td><td>" + item.lineMemo + "</td><td>" + creditedValue + "</td></tr>");
                        }
                    }

                    String total = String.Format("{0:0.00}", credit);
                    if (showDetails) total = ""; // Não exibe o total, pois ele aparece nos detalhes
                    String docEntry = "<tr><td>" + docDate + "</td><td>" + docNumber + "</td><td>" + debitAccount + "</td><td>" + creditAccount + "</td><td>" + description + "</td><td>" + total + "</td></tr>";
                    if (pageCount >= pagesToSkip) streamWriter.WriteLine(docEntry);
                    InsertNewPageIfNeeded("LIVRO DIÁRIO GERAL", tableHeader, ref itemsOnPage, ref pageCount, firstPageNum, pagesToSkip, streamWriter);
                    foreach (String detail in details)
                    {
                        if (pageCount >= pagesToSkip) streamWriter.WriteLine(detail);
                        InsertNewPageIfNeeded("LIVRO DIÁRIO GERAL", tableHeader, ref itemsOnPage, ref pageCount, firstPageNum, pagesToSkip, streamWriter);
                    }
                }
                dataConnector.CloseConnection();

                streamWriter.WriteLine("</table></div><br/><hr/>");
                streamWriter.WriteLine("<div style='width: 100%; text-align: center;'><span style='font-weight: bold; margin:0px auto;' >" + (pageCount - pagesToSkip + firstPageNum) + "</span></div>");
                streamWriter.WriteLine("</body>");
                streamWriter.WriteLine("</html>");
                streamWriter.Close();
                System.Diagnostics.Process.Start(filename);
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        private void InsertNewPageIfNeeded(String pageTitle, String tableHeader, ref int itemsOnPage, ref int pageCount, int firstPageNum, int pagesToSkip, StreamWriter streamWriter)
        {
            itemsOnPage++;
            if (itemsOnPage > 52) // quantidade de registros por página
            {
                if (pageCount >= pagesToSkip)
                {
                    streamWriter.WriteLine("</table></div><br/><hr/>");
                    streamWriter.WriteLine("<div style='width: 100%; text-align: center; page-break-after: always;'><span style='font-weight: bold; margin:0px auto;' >" + (pageCount - pagesToSkip + firstPageNum) + "</span></div>");
                    streamWriter.WriteLine("<div width='80%'><h1>" + pageTitle + "</h1>");
                    streamWriter.WriteLine("<table>");
                    if (!String.IsNullOrEmpty(tableHeader)) streamWriter.WriteLine(tableHeader);
                }
                itemsOnPage = 0;
                pageCount++;
            }
        }

        private void CreatePaginationIndex(String formUID)
        {
            SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            if (targetForm == null) return;

            try
            {
                SAPbouiCOM.Form frmStartPage = CreateSBOForm("frmStartPage", "Confiruações", 260, 160);
                SAPbouiCOM.EditText txtPage = AddSBOTextField(frmStartPage, "Page", "Página Inicial", 25, 25, 0);
                txtPage.Value = "1";
                SAPbouiCOM.Button btnOK = AddSBOButton(frmStartPage, "Ok", "Ok", 80, 80, 0);
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        private void SavePaginationIndex(String formUID)
        {
            SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            if (targetForm == null) return;

            try
            {
                // Cria o campo de usuário na tabela OADM
                // SAPbobsCOM.Company oCompany = (SAPbobsCOM.Company)sboApplication.Company.GetDICompany();
                // SAPbobsCOM.UserFieldsMD userFields = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields) as SAPbobsCOM.UserFieldsMD;
                // userFields.TableName = "OADM";
                // userFields.Name = "CustomData";
                // userFields.Type = SAPbobsCOM.BoFieldTypes.db_Numeric;
                // userFields.Description = "Outros Dados";
                // userFields.Add();

                SAPbouiCOM.EditText txtStartPage = GetSBOEditText(targetForm, "txtPage");
                dataConnector.OpenConnection("sqlServer");
                dataConnector.SqlServerConnection.ChangeDatabase(sboApplication.Company.DatabaseName);
                String query = "UPDATE OADM SET U_CustomData = " + txtStartPage.Value;
                SqlCommand command = new SqlCommand(query, dataConnector.SqlServerConnection);
                command.ExecuteNonQuery();
                dataConnector.CloseConnection();

                targetForm.Close();
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        private void DisplayOurNum(String formUID)
        {
            SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            if (targetForm == null) return;

            try
            {
                SAPbouiCOM.Item referenceItem = targetForm.Items.Item("8");
                SAPbouiCOM.EditText referenceSpecific = (SAPbouiCOM.EditText)referenceItem.Specific;
                int boeNum = int.Parse(referenceSpecific.Value);

                dataConnector.OpenConnection("sqlServer");
                dataConnector.SqlServerConnection.ChangeDatabase(sboApplication.Company.DatabaseName);
                BillOfExchangeDAO boeDAO = new BillOfExchangeDAO(dataConnector.SqlServerConnection);
                BillOfExchangeDTO billOfExchange = boeDAO.GetBillOfExchange(boeNum);
                dataConnector.CloseConnection();

                SAPbouiCOM.Item labelItem = targetForm.Items.Add("lblOurNum", BoFormItemTypes.it_STATIC);
                labelItem.Top = 100;
                labelItem.Left = 170;
                SAPbouiCOM.StaticText labelSpecific = (SAPbouiCOM.StaticText)labelItem.Specific;
                labelSpecific.Caption = "Nosso Número";

                SAPbouiCOM.Item textItem = targetForm.Items.Add("txtOurNum", BoFormItemTypes.it_EDIT);
                textItem.Top = 116;
                textItem.Left = 170;
                SAPbouiCOM.EditText textSpecific = (SAPbouiCOM.EditText)textItem.Specific;
                textSpecific.Value = billOfExchange.OurNum + "-" + billOfExchange.OurNumChk;
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        private void CreateSpecificationsTab(String formUID)
        {
            // not implemented yet
        }

        private void ResizeSpecificationsTab(String formUID)
        {
            // not implemented yet
        }

        private void OpenSpecificationsTab(String formUID)
        {
            // not implemented yet
        }

        private void CreateInstallationTab(String formUID)
        {
            SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            if (targetForm == null) return;

            try
            {
                // Cria a aba "Instalação" no form do cartão de equipamento
                SAPbouiCOM.Item lastFolder = targetForm.Items.Item("39");
                SAPbouiCOM.Item newFolder = targetForm.Items.Add("instTab", BoFormItemTypes.it_FOLDER);
                newFolder.Width = lastFolder.Width;
                newFolder.Height = lastFolder.Height;
                newFolder.Top = lastFolder.Top;
                newFolder.Left = lastFolder.Left + lastFolder.Width;
                newFolder.Visible = true;
                SAPbouiCOM.Folder installationTab = ((SAPbouiCOM.Folder)(newFolder.Specific));
                installationTab.Caption = "Instalação";
                installationTab.GroupWith("39");

                // Adiciona os campos de usuário ao cartão de equipamento
                SAPbouiCOM.EditText installationDate = AddSBOTextField(targetForm, "InstDte", "Data Instalação", 25, newFolder.Top + 35, INSTALLATION_TAB);
                installationDate.DataBind.SetBound(true, "OINS", "U_InstallationDate");
                SAPbouiCOM.EditText installationDocNum = AddSBOTextField(targetForm, "InstDoc", "Número NF Remessa", 25, newFolder.Top + 60, INSTALLATION_TAB);
                installationDocNum.DataBind.SetBound(true, "OINS", "U_InstallationDocNum");
                SAPbouiCOM.EditText bwPageCounter = AddSBOTextField(targetForm, "BwCntr", "Contador Inicial (Pb)", 25, newFolder.Top + 85, INSTALLATION_TAB);
                bwPageCounter.DataBind.SetBound(true, "OINS", "U_BwPageCounter");
                SAPbouiCOM.ComboBox technician = AddSBOCombobox(targetForm, "Techncn", "Resp. Técnico", 25, newFolder.Top + 110, INSTALLATION_TAB);
                technician.DataBind.SetBound(true, "OINS", "U_Technician");
                SAPbouiCOM.EditText sla = AddSBOTextField(targetForm, "SLA", "Nível Serviço(SLA)", 25, newFolder.Top + 135, INSTALLATION_TAB);
                sla.DataBind.SetBound(true, "OINS", "U_SLA");
                SAPbouiCOM.EditText removalDate = AddSBOTextField(targetForm, "RemvDte", "Data Devolução", 300, newFolder.Top + 35, INSTALLATION_TAB);
                removalDate.DataBind.SetBound(true, "OINS", "U_RemovalDate");
                SAPbouiCOM.EditText removalDocNum = AddSBOTextField(targetForm, "RemvDoc", "Número NF Retorno", 300, newFolder.Top + 60, INSTALLATION_TAB);
                removalDocNum.DataBind.SetBound(true, "OINS", "U_RemovalDocNum");
                SAPbouiCOM.EditText bwPageCounter2 = AddSBOTextField(targetForm, "BwCntr2", "Contador Final (Pb)", 300, newFolder.Top + 85, INSTALLATION_TAB);
                bwPageCounter2.DataBind.SetBound(true, "OINS", "U_BwPageCounter2");
                SAPbouiCOM.EditText capacity = AddSBOTextField(targetForm, "Capacty", "Capacidade", 300, newFolder.Top + 110, INSTALLATION_TAB);
                capacity.DataBind.SetBound(true, "OINS", "U_Capacity");
                SAPbouiCOM.EditText comment = AddSBOTextField(targetForm, "Comment", "Observações", 300, newFolder.Top + 135, INSTALLATION_TAB);
                comment.DataBind.SetBound(true, "OINS", "U_Comments");

                // Preenche o combo de escolha de técnicos com os técnicos cadastrados
                dataConnector.OpenConnection("sqlServer");
                dataConnector.SqlServerConnection.ChangeDatabase(sboApplication.Company.DatabaseName);
                EmployeeDAO employeeDAO = new EmployeeDAO(dataConnector.SqlServerConnection);
                List<EmployeeDTO> technicianList = employeeDAO.GetAllTechnicians();
                dataConnector.CloseConnection();
                foreach (EmployeeDTO technicianDTO in technicianList)
                    technician.ValidValues.Add(technicianDTO.empID.ToString(), technicianDTO.firstName + " " + technicianDTO.lastName);
                // Recoloca o combo de técnicos na aba de instalação ( correção de bug )
                SAPbouiCOM.Item cmbTechnician = targetForm.Items.Item("cmbTechncn");
                cmbTechnician.FromPane = INSTALLATION_TAB;
                cmbTechnician.ToPane = INSTALLATION_TAB;

                CreateAccessoriesMatrix(targetForm, 25, newFolder.Top + 165, false, INSTALLATION_TAB, false, false);
                SAPbouiCOM.Button addButton = AddSBOButton(targetForm, "Add", "Incluir Acessório", 400, newFolder.Top + 85, INSTALLATION_TAB);
                SAPbouiCOM.Button removeButton = AddSBOButton(targetForm, "Remove", "Remover Acessório", 400, newFolder.Top + 85, INSTALLATION_TAB);
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        private SAPbouiCOM.Matrix CreateAccessoriesMatrix(SAPbouiCOM.Form targetForm, int left, int top, Boolean visible, int tab, Boolean displayFirstColumn, Boolean editable)
        {
            try
            {
                // Cria o grid contendo os acessórios do equipamento
                DBDataSource dataSource = targetForm.DataSources.DBDataSources.Add("@ACCESSORIES");
                SAPbouiCOM.Item titleItem = targetForm.Items.Add("lblTitle", BoFormItemTypes.it_STATIC);
                titleItem.Left = left;
                titleItem.Top = top;
                titleItem.FromPane = tab;
                titleItem.ToPane = tab;
                titleItem.Visible = visible;
                SAPbouiCOM.StaticText titleSpecific = (SAPbouiCOM.StaticText)titleItem.Specific;
                titleSpecific.Caption = "Accessórios";
                SAPbouiCOM.Item matrixItem = targetForm.Items.Add("accessries", BoFormItemTypes.it_MATRIX);
                matrixItem.FromPane = tab;
                matrixItem.ToPane = tab;
                matrixItem.Left = left;
                matrixItem.Top = titleItem.Top + 12;
                matrixItem.Width = targetForm.ClientWidth - 40;
                matrixItem.Height = 120;
                matrixItem.Visible = visible;
                SAPbouiCOM.Matrix matrixSpecific = (SAPbouiCOM.Matrix)matrixItem.Specific;
                matrixSpecific.SelectionMode = BoMatrixSelect.ms_Single;
                SAPbouiCOM.Column column0 = matrixSpecific.Columns.Add("colmn0", BoFormItemTypes.it_EDIT);
                column0.Width = 30;
                column0.Visible = displayFirstColumn;
                column0.Editable = editable;
                column0.DataBind.SetBound(true, "@ACCESSORIES", "Code");
                SAPbouiCOM.Column column1 = matrixSpecific.Columns.Add("colmn1", BoFormItemTypes.it_EDIT);
                column1.TitleObject.Caption = "Código do Item";
                column1.Width = 90;
                column1.Editable = editable;
                column1.DataBind.SetBound(true, "@ACCESSORIES", "U_ItemCode");
                SAPbouiCOM.Column column2 = matrixSpecific.Columns.Add("colmn2", BoFormItemTypes.it_EDIT);
                column2.TitleObject.Caption = "Nome do Item";
                column2.Width = 180;
                column2.Editable = editable;
                column2.DataBind.SetBound(true, "@ACCESSORIES", "U_ItemName");
                SAPbouiCOM.Column column3 = matrixSpecific.Columns.Add("colmn3", BoFormItemTypes.it_EDIT);
                column3.TitleObject.Caption = "Quantidade";
                column3.Width = 70;
                column3.Editable = editable;
                column3.DataBind.SetBound(true, "@ACCESSORIES", "U_Amount");

                return matrixSpecific;
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
                return null;
            }
        }

        private void ResizeInstallationTab(String formUID)
        {
            SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            if (targetForm == null) return;

            try
            {
                int tabTop = 0;
                SAPbouiCOM.Item installationTab = targetForm.Items.Item("instTab");
                if (installationTab != null) tabTop = installationTab.Top;

                SAPbouiCOM.Item lblInstallationDate = targetForm.Items.Item("lblInstDte");
                lblInstallationDate.Top = tabTop + 35;
                SAPbouiCOM.Item txtInstallationDate = targetForm.Items.Item("txtInstDte");
                txtInstallationDate.Top = tabTop + 35;
                SAPbouiCOM.Item lblInstallationDocNum = targetForm.Items.Item("lblInstDoc");
                lblInstallationDocNum.Top = lblInstallationDate.Top + 25;
                SAPbouiCOM.Item txtInstallationDocNum = targetForm.Items.Item("txtInstDoc");
                txtInstallationDocNum.Top = lblInstallationDate.Top + 25;
                SAPbouiCOM.Item lblBwPageCounter = targetForm.Items.Item("lblBwCntr");
                lblBwPageCounter.Top = lblInstallationDate.Top + 50;
                SAPbouiCOM.Item txtBwPageCounter = targetForm.Items.Item("txtBwCntr");
                txtBwPageCounter.Top = lblInstallationDate.Top + 50;
                SAPbouiCOM.Item lblTechnician = targetForm.Items.Item("lblTechncn");
                lblTechnician.Top = lblInstallationDate.Top + 75;
                SAPbouiCOM.Item cmbTechnician = targetForm.Items.Item("cmbTechncn");
                cmbTechnician.Top = lblInstallationDate.Top + 75;
                SAPbouiCOM.Item lblSla = targetForm.Items.Item("lblSLA");
                lblSla.Top = lblInstallationDate.Top + 100;
                SAPbouiCOM.Item txtSla = targetForm.Items.Item("txtSLA");
                txtSla.Top = lblInstallationDate.Top + 100;
                SAPbouiCOM.Item lblRemovalDate = targetForm.Items.Item("lblRemvDte");
                lblRemovalDate.Top = tabTop + 35;
                SAPbouiCOM.Item txtRemovalDate = targetForm.Items.Item("txtRemvDte");
                txtRemovalDate.Top = tabTop + 35;
                SAPbouiCOM.Item lblRemovalDocNum = targetForm.Items.Item("lblRemvDoc");
                lblRemovalDocNum.Left = lblRemovalDate.Left;
                lblRemovalDocNum.Top = lblRemovalDate.Top + 25;
                SAPbouiCOM.Item txtRemovalDocNum = targetForm.Items.Item("txtRemvDoc");
                txtRemovalDocNum.Left = lblRemovalDocNum.Left + lblRemovalDocNum.Width + 10;
                txtRemovalDocNum.Top = lblRemovalDate.Top + 25;
                SAPbouiCOM.Item lblBwPageCounter2 = targetForm.Items.Item("lblBwCntr2");
                lblBwPageCounter2.Left = lblRemovalDate.Left;
                lblBwPageCounter2.Top = lblRemovalDate.Top + 50;
                SAPbouiCOM.Item txtBwPageCounter2 = targetForm.Items.Item("txtBwCntr2");
                txtBwPageCounter2.Left = lblBwPageCounter2.Left + lblBwPageCounter2.Width + 10;
                txtBwPageCounter2.Top = lblRemovalDate.Top + 50;
                SAPbouiCOM.Item lblCapacity = targetForm.Items.Item("lblCapacty");
                lblCapacity.Left = lblRemovalDate.Left;
                lblCapacity.Top = lblRemovalDate.Top + 75;
                SAPbouiCOM.Item txtCapacity = targetForm.Items.Item("txtCapacty");
                txtCapacity.Left = lblCapacity.Left + lblCapacity.Width + 10;
                txtCapacity.Top = lblRemovalDate.Top + 75;
                SAPbouiCOM.Item lblComment = targetForm.Items.Item("lblComment");
                lblComment.Left = lblRemovalDate.Left;
                lblComment.Top = lblRemovalDate.Top + 100;
                SAPbouiCOM.Item txtComment = targetForm.Items.Item("txtComment");
                txtComment.Left = lblComment.Left + lblComment.Width + 10;
                txtComment.Top = lblRemovalDate.Top + 100;


                SAPbouiCOM.Item titleItem = targetForm.Items.Item("lblTitle");
                titleItem.Top = lblInstallationDate.Top + 125;
                SAPbouiCOM.Item matrixItem = targetForm.Items.Item("accessries");
                matrixItem.Width = targetForm.ClientWidth - 40;
                matrixItem.Height = 120;
                SAPbouiCOM.Matrix matrixSpecific = (SAPbouiCOM.Matrix)matrixItem.Specific;
                matrixSpecific.Columns.Item(0).Width = 30;
                matrixSpecific.Columns.Item(1).Width = 90;
                matrixSpecific.Columns.Item(2).Width = 180;
                matrixSpecific.Columns.Item(3).Width = 70;
                // Ajusta o grid de acordo com o resize do form
                int totalSize = 0;
                int lastIndex = 0;
                for (int colIndex = 0; colIndex < matrixSpecific.Columns.Count; colIndex++)
                {
                    totalSize += matrixSpecific.Columns.Item(colIndex).Width;
                    lastIndex = colIndex;
                }
                int fillWidth = matrixItem.Width - totalSize - 15;
                matrixSpecific.Columns.Item(lastIndex).Width += fillWidth;

                int middle = (matrixItem.Width - matrixItem.Left) / 2;
                SAPbouiCOM.Item btnAdd = targetForm.Items.Item("btnAdd");
                btnAdd.Left = middle - 120;
                btnAdd.Top = lblInstallationDate.Top + 275;
                SAPbouiCOM.Item btnRemove = targetForm.Items.Item("btnRemove");
                btnRemove.Left = middle + 80;
                btnRemove.Top = lblInstallationDate.Top + 275;
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        private void OpenInstallationTab(String formUID)
        {
            // Altera o pane level do form para INSTALLATION_TAB, fazendo com que o SAP exiba apenas controles desta aba
            SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            if (targetForm != null) targetForm.PaneLevel = INSTALLATION_TAB;
        }

        private void SelectAccessory(int row, String formUID)
        {
            SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            if (targetForm == null) return;
            try
            {
                Item matrixItem = targetForm.Items.Item("accessries");
                Matrix matrixSpecific = (Matrix)matrixItem.Specific;
                matrixSpecific.SelectRow(row, true, false);
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        private void AddAccessory(String formUID)
        {
            SAPbouiCOM.Form parentForm = GetSBOForm(formUID);
            if (parentForm == null) return;

            int offset = parentForm.DataSources.DBDataSources.Item(0).Offset;
            String equipmentCardId = parentForm.DataSources.DBDataSources.Item(0).GetValue("insID", offset);
            if (String.IsNullOrEmpty(equipmentCardId))
            {
                sboApplication.MessageBox("Nenhum cartão de equipamento selecionado!", 0, "OK", "", "");
                return;
            }

            try
            {
                // Cria o form de inclusão de acessório
                SAPbouiCOM.Form accessoryForm = GetSBOForm("frmAccssry");
                if (accessoryForm == null)
                {
                    FormCreationParams creationPackage = (FormCreationParams)sboApplication.CreateObject(BoCreatableObjectType.cot_FormCreationParams);
                    creationPackage.UniqueID = "frmAccssry";
                    creationPackage.FormType = "customForm";
                    creationPackage.BorderStyle = BoFormBorderStyle.fbs_Fixed;
                    accessoryForm = sboApplication.Forms.AddEx(creationPackage);
                }
                accessoryForm.Width = 350;
                accessoryForm.Height = 200;
                accessoryForm.AutoManaged = true;
                accessoryForm.Title = "Incluir Acessório";
                accessoryForm.Visible = true;

                // Armazena o form que originou a chamada
                UserDataSource parentDs = accessoryForm.DataSources.UserDataSources.Add("parentDs", BoDataType.dt_SHORT_TEXT, 30);
                parentDs.Value = parentForm.UniqueID;

                // Adiciona os campos de dados referentes ao acessório
                UserDataSource equipmentCardDs = accessoryForm.DataSources.UserDataSources.Add("equipCrdDs", BoDataType.dt_SHORT_NUMBER, 10);
                equipmentCardDs.Value = equipmentCardId;
                ChooseFromListCreationParams creationParams = (ChooseFromListCreationParams)sboApplication.CreateObject(BoCreatableObjectType.cot_ChooseFromListCreationParams);
                creationParams.UniqueID = "lstItems";
                creationParams.ObjectType = ((int)SAPbobsCOM.BoObjectTypes.oItems).ToString();
                creationParams.MultiSelection = false;
                ChooseFromList list = accessoryForm.ChooseFromLists.Add(creationParams);
                SAPbouiCOM.EditText txtItemCode = AddSBOTextField(accessoryForm, "ItmCode", "Código do Item", 25, 25, 0);
                accessoryForm.DataSources.UserDataSources.Add("itmCodeDs", BoDataType.dt_SHORT_TEXT, 30);
                txtItemCode.DataBind.SetBound(true, "", "itmCodeDs");
                txtItemCode.ChooseFromListUID = list.UniqueID;
                txtItemCode.ChooseFromListAlias = "ItemCode";
                SAPbouiCOM.EditText txtItemName = AddSBOTextField(accessoryForm, "ItmName", "Nome do Item", 25, 50, 0);
                accessoryForm.DataSources.UserDataSources.Add("itmNameDs", BoDataType.dt_SHORT_TEXT, 200);
                txtItemName.DataBind.SetBound(true, "", "itmNameDs");
                SAPbouiCOM.EditText txtAmount = AddSBOTextField(accessoryForm, "Amount", "Quantidade", 25, 75, 0);
                accessoryForm.DataSources.UserDataSources.Add("amountDs", BoDataType.dt_SHORT_NUMBER, 10);
                txtAmount.DataBind.SetBound(true, "", "amountDs");
                SAPbouiCOM.Button btnAdd = AddSBOButton(accessoryForm, "Add", "Incluir", 140, 120, 0);
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        private void RemoveAccessory(String formUID)
        {
            SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            if (targetForm == null) return;

            int offset = targetForm.DataSources.DBDataSources.Item(0).Offset;
            String equipmentCardId = targetForm.DataSources.DBDataSources.Item(0).GetValue("insID", offset);
            if (String.IsNullOrEmpty(equipmentCardId))
            {
                sboApplication.MessageBox("Nenhum cartão de equipamento selecionado!", 0, "OK", "", "");
                return;
            }

            SAPbouiCOM.Item matrixItem = targetForm.Items.Item("accessries");
            SAPbouiCOM.Matrix matrixSpecific = (SAPbouiCOM.Matrix)matrixItem.Specific;
            int? selectedRow = null;
            for (int rowIndex = 1; rowIndex <= matrixSpecific.Columns.Item(0).Cells.Count; rowIndex++)
            {
                if (matrixSpecific.IsRowSelected(rowIndex))
                    selectedRow = rowIndex;
            }
            if (selectedRow == null)
            {
                sboApplication.MessageBox("Nenhum acessório selecionado!", 0, "OK", "", "");
                return;
            }

            SAPbouiCOM.EditText code = (SAPbouiCOM.EditText)matrixSpecific.Columns.Item(0).Cells.Item(selectedRow).Specific;
            SAPbobsCOM.UserTable accessoryTable = sboCompany.UserTables.Item("ACCESSORIES");
            if (accessoryTable.GetByKey(code.Value)) accessoryTable.Remove();
            // Limpa a seleção e recarrega os acessórios do equipamento
            matrixSpecific.ClearSelections();
            ReloadAccessories(targetForm);
        }

        private void ChooseItem(ItemEvent pVal)
        {
            SAPbouiCOM.Form targetForm = GetSBOForm(pVal.FormUID);
            if (targetForm == null) return;

            try
            {
                IChooseFromListEvent chooseEvent = (IChooseFromListEvent)pVal;
                DataTable itemsResult = chooseEvent.SelectedObjects;

                SAPbouiCOM.Item itemName = targetForm.Items.Item("txtItmName");
                itemName.Click(BoCellClickType.ct_Regular);
                SAPbouiCOM.EditText itemNameSpecific = (SAPbouiCOM.EditText)itemName.Specific;
                itemNameSpecific.Value = (String)itemsResult.GetValue(1, 0);

                SAPbouiCOM.Item itemCode = targetForm.Items.Item("txtItmCode");
                itemCode.Click(BoCellClickType.ct_Regular);
                SAPbouiCOM.EditText itemCodeSpecific = (SAPbouiCOM.EditText)itemCode.Specific;
                itemCodeSpecific.Value = (String)itemsResult.GetValue(0, 0);
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        private void SaveNewAccessory(String formUID)
        {
            SAPbouiCOM.Form accessoryForm = GetSBOForm(formUID);
            if (accessoryForm == null) return;

            // Recupera o parent do form de acessórios
            UserDataSource parentDs = accessoryForm.DataSources.UserDataSources.Item("parentDs");
            SAPbouiCOM.Form parentForm = GetSBOForm(parentDs.Value);
            if (parentForm == null) return;

            // Insere o registro na tabela
            try
            {
                UserDataSource equipmentCardDs = GetSBOUserDataSource(accessoryForm, "equipCrdDs");
                SAPbouiCOM.EditText txtItemCode = GetSBOEditText(accessoryForm, "txtItmCode");
                SAPbouiCOM.EditText txtItemName = GetSBOEditText(accessoryForm, "txtItmName");
                SAPbouiCOM.EditText txtAmount = GetSBOEditText(accessoryForm, "txtAmount");

                if (String.IsNullOrEmpty(txtItemCode.Value) || String.IsNullOrEmpty(txtItemName.Value) || String.IsNullOrEmpty(txtAmount.Value))
                {
                    sboApplication.MessageBox("É necessário preencher todos os campos!", 0, "OK", "", "");
                    return;
                }

                int rowCount = 0;
                int recordId = 0;
                String query = "";
                SqlCommand command = null;
                SqlDataReader dataReader = null;
                dataConnector.OpenConnection("sqlServer");
                dataConnector.SqlServerConnection.ChangeDatabase(sboApplication.Company.DatabaseName);
                query = "SELECT COUNT(*) FROM [@ACCESSORIES]";
                command = new SqlCommand(query, dataConnector.SqlServerConnection);
                dataReader = command.ExecuteReader();
                if (dataReader.Read()) rowCount = (int)dataReader[0];
                dataReader.Close();
                if (rowCount > 0)
                {
                    query = "SELECT MAX(CAST(Code AS INT)) FROM [@ACCESSORIES]";
                    command = new SqlCommand(query, dataConnector.SqlServerConnection);
                    dataReader = command.ExecuteReader();
                    if (dataReader.Read()) recordId = (int)dataReader[0];
                    dataReader.Close();
                }
                dataConnector.CloseConnection();

                SAPbobsCOM.UserTable accessoryTable = sboCompany.UserTables.Item("ACCESSORIES");
                accessoryTable.Code = (recordId + 1).ToString();
                accessoryTable.Name = (recordId + 1).ToString();
                accessoryTable.UserFields.Fields.Item("U_InsId").Value = equipmentCardDs.Value;
                accessoryTable.UserFields.Fields.Item("U_ItemCode").Value = txtItemCode.Value;
                accessoryTable.UserFields.Fields.Item("U_ItemName").Value = txtItemName.Value;
                accessoryTable.UserFields.Fields.Item("U_Amount").Value = txtAmount.Value;
                int addResult = accessoryTable.Add();
                if (addResult != 0)
                {
                    int errorCode;
                    String errorMessage;
                    sboCompany.GetLastError(out errorCode, out errorMessage);
                    sboApplication.MessageBox("Falha ao inserir registro! " + errorMessage, 0, "Ok", "", "");
                }
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }

            // Atualiza o grid no parent form
            ReloadAccessories(parentForm);
            // Fecha o form de inclusão de acessório após ter adicionado o registro na tabela
            accessoryForm.Close();
        }

        private void ReloadAccessories(SAPbouiCOM.Form targetForm)
        {
            if (targetForm == null) return;

            int offset = targetForm.DataSources.DBDataSources.Item(0).Offset;
            String equipmentCardId = targetForm.DataSources.DBDataSources.Item(0).GetValue("insID", offset);

            DBDataSource dataSource = targetForm.DataSources.DBDataSources.Item("@ACCESSORIES");
            Conditions conditionSet = new Conditions();
            Condition condition = conditionSet.Add();
            condition.Alias = "U_InsID";
            condition.Operation = BoConditionOperation.co_EQUAL;
            condition.CondVal = equipmentCardId;
            dataSource.Query(conditionSet);

            SAPbouiCOM.Item matrixItem = targetForm.Items.Item("accessries");
            SAPbouiCOM.Matrix matrixSpecific = (SAPbouiCOM.Matrix)matrixItem.Specific;
            matrixSpecific.LoadFromDataSource();
        }

        private void SetDefaultCarrier(String formUID)
        {
            SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            if (targetForm == null) return;

            // Realiza a operação em uma thread separada
            BackgroundWorker backgroundWorker = new BackgroundWorker();
            backgroundWorker.DoWork += new DoWorkEventHandler(UpdateCarrierField);
            backgroundWorker.RunWorkerAsync(targetForm);

            System.Windows.Forms.Application.DoEvents();
        }

        private void UpdateCarrierField(Object sender, DoWorkEventArgs e)
        {
            SAPbouiCOM.Form targetForm = (SAPbouiCOM.Form)e.Argument;
            Thread.Sleep(600); // aguarda alguns milisegundos até o carregamento dos campos
            try
            {
                dataConnector.OpenConnection("sqlServer");
                dataConnector.SqlServerConnection.ChangeDatabase(sboApplication.Company.DatabaseName);
                BusinessPartnerDAO bpDAO = new BusinessPartnerDAO(dataConnector.SqlServerConnection);
                BusinessPartnerDTO defaultCarrier = bpDAO.GetDefaultCarrier();
                dataConnector.CloseConnection();

                SAPbouiCOM.Item textItem = targetForm.Items.Item("2022");
                textItem.Click(BoCellClickType.ct_Regular);
                SAPbouiCOM.EditText textSpecific = (SAPbouiCOM.EditText)textItem.Specific;
                if (String.IsNullOrEmpty(textSpecific.Value))
                    textSpecific.Value = defaultCarrier.CardCode;
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        // Recebe como parâmetro o DocCode da tabela RDOC  -- SELECT DocCode, DocName, PaperSize, Template FROM RDOC WHERE Category = 'C'
        private void ExportCrystalReport(String docCode, String filePath)
        {
            // Specify the table and blob field
            SAPbobsCOM.BlobParams blobParams = (SAPbobsCOM.BlobParams)sboCompany.GetCompanyService().GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlobParams);
            blobParams.Table = "RDOC";
            blobParams.Field = "Template";

            // Specify the file name to which to write the blob
            blobParams.FileName = filePath;

            // Specify the key field and value of the row from which to get the blob
            SAPbobsCOM.BlobTableKeySegment keySegment;
            keySegment = blobParams.BlobTableKeySegments.Add();
            keySegment.Name = "DocCode";
            keySegment.Value = docCode;

            // Save the blob to the file
            sboCompany.GetCompanyService().SaveBlobToFile(blobParams);
        }

        private SAPbouiCOM.Form GetSBOForm(String formName)
        {
            SAPbouiCOM.Form oForm = null;
            try
            {
                oForm = sboApplication.Forms.Item(formName);
            }
            catch
            {
                return null;
            }
            return oForm;
        }

        private SAPbouiCOM.Form CreateSBOForm(String formName, String formTitle, int width, int height)
        {
            SAPbouiCOM.Form oForm = GetSBOForm(formName);
            if (oForm != null) return oForm; // O form já existe, retorna o form criado anteriormente

            SAPbouiCOM.FormCreationParams fcp = ((SAPbouiCOM.FormCreationParams)(sboApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));
            fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;
            fcp.FormType = "custom";
            fcp.UniqueID = formName;

            oForm = sboApplication.Forms.AddEx(fcp);
            oForm.Title = formTitle;
            oForm.Width = width;
            oForm.Height = height;
            oForm.AutoManaged = true;
            oForm.Visible = true;

            return oForm;
        }

        private SAPbouiCOM.UserDataSource GetSBOUserDataSource(SAPbouiCOM.Form sboForm, String dsName)
        {
            if (sboForm == null) return null;

            SAPbouiCOM.UserDataSource sboUserDataSource = null;
            try
            {
                sboUserDataSource = sboForm.DataSources.UserDataSources.Item(dsName);
            }
            catch
            {
                return null;
            }
            return sboUserDataSource;
        }

        private SAPbouiCOM.Item AddSBOLabel(SAPbouiCOM.Form sboForm, String name, String caption, int left, int top, int tab)
        {
            SAPbouiCOM.Item labelItem = sboForm.Items.Add("lbl" + name, BoFormItemTypes.it_STATIC);
            labelItem.Left = left;
            labelItem.Top = top;
            labelItem.Width = 106;
            labelItem.FromPane = tab;
            labelItem.ToPane = tab;
            SAPbouiCOM.StaticText labelSpecific = (SAPbouiCOM.StaticText)labelItem.Specific;
            labelSpecific.Caption = caption;

            return labelItem;
        }

        private SAPbouiCOM.EditText AddSBOTextField(SAPbouiCOM.Form sboForm, String name, String caption, int left, int top, int tab, Boolean doLineBreaks)
        {
            if (sboForm == null) return null;

            try
            {
                SAPbouiCOM.Item labelItem = AddSBOLabel(sboForm, name, caption, left, top, tab);

                BoFormItemTypes itemType = BoFormItemTypes.it_EDIT;
                if (doLineBreaks) itemType = BoFormItemTypes.it_EXTEDIT;
                SAPbouiCOM.Item editItem = sboForm.Items.Add("txt" + name, itemType);
                editItem.Left = left + labelItem.Width + 10;
                editItem.Top = top;
                editItem.FromPane = tab;
                editItem.ToPane = tab;
                editItem.AffectsFormMode = true;
                editItem.SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_All, BoModeVisualBehavior.mvb_Default);
                SAPbouiCOM.EditText editSpecific = (SAPbouiCOM.EditText)editItem.Specific;
                editSpecific.Value = "";

                return editSpecific;
            }
            catch
            {
                return null;
            }
        }

        private SAPbouiCOM.EditText AddSBOTextField(SAPbouiCOM.Form sboForm, String name, String caption, int left, int top, int tab)
        {
            return AddSBOTextField(sboForm, name, caption, left, top, tab, false);
        }

        private SAPbouiCOM.ComboBox AddSBOCombobox(SAPbouiCOM.Form sboForm, String name, String caption, int left, int top, int tab)
        {
            if (sboForm == null) return null;

            try
            {
                SAPbouiCOM.Item labelItem = AddSBOLabel(sboForm, name, caption, left, top, tab);

                SAPbouiCOM.Item comboItem = sboForm.Items.Add("cmb" + name, BoFormItemTypes.it_COMBO_BOX);
                comboItem.Left = left + labelItem.Width + 10;
                comboItem.Top = top;
                comboItem.FromPane = tab;
                comboItem.ToPane = tab;
                comboItem.Visible = true;
                comboItem.AffectsFormMode = true;
                comboItem.SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_All, BoModeVisualBehavior.mvb_Default);
                SAPbouiCOM.ComboBox comboSpecific = (SAPbouiCOM.ComboBox)comboItem.Specific;

                return comboSpecific;
            }
            catch
            {
                return null;
            }
        }

        private SAPbouiCOM.Button AddSBOButton(SAPbouiCOM.Form sboForm, String name, String caption, int left, int top, int tab)
        {
            if (sboForm == null) return null;

            try
            {
                int buttonWidth = caption.Length * 6;
                if (buttonWidth < 80) buttonWidth = 80;

                SAPbouiCOM.Item buttonItem = sboForm.Items.Add("btn" + name, BoFormItemTypes.it_BUTTON);
                buttonItem.FromPane = tab;
                buttonItem.ToPane = tab;
                buttonItem.Left = left;
                buttonItem.Top = top;
                buttonItem.Width = buttonWidth;
                buttonItem.Height = 20;
                SAPbouiCOM.Button buttonSpecific = (SAPbouiCOM.Button)buttonItem.Specific;
                buttonSpecific.Caption = caption;

                return buttonSpecific;
            }
            catch
            {
                return null;
            }
        }

        private SAPbouiCOM.EditText GetSBOEditText(SAPbouiCOM.Form sboForm, String itemUID)
        {
            if (sboForm == null) return null;

            SAPbouiCOM.Item editTextItem = null;
            SAPbouiCOM.EditText editTextSpecific = null;
            try
            {
                editTextItem = sboForm.Items.Item(itemUID);
                editTextSpecific = (SAPbouiCOM.EditText)editTextItem.Specific;
            }
            catch
            {
                return null;
            }

            return editTextSpecific;
        }

        private SAPbouiCOM.ComboBox GetSBOComboBox(SAPbouiCOM.Form sboForm, String itemUID)
        {
            if (sboForm == null) return null;

            SAPbouiCOM.Item comboBoxItem = null;
            SAPbouiCOM.ComboBox comboBoxSpecific = null;
            try
            {
                comboBoxItem = sboForm.Items.Item(itemUID);
                comboBoxSpecific = (SAPbouiCOM.ComboBox)comboBoxItem.Specific;
            }
            catch
            {
                return null;
            }

            return comboBoxSpecific;
        }
    }

}

