using Microsoft.Office.InfoPath;
using System;
using System.Xml;
using System.Xml.XPath;
using System.Text.RegularExpressions;


namespace Droneskjema
{
    public partial class FormCode
    {
        // Member variables are not supported in browser-enabled forms.
        // Instead, write and read these values from the FormState
        // dictionary using code such as the following:
        //
        // private object _memberVariable
        // {
        //     get
        //     {
        //         return FormState["_memberVariable"];
        //     }
        //     set
        //     {
        //         FormState["_memberVariable"] = value;
        //     }
        // }

        // NOTE: The following procedure is required by Microsoft InfoPath.
        // It can be modified using Microsoft InfoPath.

        // Developer notes:
        // https://blogs.msdn.microsoft.com/infopath/2006/11/28/the-xsinil-attribute/

        
        private object _feeForOppstart
        {
            get
            {
                return FormState["_feeForOppstart"];
            }
            set
            {
                FormState["_feeForOppstart"] = value;
            }
        }


        private object _feeForAvslutning
        {
            get
            {
                return FormState["_feeForAvslutning"];
            }
            set
            {
                FormState["_feeForAvslutning"] = value;
            }
        }

        
        
        
        public void InternalStartup()
        {
            EventManager.XmlEvents["/melding/Organisasjon/organisasjonsnummer"].Changed += new XmlChangedEventHandler(organisasjonsnummer_Changed);
            EventManager.XmlEvents["/melding/Skjemadata/erOppstart"].Changed += new XmlChangedEventHandler(erOppstart_Changed);
            EventManager.XmlEvents["/uictrl/land_pulldown_data", "GuiElementControl"].Changed += new XmlChangedEventHandler(GuiElementControl__land_pulldown_data_Changed);
            EventManager.XmlEvents["/melding/Organisasjon/land"].Changed += new XmlChangedEventHandler(Organisasjon_land_Changed);
            EventManager.XmlEvents["/melding/Skjemadata/erForetak"].Changed += new XmlChangedEventHandler(erForetak_Changed);
        }


        public void erForetak_Changed(object sender, XmlEventArgs e)
        {
            string nullMelding = "";
            
            if (string.Equals(e.Site.InnerXml.ToString(), "false"))
            {
                SetNodeToString("/melding/Organisasjon/organisasjonsnummer", "privatperson", nullMelding);
                SetNodeToString("/melding/Organisasjon/adresse", "privatperson", nullMelding);
                SetNodeToString("/melding/Organisasjon/e-post", "privatperson", nullMelding);
                SetNodeToString("/melding/Organisasjon/land", "privatperson", nullMelding);
                SetNodeToString("/melding/Organisasjon/navn", "privatperson", nullMelding);
                SetNodeToString("melding/Organisasjon/postnummer", "privatperson", nullMelding);
                SetNodeToString("/melding/Organisasjon/poststed", "privatperson", nullMelding);
                SetNodeToString("/melding/Organisasjon/telefon", "privatperson", nullMelding);
            }
            else if (string.Equals(e.Site.InnerXml.ToString(), "true"))
            {
                SetNodeToString("/melding/Organisasjon/organisasjonsnummer", null, nullMelding);
                SetNodeToString("/melding/Organisasjon/adresse", null, nullMelding);
                SetNodeToString("/melding/Organisasjon/e-post", null, nullMelding);
                SetNodeToString("/melding/Organisasjon/land", null, nullMelding);
                SetNodeToString("/melding/Organisasjon/navn", null, nullMelding);
                SetNodeToString("melding/Organisasjon/postnummer", null, nullMelding);
                SetNodeToString("/melding/Organisasjon/poststed", null, nullMelding);
                SetNodeToString("/melding/Organisasjon/telefon", null, nullMelding);
            }
            else
            {
                // do nothing ... initial state
            }
        }

        
        
        
        
        public void organisasjonsnummer_Changed(object sender, XmlEventArgs e)
        {
            
            // Only run test if ifForetak is true;
            if (!GetMainDataSource_isForetak())
            {
                return;
            }
            
            
            string errorKey = "ORGNUM";

            ValidationResult result = Validate_organisasjonsnummer(e.Site.InnerXml);
            if (result.IsValid == true)
            {
                ReportError(e, errorKey, "Organisasjonsnummer", "Henter data fra Enhetsregisteret");
                result = GetOrgNumberData(e.Site.InnerXml);
                if (result.IsValid == true)
                {
                    DeleteErrorKey(errorKey);
                }
                else
                {
                    ReportError(e, errorKey, "Organisasjonsnummer", result.ErrorMsg);
                }
            }
            else
            {
                ReportError(e, errorKey, "Organisasjonsnummer", result.ErrorMsg);
            }
        }

        
        
        public void erOppstart_Changed(object sender, XmlEventArgs e)
        {
            string errorKey = "OPPSTARTCHANGED";
            DeleteErrorKey(errorKey);

            bool isGotFeesFromCodeList = GetRpasFeeFromCodeList();

            if (!isGotFeesFromCodeList)
            {
                ReportError(e, errorKey, "Avgift", "Feil ved utlesing av avgift, prøv igjen senere");
            }

            if (!String.IsNullOrEmpty(e.Site.InnerXml) && Convert.ToBoolean(e.Site.InnerXml) == true)
            {
                SetNodeToString("/melding/Betaling/sum", _feeForOppstart.ToString(), String.Empty);
            }
            else if (!String.IsNullOrEmpty(e.Site.InnerXml) && Convert.ToBoolean(e.Site.InnerXml) == false)
            {
                // Avslutning, ingen betaling, sum settes til _feeForAvslutning
                SetNodeToString("/melding/Betaling/sum", _feeForAvslutning.ToString(), String.Empty);
            }
        }



        public void GuiElementControl__land_pulldown_data_Changed(object sender, XmlEventArgs e)
        {
            // This should hit in the case a value is set on the pulldown list for the land.
            // If the system is operating in the SetGuiCtrlNode("/uictrl/land", "USEPULLDOWN"); mode, 
            // the current data in /uictrl/land_pulldown_data should be copied over to 
            // /melding/Organisasjon/land

            if (isSystemIn_USEPULLDOWN())
            {
                SetNodeToString("/melding/Organisasjon/land", get_land_pulldown_data(), "");   
            }
        }

        public string get_land_pulldown_data()
        {
            try {
                DataSource guiCtrl = DataSources["GuiElementControl"];
                XPathNavigator guiNav = guiCtrl.CreateNavigator().SelectSingleNode("/uictrl/land_pulldown_data", NamespaceManager);
                return guiNav.InnerXml.ToString();
            }
            catch
            {
                return "";
            }
        }
        
        
        public bool isSystemIn_USEPULLDOWN()
        {
            try
            {
                DataSource guiCtrl = DataSources["GuiElementControl"];
                XPathNavigator guiNav = guiCtrl.CreateNavigator().SelectSingleNode("/uictrl/land", NamespaceManager);

                if (string.Equals(guiNav.InnerXml.ToString(), "USEPULLDOWN")) return true;
                else return false;
            }
            catch
            {
                return false;
            }
       
        }


        public void Organisasjon_land_Changed(object sender, XmlEventArgs e)
        {
            string errorKey = "FORETAKADRLAND";

            if (isSystemIn_USEPULLDOWN() && (string.Equals(e.Site.InnerXml.ToString(), "") || string.Equals(e.Site.InnerXml.ToString(), "ikkeValgt")))
            {
                ReportError(e, errorKey, "Land", "Velg land for foretakets adresse fra \"Land\" listen");
            }
            else
            {
                DeleteErrorKey(errorKey);
            }
        }

        
        public int GetTheFormLanguageCode()
        {
            int cc = 0;

            if (FormState.Contains("Language"))
            {
                if (FormState["Language"].Equals(1044)) cc = 1044;
                if (FormState["Language"].Equals(2068)) cc = 2068;
                if (FormState["Language"].Equals(1033)) cc = 1033;
            }
            else
            {
                cc = 1044; // defaults to bokmaal
            }
            return cc;
        }

        
        public void DeleteNil(XPathNavigator node)
        {
            if (node.MoveToAttribute("nil", "http://www.w3.org/2001/XMLSchema-instance"))
                node.DeleteSelf();
        }



        public ValidationResult GetOrgNumberData(string orgNumber)
        {
            no.altinn.RegisterER.RegisterERInfoPathSF client = new no.altinn.RegisterER.RegisterERInfoPathSF();

            try
            {
                no.altinn.RegisterER.OrganizationRegesterBEV2 data = client.GetOrganizationRegisterDataV2(1234, true, orgNumber);
                PopulateDataForOrgNumber(data);
                return new ValidationResult(true, "Data ble hentet fra Enhetsregisteret");
            }
            catch 
            {
                return new ValidationResult(false, "Feil ved uthenting av data fra Enhetsregisteret");
            }
        }


        public void SetGuiCtrlNode(string node, string value)
        {
            DataSource guiCtrl = DataSources["GuiElementControl"];
            XPathNavigator guiNav = guiCtrl.CreateNavigator().SelectSingleNode(node, NamespaceManager);
            guiNav.SetValue(value);
        }

        
        public void SetNodeToString(string xpath, string value, string nillValue)
        {
            XPathNavigator mainData = MainDataSource.CreateNavigator();
            XPathNavigator node = mainData.SelectSingleNode(xpath, NamespaceManager);
            
            if (!String.IsNullOrEmpty(value))
            {
                DeleteNil(node);
                node.SetValue(value);
            }
            else if (!String.IsNullOrEmpty(nillValue))
            {
                DeleteNil(node);
                node.SetValue(nillValue);
            } else {
                DeleteNil(node);
                node.SetValue("");
            }
        }


        public bool GetMainDataSource_isForetak()
        {
            try
            {
                XPathNavigator mainData = MainDataSource.CreateNavigator();
                XPathNavigator node = mainData.SelectSingleNode("/melding/Skjemadata/erForetak", NamespaceManager);
                
                if (string.Equals(node.InnerXml.ToString(), "true")) return true;
                else return false;
            }
            catch
            {
                return false;
            }
        }
        
        
        
        public void SetGuiCtrlData_CheckEmptyValue_SetCanEdit(string fieldValue, string guiNodeXpath)
        {
            if (String.IsNullOrEmpty(fieldValue))
            {
                SetGuiCtrlNode(guiNodeXpath, "CanEdit");               
            }
        }

        
        public void PopulateDataForOrgNumber(no.altinn.RegisterER.OrganizationRegesterBEV2 data)
        {
            string nullMelding = "";

            SetNodeToString("/melding/Organisasjon/adresse", data.BusinessAddress, nullMelding);
            SetGuiCtrlData_CheckEmptyValue_SetCanEdit(data.BusinessAddress, "/uictrl/adresse");

            SetNodeToString("/melding/Organisasjon/e-post", data.EMailAddress, nullMelding);
            SetGuiCtrlData_CheckEmptyValue_SetCanEdit(data.EMailAddress, "/uictrl/e-post");

            LT_LandCountryInfo countryName = FindAltinnCodeListNameForCountryCode(data.CountryCode);
            if (countryName != null)
            {
                SetNodeToString("/melding/Organisasjon/land", countryName.BackendCode, nullMelding);
                SetGuiCtrlNode("/uictrl/land_display_field", countryName.Name);
                SetGuiCtrlNode("/uictrl/land", "USETEXTBOX");
            }
            else
            {
                SetGuiCtrlNode("/uictrl/land", "USEPULLDOWN");
                SetNodeToString("/melding/Organisasjon/land", "ikkeValgt", nullMelding);
            }
    
            SetNodeToString("/melding/Organisasjon/navn", data.Name, nullMelding);
            SetGuiCtrlData_CheckEmptyValue_SetCanEdit(data.Name, "/uictrl/navn");

            SetNodeToString("melding/Organisasjon/postnummer", data.BusinessPostCode, nullMelding);
            SetGuiCtrlData_CheckEmptyValue_SetCanEdit(data.BusinessPostCode, "/uictrl/postnummer");

            SetNodeToString("/melding/Organisasjon/poststed", data.BusinessPostCity, nullMelding);
            SetGuiCtrlData_CheckEmptyValue_SetCanEdit(data.BusinessPostCity, "/uictrl/poststed");

            SetNodeToString("/melding/Organisasjon/telefon", data.TelephoneNumber, nullMelding);
            SetGuiCtrlData_CheckEmptyValue_SetCanEdit(data.TelephoneNumber, "/uictrl/telefon");


        }


        public no.altinn.infopathCodeList.CodeList GetAltinnCodeList(string codelistName, int language)
        {
            // returns a code list from the internal Altinn codelist resources 
            try
            {
                no.altinn.infopathCodeList.ServiceMetaDataInfoPathSF request = new no.altinn.infopathCodeList.ServiceMetaDataInfoPathSF();
                return request.GetCodeList(codelistName, 0, true, language, true);
            }
            catch 
            {
                return null;
            }
        }




        public LT_LandCountryInfo FindAltinnCodeListNameForCountryCode(string countryCode)
        {
            // Altinn prod sends country code as two letters XX per https://no.wikipedia.org/wiki/ISO_3166-1_alfa-2
            // Match letters with "Value3" field of LT_Land codelist (XX)
            // Return the country name and the Value 2 and LT code in the Code field

            LT_LandCountryInfo returnData = null;


            // Did not get any data
            if (String.IsNullOrEmpty(countryCode)) return returnData; // null

            // Did no get a valid two digit CC
            Match CCValid = Regex.Match(countryCode.ToUpper(), "^[A-Z]{2}$");
            if (!CCValid.Success) return returnData; // null

            try
            {
                no.altinn.infopathCodeList.CodeList codeListLT_Land = GetAltinnCodeList("LT_Land", GetTheFormLanguageCode());

                for (int lt = 0; lt < codeListLT_Land.CodeListRows.Length; lt++)
                {
                    if (codeListLT_Land.CodeListRows[lt].Value3.ToUpper().Contains(countryCode.ToUpper()))
                    {
                        returnData = new LT_LandCountryInfo(codeListLT_Land.CodeListRows[lt].Value2, codeListLT_Land.CodeListRows[lt].Code); 
                    }
                }
            }
            catch { }

            return returnData;
        }

        
        
        public bool GetRpasFeeFromCodeList()
        {
            // Kodeliste: LT_RPAS_AVGIFT
            // Verdi 1: Avgift for oppstart
            // Verdi 2: Avgift for avslutting (not implemented)
            // Usage -- the code will pick the last entry in the list

            bool status = false;

            try
            {
                no.altinn.infopathCodeList.CodeList feeCodeList = GetAltinnCodeList("LT_RPAS_AVGIFT", 1044);

                if (feeCodeList.CodeListRows.Length > 0)
                {
                    _feeForOppstart = feeCodeList.CodeListRows[feeCodeList.CodeListRows.Length - 1].Value1;
                    _feeForAvslutning = feeCodeList.CodeListRows[feeCodeList.CodeListRows.Length - 1].Value2;
                    status = true;
                }
            }
            catch { }

            return status;
        }




        public void DeleteErrorKey(string errorKey)
        {
            try
            {
                this.Errors.Delete(errorKey);
            }
            catch (Exception) { }
        }

        
        public void ReportError(XmlEventArgs e, string errorKey, string keyword, string message)
        {
            DeleteErrorKey(errorKey);
            this.Errors.Add(e.Site, errorKey, keyword, message);
        }



        public ValidationResult Validate_organisasjonsnummer(string orgnum)
        {
            // Org number valid format: 
            //    https://www.brreg.no/om-oss/samfunnsoppdraget-vart/registera-vare/einingsregisteret/organisasjonsnummeret/
            ValidationResult result;

            Match isOrgNumValid = Regex.Match(orgnum, "^([0-9])([0-9])([0-9])([0-9])([0-9])([0-9])([0-9])([0-9])([0-9])$");

            if (isOrgNumValid.Success)
            {
                int products = Convert.ToInt32(isOrgNumValid.Groups[1].Value) * 3 +
                               Convert.ToInt32(isOrgNumValid.Groups[2].Value) * 2 +
                               Convert.ToInt32(isOrgNumValid.Groups[3].Value) * 7 +
                               Convert.ToInt32(isOrgNumValid.Groups[4].Value) * 6 +
                               Convert.ToInt32(isOrgNumValid.Groups[5].Value) * 5 +
                               Convert.ToInt32(isOrgNumValid.Groups[6].Value) * 4 +
                               Convert.ToInt32(isOrgNumValid.Groups[7].Value) * 3 +
                               Convert.ToInt32(isOrgNumValid.Groups[8].Value) * 2;

                int controlDigit = 11 - (products % 11);
                if (controlDigit == 11)
                {
                    controlDigit = 0;
                }

                if (controlDigit == Convert.ToInt32(isOrgNumValid.Groups[9].Value))
                {
                    result = new ValidationResult(true, "Organisasjonsnummeret er gyldig");
                }
                else
                {
                    result = new ValidationResult(false, "Organisasjonsnummeret har feil kontrollverdi");
                }
            }
            else
            {
                result = new ValidationResult(false, "Organisasjonsnummeret er ikke gyldig");
            }
            return result;
        }

    }




    public class ValidationResult
    {
        private bool _isValid;
        private string _errorMessage;

        public bool IsValid
        {
            get { return this._isValid; }
            set { this._isValid = value; }
        }

        public string ErrorMsg
        {
            get { return _errorMessage; }
            set { _errorMessage = value; }
        }

        public ValidationResult()
        {
        }

        public ValidationResult(bool isValid, string errorMsg)
        {
            IsValid = isValid;
            ErrorMsg = errorMsg;
        }
    }

    
    
    public class LT_LandCountryInfo
    {
        private string _name;
        private string _backendCode;

        public string Name
        {
            get { return this._name; }
            set { this._name = value; }
        }

        public string BackendCode
        {
            get { return _backendCode; }
            set { _backendCode = value; }
        }

        public LT_LandCountryInfo()
        {
        }

        public LT_LandCountryInfo(string name, string backendCode)
        {
            Name = name;
            BackendCode = backendCode;
        }
    }
}
