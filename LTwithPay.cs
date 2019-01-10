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
        }


        public int GetTheFormLanguageCode()
        {
            int cc = 0;

            if (FormState.Contains("Language"))
            {
                if (FormState["Language"].Equals(1044)) cc =  1044;
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
            catch (Exception e) 
            {
                return new ValidationResult(false, "Feil ved uthenting av data, prøv igjen senere");
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

        public void SetGuiCtrlForEmptyErData(string fieldValue, string guiNodeXpath)
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
            SetGuiCtrlForEmptyErData(data.BusinessAddress, "/uictrl/adresse");

            SetNodeToString("/melding/Organisasjon/e-post", data.EMailAddress, nullMelding);
            SetGuiCtrlForEmptyErData(data.EMailAddress, "/uictrl/e-post");

            LT_LandCountryInfo countryName = FindAltinnCodeListNameForCountryCode(data.CountryCode);
            if (countryName != null) 
            {
                SetNodeToString("/melding/Organisasjon/land", countryName.BackendCode, nullMelding);
                SetGuiCtrlNode("/uictrl/land_display_field", countryName.Name);
                SetGuiCtrlNode("/uictrl/land", "USETEXTBOX");
            }
            else
                SetGuiCtrlNode("/uictrl/land", "USEPULLDOWN");
            SetNodeToString("/melding/Organisasjon/land", null, null);
    
            SetNodeToString("/melding/Organisasjon/navn", data.Name, nullMelding);
            SetGuiCtrlForEmptyErData(data.Name, "/uictrl/navn");

            SetNodeToString("melding/Organisasjon/postnummer", data.BusinessPostCode, nullMelding);
            SetGuiCtrlForEmptyErData(data.BusinessPostCode, "/uictrl/postnummer");

            SetNodeToString("/melding/Organisasjon/poststed", data.BusinessPostCity, nullMelding);
            SetGuiCtrlForEmptyErData(data.BusinessPostCity, "/uictrl/poststed");

            SetNodeToString("/melding/Organisasjon/telefon", data.TelephoneNumber, nullMelding);
            SetGuiCtrlForEmptyErData(data.TelephoneNumber, "/uictrl/telefon");
        }


        public void organisasjonsnummer_Changed(object sender, XmlEventArgs e)
        {
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
