using System;
using System.IO;
using Microsoft.MetadirectoryServices;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Xml.Linq;
using System.Xml;
using System.Text;
using System.Net;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Vigo.Bas.ManagementAgent.Log;
using System.Diagnostics;

namespace FimSync_Ezma
{
    public class EzmaExtension :
    IMAExtensible2CallExport,
    IMAExtensible2CallImport,
    //IMAExtensible2FileImport,
    //IMAExtensible2FileExport,
    //IMAExtensible2GetHierarchy,
    IMAExtensible2GetSchema,
    IMAExtensible2GetCapabilities,
    IMAExtensible2GetParameters
    //IMAExtensible2GetPartitions
    {
        //KeyedCollection<string, ConfigParameter> exportConfigParameters;

        #region Page Size
        // Variables for page size
        private int m_importDefaultPageSize = 12;
        private int m_importMaxPageSize = 50;
        private int m_exportDefaultPageSize = 10;
        private int m_exportMaxPageSize = 20;
        KeyedCollection<string, ConfigParameter> exportconfigParameters;
        public string MachineName = Environment.GetEnvironmentVariable("COMPUTERNAME");
        public string outXml;

        public Dictionary<string, List<XElement>> positions = new Dictionary<string, List<XElement>>();
        public Dictionary<string, string> leadersID = new Dictionary<string, string>();
        Dictionary<string, string> _personOU = new Dictionary<string, string>();

        List<XElement> persons = new List<XElement>();
        public List<XElement> orgs = new List<XElement>();
        List<string> allreadyimported = new List<string>();

        // Bulids CN and OU strings
        NameValueCollection parent = new NameValueCollection();
        NameValueCollection OrgName = new NameValueCollection();

        public Dictionary<string, Tuple<string, string, string>> hiracyParents = new Dictionary<string, Tuple<string, string, string>>();
        
        // Groups Dictionary
        public Dictionary<string, Tuple<List<string>, string>> groups = new Dictionary<string, Tuple<List<string>, string>>();
        XmlWriterSettings xmlsettings = new XmlWriterSettings();
        bool debug;
        

        public string _ssn;


        public int ImportMaxPageSize
        {
            get
            {
                return m_importMaxPageSize;
            }
        }
        public int ImportDefaultPageSize
        {
            get
            {
                return m_importDefaultPageSize;
            }
        }
        public int ExportDefaultPageSize
        {
            get
            {
                return m_exportDefaultPageSize;
            }
            set
            {
                m_exportDefaultPageSize = value;
            }
        }
        public int ExportMaxPageSize
        {
            get
            {
                return m_exportMaxPageSize;
            }
            set
            {
                m_exportMaxPageSize = value;
            }
        }

        #endregion

        #region Capabilities, Config parameters and Schema
        public EzmaExtension()
        {
        }

        public MACapabilities Capabilities
        {
            // Returns the capabilities that this management agent has
            get
            {
                MACapabilities myCapabilities = new MACapabilities();

                myCapabilities.ConcurrentOperation = true;
                myCapabilities.ObjectRename = false;
                myCapabilities.DeleteAddAsReplace = false;
                myCapabilities.DeltaImport = false;
                myCapabilities.DistinguishedNameStyle = MADistinguishedNameStyle.None;
                myCapabilities.FullExport = true;
                myCapabilities.ExportType = MAExportType.ObjectReplace;
                myCapabilities.NoReferenceValuesInFirstExport = true;
                myCapabilities.Normalizations = MANormalizations.None;
                myCapabilities.ObjectConfirmation = MAObjectConfirmation.NoDeleteConfirmation;

                return myCapabilities;
            }
        }

        public IList<ConfigParameterDefinition> GetConfigParameters(KeyedCollection<string, ConfigParameter> configParameters, ConfigParameterPage page)
        {
            List<ConfigParameterDefinition> configParametersDefinitions = new List<ConfigParameterDefinition>();
            
            switch (page)
            {
                case ConfigParameterPage.Connectivity:
                    // Parametere for webservice
                    
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateStringParameter("GetAnsattList uri", "", @"https://bsbqa.bluegarden.net/Synchronous/GetAnsattList/v31/GetAnsattList"));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateStringParameter("GetOrgList uri", "", @"https://bsbqa.bluegarden.net/Synchronous/Organisation/Structure/v31/OrgStructure"));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateStringParameter("Eksport uri", "", @"https://bsbqa.bluegarden.net/Synchronous/GetAnsattList/v31/UpdateAnsatt"));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateStringParameter("Timeout (sek)", "", "600"));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateCheckBoxParameter("Enable Debug", false));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateDividerParameter());
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateStringParameter("OrgUnitId", "", "0000000000"));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateCheckBoxParameter("Kun aktive OrgUnits", true));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateLabelParameter("0000000000 henter alle personer fra alle org"));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateStringParameter("SourceCompany", "", "212"));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateStringParameter("SourceSystem", "", "HFK")); 
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateStringParameter("Arbeidsgiver", "", "2121200"));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateDividerParameter());
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateCheckBoxParameter("Lag grupper fra OrgList", true));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateStringParameter("Gruppeprefix", "", "s"));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateStringParameter("Gruppesuffix", "", ""));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateDividerParameter());

                    // Username and Password
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateStringParameter("Username", "", "bsbclient_212"));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateEncryptedStringParameter("Password", "", ""));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateStringParameter("Domain", ""));

                    break;
                case ConfigParameterPage.Global:
                case ConfigParameterPage.Partition:
                case ConfigParameterPage.RunStep:
                    break;
            }

            return configParametersDefinitions;
        }

        public ParameterValidationResult ValidateConfigParameters(KeyedCollection<string, ConfigParameter> configParameters, ConfigParameterPage page)
        {

            // Configuration validation

            ParameterValidationResult myResults = new ParameterValidationResult();

            // Try webservices, stop if unavaiable
            try
            {
                String AllPersons = ValidateWebservice(configParameters, "getAnsattList");
                
                if (AllPersons.Contains("soapenv:Server"))
                {
                    myResults.Code = ParameterValidationResultCode.Success;
                }
                else
                {
                    myResults.ErrorMessage = "Cannot Connect to WebService";
                    myResults.Code = ParameterValidationResultCode.Failure;
                }
            }
            catch (Exception e)
            {
                Logger.Log.Error("Error message: " + e.Message);
                myResults.ErrorMessage = "Error message: " + e.Message + Environment.NewLine + "Kan ikke koble til webservice";
                myResults.Code = ParameterValidationResultCode.Failure;
                
            }

            return myResults;
        }

        public Schema GetSchema(KeyedCollection<string, ConfigParameter> configParameters)
        {
            // Create CS Schema type person
            SchemaType person = SchemaType.Create("person", true);

            // Anchor
            person.Attributes.Add(SchemaAttribute.CreateAnchorAttribute("PersonID", AttributeType.String));

            // Attributes
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("PersonOU", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("UserRole", AttributeType.String, AttributeOperation.ImportOnly)); 
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("StartDato_person", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("SluttDato_person", AttributeType.String, AttributeOperation.ImportOnly)); 
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("Signatur", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("UserId", AttributeType.String, AttributeOperation.ExportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("Email", AttributeType.String, AttributeOperation.ExportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("Fornavn", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("Etternavn", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("Mobiltelefon", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("Telefon_Bosted", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("Telefon_Intern", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("AnsattNummer", AttributeType.String, AttributeOperation.ImportExport));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("Fodselsnummer", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("Arbeidsgiver", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("ArbeidsgiverNummer", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("AdresseLinje1_Bosted", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("AdresseLinje2_Bosted", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("AdresseLinje1_Intern", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("AdresseLinje2_Intern", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("Postnummer_Bosted", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("Poststed_Bosted", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("Postnummer_Intern", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("Poststed_Intern", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("Leder", AttributeType.Reference, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("Pri_Arbeidsforhold_status", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("Pri_Arbeidsforhold_nummer", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("Pri_AnsattforholdsKode", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("Pri_OrgUnitName", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("Pri_OrgUnitId", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("Pri_OrgUnitIdRef", AttributeType.Reference, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("Pri_IsHovedarbeidsforhold", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("Pri_Startdato_stilling", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("Pri_Stoppdato_stilling", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("Pri_Stillingskode_Kodeverdi", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("Pri_Stillingskode_Beskrivelse", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateMultiValuedAttribute("Sec_Arbeidsforhold_nummer", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateMultiValuedAttribute("Sec_OrgUnitName", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateMultiValuedAttribute("Sec_OrgUnitId", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateMultiValuedAttribute("Sec_Startdato_stilling", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateMultiValuedAttribute("Sec_AnsattforholdsKode", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateMultiValuedAttribute("Sec_Stillingskode_Kodeverdi", AttributeType.String, AttributeOperation.ImportOnly));
            person.Attributes.Add(SchemaAttribute.CreateMultiValuedAttribute("Sec_Stillingskode_Beskrivelse", AttributeType.String, AttributeOperation.ImportOnly));
            
            // Create CS type group
            SchemaType group = SchemaType.Create("group", true);

            // Anchor
            group.Attributes.Add(SchemaAttribute.CreateAnchorAttribute("groupId", AttributeType.String, AttributeOperation.ImportOnly));

            // Attributes
            group.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("groupName", AttributeType.String, AttributeOperation.ImportOnly));
            group.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("groupSourceId", AttributeType.String, AttributeOperation.ImportOnly));
            group.Attributes.Add(SchemaAttribute.CreateMultiValuedAttribute("members", AttributeType.Reference, AttributeOperation.ImportOnly));
            group.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("groupOU", AttributeType.String, AttributeOperation.ImportOnly));


            // Create CS type unit
            SchemaType unit = SchemaType.Create("unit", true);

            // Anchor
            unit.Attributes.Add(SchemaAttribute.CreateAnchorAttribute("orgunitid", AttributeType.String, AttributeOperation.ImportOnly));

            // Attributes
            unit.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("OrgNavn", AttributeType.String, AttributeOperation.ImportOnly));
            unit.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("UnitId", AttributeType.String, AttributeOperation.ImportOnly));
            unit.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("ParentGroupId", AttributeType.String, AttributeOperation.ImportOnly));
            unit.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("ErAktiv", AttributeType.String, AttributeOperation.ImportOnly));
            unit.Attributes.Add(SchemaAttribute.CreateSingleValuedAttribute("unitOU", AttributeType.String, AttributeOperation.ImportOnly)); 


            // Return schema
            Schema schema = Schema.Create();
            schema.Types.Add(person);
            schema.Types.Add(group);
            schema.Types.Add(unit);

            return schema;
        }
        #endregion

        #region Get Infile methods

        //private String GetXmlfile(KeyedCollection<string, ConfigParameter> configParameters)
        //{
        //    try
        //    {   //Log request
        //        string infile = configParameters["Mappe"].Value + configParameters["Inn xml"].Value;
        //        Log("Trying to read: " + infile);
        //        FileStream fileStream = new FileStream(infile, FileMode.Open, FileAccess.Read);
        //        StreamReader reader = new StreamReader(fileStream);

        //        // Read all content from stream
        //        string xml = reader.ReadToEnd();

        //        // Cleanup
        //        reader.Close();
        //        fileStream.Close();

        //        // Return, but remove the namespace
        //        return xml;
        //    }
        //    catch (Exception e)
        //    {
        //        throw new Exception("Message: " + e.Message + " StackTrace: " + e.StackTrace);
        //    }
        //}

        //private String GetCsvfile(KeyedCollection<string, ConfigParameter> configParameters)
        //{
        //    try
        //    {   //Log request
        //        string csvfile = configParameters["Mappe"].Value + configParameters["Inn csv"].Value;
        //        Log("Leser: " + csvfile);

        //        string[] lines = File.ReadAllLines(csvfile, Encoding.GetEncoding("ISO-8859-1"));

        //        XDocument csvtoxml = new XDocument
        //            (new XElement
        //                ("units", from str in lines.Where((line, index) => index > 0)
        //                    let columns = str.Split(';')
        //                    select new XElement("unit",
        //                        new XElement("ID", columns[0]),
        //                        new XElement("Navn", columns[1]),
        //                        new XElement("Eier_ID", columns[2]),
        //                        new XElement("Leder", columns[3]),
        //                        new XElement("Aktiv", columns[4]),
        //                        new XElement("Kostnader", columns[5]),
        //                        new XElement("Personer", columns[6]),
        //                        new XElement("Orgnr_bedrift", columns[7]),
        //                        new XElement("Orgledd", columns[8]),
        //                        new XElement("Kostnadsbaerer", columns[9]))));

        //        string xml = csvtoxml.ToString(SaveOptions.DisableFormatting);
        //        return xml;
        //    }
        //    catch (Exception e)
        //    {
        //        throw new Exception("Message: " + e.Message + " StackTrace: " + e.StackTrace);
        //    }
        //}


        #endregion


        #region Private SOAP and WS methods
        private NetworkCredential GetNetworkCredential(KeyedCollection<string, ConfigParameter> configParameters)
        {

            string username = null;
            //Collect username and password, domain and return to connect to WS
            if (configParameters["Username"].Value != "")
            {
                if (configParameters["Domain"].Value != null && configParameters["Domain"].Value != "")
                {
                    username = configParameters["Domain"].Value + "\\" + configParameters["Username"].Value;
                }
                else
                {
                    username = configParameters["Username"].Value;
                }
            }
            return new NetworkCredential(username, configParameters["Password"].SecureValue);
        }
        

        private String GetAllActivePersons(KeyedCollection<string, ConfigParameter> configParameters, string OrgUnitId)
        {
            
            string SOAPAction = "getAnsattList";
            string tm = DateTime.Now.ToString("yyyyMMdd-HHmmssfff"),
                SourceCompany = configParameters["SourceCompany"].Value,
                SourceSystem = configParameters["SourceSystem"].Value,
                Arbeidsgiver = configParameters["Arbeidsgiver"].Value,
                SOAPxml = @"
               <soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/""
               xmlns:v3=""http://bluemsg.bluegarden.no/object/v3""
               xmlns:v31=""http://ansattlist.bluegarden.no/service/v3_1"">
                  <soapenv:Header>
                    <v3:BSBHeader>
                     <SourceCompany>" + SourceCompany + @"</SourceCompany>
                     <SourceEmployer>*</SourceEmployer>
                     <SourceSystem>" + SourceSystem + @"</SourceSystem>
                     <SourceUser>" + MachineName + @"</SourceUser>
                     <MessageId>" + tm + @"</MessageId>
                  </v3:BSBHeader>
               </soapenv:Header>
               <soapenv:Body>
                  <v31:GetAnsattListRequestMessage>
	             <v31:Arbeidsgiver>" + Arbeidsgiver + @"</v31:Arbeidsgiver>
                     <v31:OrgUnitId>" + OrgUnitId + @"</v31:OrgUnitId>
                  </v31:GetAnsattListRequestMessage>
               </soapenv:Body>
            </soapenv:Envelope>";
            if (debug == true)
            {
                Logger.Log.Debug("Henter personer fra " + configParameters["GetAnsattList uri"].Value + " under enhet " + OrgUnitId);
            }

            // Send generic request for all users with credentials, and return the string containing the XML result
            return GenericRequest(
                configParameters["GetAnsattList uri"].Value,
                GetNetworkCredential(configParameters),
                "POST",
                SOAPxml,
                int.Parse(configParameters["Timeout (sek)"].Value), SOAPAction);

        }
        
        private String UpdatePerson(KeyedCollection<string, ConfigParameter> configparameters, string SOAPxml)
        {
            return GenericRequest(
                configparameters["Eksport uri"].Value,
                GetNetworkCredential(configparameters),
                "POST", SOAPxml,
                int.Parse(configparameters["Timeout (sek)"].Value), "updatePerson");
        }
        private String GetAllOrg(KeyedCollection<string, ConfigParameter> configParameters)
        {
            string tm = DateTime.Now.ToString("yyyyMMdd-HHmmssfff"),
                SourceCompany = configParameters["SourceCompany"].Value, 
                SourceSystem = configParameters["SourceSystem"].Value,
                Arbeidsgiver = configParameters["Arbeidsgiver"].Value,
                OrgUnitId = configParameters["OrgUnitId"].Value,
                SOAPAction = "getOrgList",
                SOAPxml = @"
            <soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" 
            xmlns:he=""http://bluemsg.bluegarden.no/object/v3""
            xmlns:v31=""http://bluegarden.no/organisation/structure/service/v31"">
               <soapenv:Header>
                  <he:BSBHeader>
                     <SourceCompany>" + SourceCompany + @"</SourceCompany>
                     <SourceEmployer>*</SourceEmployer>
                     <SourceSystem>" + SourceSystem + @"</SourceSystem>
                     <SourceUser>" + MachineName + @"</SourceUser>
                     <MessageId>" + tm + @"</MessageId>
                  </he:BSBHeader>
               </soapenv:Header>
               <soapenv:Body>
                  <v31:getOrgListRequest>
	             <v31:Arbeidsgiver>" + Arbeidsgiver + @"</v31:Arbeidsgiver>
                     <v31:OrgUnitId>" + OrgUnitId + @"</v31:OrgUnitId>
                  </v31:getOrgListRequest>
               </soapenv:Body>
            </soapenv:Envelope>";
            Logger.Log.Info("Henter organisasjonsinfo fra " + configParameters["GetOrgList uri"].Value);

            // Send generic request for all users with credentials, and return the string containing the XML result
            return GenericRequest(
                configParameters["GetOrgList uri"].Value,
                GetNetworkCredential(configParameters),
                "POST", SOAPxml,
                int.Parse(configParameters["Timeout (sek)"].Value), SOAPAction);
        }


        private String ValidateWebservice(KeyedCollection<string, ConfigParameter> configParameters, string SOAPAction)
        {
            // Send generic request for all users with credentials, and return the string containing the XML result
            
            return GenericRequest(configParameters["GetAnsattList uri"].Value,
                GetNetworkCredential(configParameters),
                "GET",
                null,
                int.Parse(configParameters["Timeout (sek)"].Value), SOAPAction);

        }



        private String GenericRequest(String uri, NetworkCredential credentials, String method, String SOAPxmlBody, int timeoutSec, String SOAPAction)
        {
            
            string soapResult = null;

            try
            {
                // Log request debug
                if (debug == true)
                {
                    Logger.Log.Debug("Request: " + uri);
                    Logger.Log.Debug("Request method: " + method);
                    Logger.Log.Debug("SOAPAction: " + SOAPAction);
                }


                // Create request and assign credentials
                HttpWebRequest webRequest = (HttpWebRequest)WebRequest.Create(uri);

                if (credentials != null)
                {
                    webRequest.Credentials = credentials;
                }
                // Even If certificates not is valid, it moves on
                System.Net.Security.RemoteCertificateValidationCallback BadCertificates = new System.Net.Security.RemoteCertificateValidationCallback(delegate { return true; });
                System.Net.ServicePointManager.ServerCertificateValidationCallback = BadCertificates;

                if (method == "POST")
                {
                    // Send POST body
                    webRequest.ContentType = "text/xml;charset=\"utf-8\"";
                    webRequest.Accept = "text/xml";
                    webRequest.Method = method;
                    webRequest.Timeout = timeoutSec * 1000;
                    webRequest.KeepAlive = true;
                    webRequest.Headers.Add(@"SOAPAction:" + SOAPAction);

                    XmlDocument soapEnvelopeXml = new XmlDocument();
                    soapEnvelopeXml.LoadXml(SOAPxmlBody);

                    using (Stream stream = webRequest.GetRequestStream())
                    {
                        soapEnvelopeXml.Save(stream);
                    }

                    using (WebResponse response = webRequest.GetResponse())
                    {
                        using (StreamReader rd = new StreamReader(response.GetResponseStream()))
                        {
                            soapResult = rd.ReadToEnd();
                            rd.Close();
                        }
                    }
                }
                else
                {
                    webRequest.Method = method;
                    webRequest.Timeout = timeoutSec * 1000;
                    webRequest.PreAuthenticate = true;

                    try
                    {
                        using (WebResponse response = webRequest.GetResponse())
                        {
                            using (StreamReader rd = new StreamReader(response.GetResponseStream()))
                            {
                                soapResult = rd.ReadToEnd();
                                rd.Close();
                            }
                        }

                    }
                    catch (WebException e)
                    {
                        if (e.Message.Contains("500"))
                        {
                            using (var stream = e.Response.GetResponseStream())
                            using (var reader = new StreamReader(stream))
                            {
                                soapResult = reader.ReadToEnd();
                            }
                        }
                        else
                        {
                            soapResult = @"<soapenv:Envelope xmlns:soapenv = ""http://schemas.xmlsoap.org/soap/envelope/""
                            <soapenv:Body>
                            <soapenv:Fault>
                            <faultcode>error</faultcode>
                            </soapenv:Fault>
                            </soapenv:Body>
                            </soapenv:Envelope>";
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Log.Error(string.Format("Errorcode: {0}", ex.Message));
                    }
                }
            }
            catch (WebException e)
            {
                Logger.Log.Error("Message: " + e.Message);
                if (SOAPAction == "getOrgList" && method == "POST")
                {
                    throw new Exception("Kan ikke hente ut organisajoner. Webservice svarte ikke innen timeout");
                }

            }
            
            return soapResult;
            
        }
        #endregion

        public String GetOUString(string GroupId, string OrgNavn)
        {
            OrgNavn = Regex.Replace(OrgNavn, @"[^a-zA-Z0-9ÆØÅæøå\-_ ]", "").TrimStart().TrimEnd();
            string parentstring = null;
            while (parent[GroupId] != null && parent[GroupId] != "2121200/0000000000")
            {
                if (parentstring != null)
                {
                    parentstring += ",OU=";
                }
                parentstring += OrgNavn;
                GroupId = parent[GroupId];
                OrgNavn = OrgName[GroupId];
            }
            if (parentstring != null)
            {
                parentstring = "OU=" + parentstring;
                parentstring += ",OU=";
                parentstring += OrgName[GroupId];
            }
            
            return parentstring;
        }

        


        #region Import methods
        /*
         * Attributes used during import 
         */
        List<ImportListItem> ImportedObjectsList;
        int GetImportEntriesIndex, GetImportEntriesPageSize;

        public OpenImportConnectionResults OpenImportConnection(KeyedCollection<string, ConfigParameter> configParameters,
                                                Schema types,
                                                OpenImportConnectionRunStep importRunStep)
        {
            Logger.Log.Info("Starter import");

            // Instantiate ImportedObjectsList
            ImportedObjectsList = new List<ImportListItem>();

            // Debug true/false
            if (configParameters["Enable Debug"].Value == "1")
            {
                debug = true;
            }

            // Find only active units true/false
            bool _onlyActiveUnits = false;
            if (configParameters["Kun aktive OrgUnits"].Value == "1")
            {
                _onlyActiveUnits = true;
            }

            #region Get all units
            Stopwatch getOrganizationsStopwatch = new Stopwatch();
            getOrganizationsStopwatch.Start();

            // Get all orgunits
            String _GetAllOrg = GetAllOrg(configParameters);

            getOrganizationsStopwatch.Stop();
            if (debug == true)
            {
                Logger.Log.Debug(string.Format("GetOrganizations fullførte på {0} sekunder", getOrganizationsStopwatch.Elapsed.Seconds));
            }
            // Define XML Namespaces
            XNamespace se = "http://schemas.xmlsoap.org/soap/envelope/";
            XNamespace v3person = "http://ansattlist.bluegarden.no/service/v3_1";
            XNamespace v31person = "http://ansattlist.bluegarden.no/object/v3_1";
            XNamespace v31 = "http://bluegarden.no/organisation/structure/service/v31";
            XNamespace v311 = "http://bluegarden.no/organisation/structure/object/v31";

            // Define NameValueCollection to be used with hirarchy
            //NameValueCollection parent = new NameValueCollection();
            //NameValueCollection OrgName = new NameValueCollection();

            
            // OrgUnits
            int orgcount = 0;
            Logger.Log.Info("Bygger OU-struktur");
            foreach (XElement unit in XDocument.Parse(_GetAllOrg).
                Element(se + "Envelope").
                Element(se + "Body").
                Element(v31 + "getOrgListResponse").
                Element(v311 + "OrgList").
                Elements(v311 + "OrgUnit"))

            {
                string GroupId = null;
                try
                {
                    GroupId = unit.Element(v311 + "GroupId").Value;
                    string ParentGroupId = unit.Element(v311 + "ParentGroupId").Value;
                    string _OrgName = Regex.Replace(unit.Element(v311 + "OrgNavn").Value, @"[^a-zA-Z0-9ÆØÅæøå\-_ ]", "").TrimStart().TrimEnd();
                    string ErAktiv = unit.Element(v311 + "ErAktiv").Value;

                    if (GroupId != "2121200/0000000000")
                    {
                        if (_onlyActiveUnits == true)
                        {
                            if (ErAktiv == "true")
                            {
                                parent[GroupId] = ParentGroupId;
                                OrgName[GroupId] = _OrgName;
                                orgcount++;
                                orgs.Add(unit);
                            }
                        }
                        else
                        {
                            parent[GroupId] = ParentGroupId;
                            OrgName[GroupId] = _OrgName;
                            orgcount++;
                            orgs.Add(unit);
                        }
                    }
                }
                catch (Exception e)
                {
                    Logger.Log.Error("Kan ikke bygge OU-struktur for " + GroupId);
                    Logger.Log.Error(e);
                }
                

            }
            Logger.Log.Info("OU-struktur ferdig");
            Logger.Log.Info(string.Format("GetOrganizations fant {0} enheter", orgcount));

            #endregion

            #region Get all persons in all units
            Stopwatch getPersonsPrOrgStopwatch = new Stopwatch();
            getPersonsPrOrgStopwatch.Start();

            int countperstotal = 0;

            // Get persons foreach orgunit
            foreach (XElement xunit in orgs)
            {
                string orgunitid = null;
                string ErAktiv = null;
                string _OrgName = null;
                string GroupId = null;
                string ParentGroupId = null;

                // Set strings with XML values foreach OrgUnit
                try
                {
                    orgunitid = xunit.Element(v311 + "OrgUnitId").Value;
                    _OrgName = xunit.Element(v311 + "OrgNavn").Value;
                    GroupId = xunit.Element(v311 + "GroupId").Value;
                    ParentGroupId = xunit.Element(v311 + "ParentGroupId").Value;
                    ErAktiv = xunit.Element(v311 + "ErAktiv").Value;
                }
                catch (Exception e)
                {
                    Logger.Log.Error("Feil med en eller flere verdier i XML for organisasjon " + orgunitid);
                    Logger.Log.Error(e.ToString());
                }

                // If defined in Agent Configuration
                if (_onlyActiveUnits == true)
                {
                    // Make sure not to import OrgUnit more than once and pick only Active Units
                    if (!allreadyimported.Contains(orgunitid) && ErAktiv == "true")
                    {
                        string oustring = GetOUString(GroupId, _OrgName);

                        // Parse as Organization
                        Unit unit = new Unit(xunit, oustring);
                        Logger.Log.Info("La til organisasjon: " + orgunitid);
                        ImportedObjectsList.Add(new ImportListItem() { unit = unit });
                        allreadyimported.Add(orgunitid);
                    }
                }
                else
                {
                    // Make sure not to import OrgUnit more than once
                    if (!allreadyimported.Contains(orgunitid))
                    {
                        string OUstring = GetOUString(GroupId, _OrgName);

                        // Parse as Organization
                        Unit unit = new Unit(xunit, OUstring);
                        Logger.Log.Info("La til organisasjon: " + orgunitid);
                        ImportedObjectsList.Add(new ImportListItem() { unit = unit });
                        allreadyimported.Add(orgunitid);
                    }
                }

                // Get and parse all persons in unit
                try
                {
                    Stopwatch getAllPersonsStopwatch = new Stopwatch();
                    getAllPersonsStopwatch.Start();

                    // Get users from WS
                    String _GetAllPersons = GetAllActivePersons(configParameters, orgunitid);
                    getAllPersonsStopwatch.Stop();
                    if (debug == true)
                    {
                        Logger.Log.Debug(string.Format("Webservice kallet GetAllPersons for Orgunit {0} fullførte på {1} sekunder", orgunitid, getAllPersonsStopwatch.Elapsed.Seconds));
                    }
                    int countpers = 0;
                    

                    // Parse all persons positions
                    foreach (XElement xperson in XDocument.Parse(_GetAllPersons).
                        Element(se + "Envelope").Element(se + "Body").
                        Element(v3person + "GetAnsattListResultMessage").
                        Element(v31person + "AnsattList").
                        Elements(v31person + "Ansatt"))
                    {
                        // Define strings
                        string PersonID, orgUnitId_pers;
                        string Leder = null;

                        // Set values to strings
                        PersonID = xperson.Element(v31person + "PersonID").Value.Replace("#", "");
                        orgUnitId_pers = xperson.Element(v31person + "Arbeidsforhold").Element(v31person + "OrgUnitId").Value;
                        try
                        {
                            Leder = xperson.Element(v31person + "Arbeidsforhold").Element(v31person + "Leder").Value;
                        }
                        catch
                        { }

                        // Parse all positions to 
                        if (!positions.ContainsKey(PersonID))
                        {
                            positions.Add(PersonID, new List<XElement>());
                        }
                        positions[PersonID].Add(xperson.Element(v31person + "Arbeidsforhold"));
                        persons.Add(xperson);
                        if (Leder != null)
                        if(!leadersID.ContainsKey(Leder))
                        {
                            leadersID.Add(Leder, null);
                        }

                        countpers++;
                        countperstotal++;
                    }
                    Logger.Log.Info(string.Format("Fant {0} personer i OrgUnit {1}", countpers, orgunitid));
                }
                catch (Exception e)
                {
                    Logger.Log.Error("GetAnsattList for OrgUnit " + orgunitid + " feilet");
                    if (debug == true)
                    {
                        Logger.Log.Debug(e.ToString());
                    }
                }
                
            }
            Logger.Log.Info(string.Format("Fant {0} stillinger", countperstotal));

            #endregion
            
            // Create leaders refrences
            foreach (XElement xperson in persons)
            {
                string PersonID = null, Signatur = null, Leder = null, GroupId, OrgUnitName, isHovedarb = null ,isaktivarb = null;
                
                PersonID = xperson.Element(v31person + "PersonID").Value.Replace("#", "");
                try
                {
                    isaktivarb = xperson.Element(v31person + "Arbeidsforhold").Attribute("status").Value;
                    isHovedarb = xperson.Element(v31person + "Arbeidsforhold").Element(v31person + "IsHovedarbeidsforhold").Value;

                    // Adds just active emplyments and primary ones to _personOU dictionary for later use
                    if (isaktivarb == "Aktiv" && isHovedarb == "J")
                    {
                        OrgUnitName = xperson.Element(v31person + "Arbeidsforhold").Element(v31person + "OrgUnitName").Value;
                        GroupId = xperson.Element(v31person + "Arbeidsforhold").Element(v31person + "OrgUnitId").Value;
                        GroupId = configParameters["Arbeidsgiver"].Value + @"/" + string.Format("{0:D10}", Convert.ToInt64(GroupId));

                        string OUstring = GetOUString(GroupId, OrgUnitName);
                        if (!_personOU.ContainsKey(PersonID))
                        {
                            _personOU.Add(PersonID, OUstring);
                        }
                    }
                }
                catch
                {
                    if (debug == true)
                    {
                        Logger.Log.Debug(string.Format("kunne ikke lage PersonOU for {0}", PersonID));
                    }
                }

                try
                {
                    Signatur = xperson.Element(v31person + "Signatur").Value;
                    Leder = xperson.Element(v31person + "Arbeidsforhold").Element(v31person + "Leder").Value;

                    if (Signatur != null && Leder != null)
                    {
                        foreach (KeyValuePair<string, string> leader in leadersID)
                        {
                            if (leader.Key == Signatur)
                            {
                                leadersID[Signatur] = PersonID;
                            }
                        }
                    }


                }
                catch
                {
                    if (debug == true)
                    {
                        Logger.Log.Debug(string.Format("kunne ikke legge til lederreferanse for {0}", PersonID));
                    }
                }

        }

            Logger.Log.Info("Legger til personer og ansattforhold");
            // Parse all elements in Ansatt with all positions
            foreach (XElement xperson in persons)
            {
                string PersonID = null;
                try
                {
                    // Define strings
                    string Fornavn, Etternavn, OrgUnitName; //, GroupId;

                    // Set values to strings
                    PersonID = xperson.Element(v31person + "PersonID").Value.Replace("#", "");
                    Fornavn = xperson.Element(v31person + "Fornavn").Value;
                    Etternavn = xperson.Element(v31person + "Etternavn").Value;
                    OrgUnitName = xperson.Element(v31person + "Arbeidsforhold").Element(v31person + "OrgUnitName").Value;
                    //GroupId = xperson.Element(v31person + "Arbeidsforhold").Element(v31person + "OrgUnitId").Value;
                    //GroupId = configParameters["Arbeidsgiver"].Value + @"/" + string.Format("{0:D10}", Convert.ToInt64(GroupId));

                    //string OUstring = GetOUString(GroupId, OrgUnitName);

                    // Try get _personOU value
                    

                    // Make sure to only import person object once
                    if (!allreadyimported.Contains(PersonID))
                    {
                        // Parse as Person

                        string OUstring;
                        _personOU.TryGetValue(PersonID, out OUstring);

                        Person person = new Person(xperson, groups, positions, OUstring, leadersID);

                        if (debug == true)
                        {
                            Logger.Log.Debug(string.Format("La til person: ({0}) {1} {2}", PersonID, Fornavn, Etternavn));
                        }
                        ImportedObjectsList.Add(new ImportListItem() { person = person });
                        allreadyimported.Add(PersonID);
                    }

                    if (debug == true)
                    {
                        Logger.Log.Debug(string.Format("La til info om arbeidssted for person: ({0}) {1} {2}", PersonID, Fornavn, Etternavn));
                    }
                }
                catch (Exception e)
                {
                    Logger.Log.Error(PersonID + " ble ikke lagt til p.g.a en feil");
                    if (debug == true)
                    {
                        Logger.Log.Debug(e.ToString());
                    }
                }
            }
            
            getPersonsPrOrgStopwatch.Stop();
            Logger.Log.Info(string.Format("Henting av alle personer fra alle orgunits fullførte på {0} minutter", getPersonsPrOrgStopwatch.Elapsed.Minutes));

            #region Groups Creation
            // Create groups from Unit values from person
            if (configParameters["Lag grupper fra OrgList"].Value == "1")
            {
                int grpcount = 0;
                Logger.Log.Info("Legger til grupper");
                try
                {
                    foreach (var _group in groups)
                    {
                        string orgunitid = configParameters["Arbeidsgiver"].Value + @"/" + string.Format("{0:D10}", Convert.ToInt64(_group.Key.ToString()));
                        string OUstring = GetOUString(orgunitid, _group.Value.Item2);

                        // Parse groups; key, members, Name, Description, OU-value
                        Group group = new Group(_group.Key, _group.Value.Item1, _group.Value.Item2, configParameters, OUstring);
                        if (debug == true)
                        {
                            Logger.Log.Debug("La til gruppe: " + configParameters["Gruppeprefix"].Value + _group.Key + configParameters["Gruppesuffix"].Value);
                        }
                        ImportedObjectsList.Add(new ImportListItem() { group = group });
                        grpcount++;
                    }
                    Logger.Log.Info(string.Format("Laget {0} grupper", grpcount));
                }
                catch (Exception e)
                {
                    Logger.Log.Error("Kunne ikke lage grupper. Feil " + e);
                    throw new Exception("Feilet på gruppegenerering. Feil " + e);
                }
                finally
                {
                    Logger.Log.Info("Gruppeimport ferdig");
                }
            }
            
            #endregion
            // Set index values and page size
            GetImportEntriesIndex = 0;
            GetImportEntriesPageSize = importRunStep.PageSize;
            return new OpenImportConnectionResults();
        }

        public CloseImportConnectionResults CloseImportConnection(CloseImportConnectionRunStep importRunStepInfo)
        {
            Logger.Log.Info("Import ferdig");
            return new CloseImportConnectionResults();
        }

        public GetImportEntriesResults GetImportEntries(GetImportEntriesRunStep importRunStep)
        {
            /* This method will be invoked multiple times, once for each "page" */

            List<CSEntryChange> csentries = new List<CSEntryChange>();
            GetImportEntriesResults importReturnInfo = new GetImportEntriesResults();

            // If no result, return empty success
            if (ImportedObjectsList == null || ImportedObjectsList.Count == 0)
            {
                importReturnInfo.CSEntries = csentries;
                return importReturnInfo;
            }
            // Parse a full page or to the end
            for (int currentPage = 0;
                GetImportEntriesIndex < ImportedObjectsList.Count && currentPage < GetImportEntriesPageSize;
                GetImportEntriesIndex++, currentPage++)
            {
                // Add persons to CsEntry
                if (ImportedObjectsList[GetImportEntriesIndex].person != null)
                {
                    csentries.Add(ImportedObjectsList[GetImportEntriesIndex].person.GetCSEntryChange());
                }

                // Add groups to CsEntry
                if (ImportedObjectsList[GetImportEntriesIndex].group != null)
                {
                    csentries.Add(ImportedObjectsList[GetImportEntriesIndex].group.GetCSEntryChange());
                }

                // Add Units to CsEntry
                if (ImportedObjectsList[GetImportEntriesIndex].unit != null)
                {
                    csentries.Add(ImportedObjectsList[GetImportEntriesIndex].unit.GetCSEntryChange());
                }
            }
            

            // Store and return
            importReturnInfo.CSEntries = csentries;
            importReturnInfo.MoreToImport = GetImportEntriesIndex < ImportedObjectsList.Count;
            return importReturnInfo;
        }

        #endregion

        #region Export methods
        public void OpenExportConnection(KeyedCollection<string, ConfigParameter> configParameters,
                              Schema types,
                              OpenExportConnectionRunStep exportRunStep)
        {
            Logger.Log.Info("Starter Export");
            exportconfigParameters = configParameters;
            return;
        }

        public PutExportEntriesResults PutExportEntries(IList<CSEntryChange> csentries)
        {
            StringWriter StringWriter;
            xmlsettings.Indent = true;
            xmlsettings.Encoding = Encoding.GetEncoding("iso-8859-1");
            
            foreach (CSEntryChange csentryChange in csentries)
            {
                StringWriter = new StringWriter();
                XmlWriter xmlWriter = XmlWriter.Create(StringWriter, xmlsettings);
                switch (csentryChange.ObjectType)
                {
                    case "person":

                        switch (csentryChange.ObjectModificationType)
                        {
                            case ObjectModificationType.Add:
                            case ObjectModificationType.Delete:
                            case ObjectModificationType.Replace:
                            case ObjectModificationType.Update:
                                try
                                {
                                    const string Ns = "http://ansatt.bluegarden.no/updateansatt/service/v3";
                                    const string Prefix = "v31";

                                    xmlWriter.WriteStartDocument(true);
                                    xmlWriter.WriteStartElement(Prefix, "UpdateAnsattRequest", Ns);
                                    xmlWriter.WriteAttributeString("xmlns", Prefix, null, Ns);

                                    xmlWriter.WriteStartElement(Prefix, "AnsattNummer", null);
                                    xmlWriter.WriteString(csentryChange.DN.Substring(7,5));
                                    xmlWriter.WriteEndElement();

                                    foreach (var attribute in csentryChange.ChangedAttributeNames)
                                    {
                                        switch (attribute.ToLower())
                                        {
                                            case "email":
                                            case "userid":

                                                string attrvalue = csentryChange.AttributeChanges[attribute].ValueChanges[0].Value.ToString();
                                                xmlWriter.WriteStartElement(Prefix, attribute, null);
                                                xmlWriter.WriteString(attrvalue);
                                                xmlWriter.WriteEndElement();

                                                break;
                                        }
                                    }
                                }
                                catch
                                { }
                                finally
                                {
                                    xmlWriter.WriteEndElement();
                                    xmlWriter.WriteEndDocument();
                                }

                                break;
                        }
                        break;
                }
                xmlWriter.Close();
                xmlWriter.Dispose();
                outXml = StringWriter.ToString();
                
                if (outXml != string.Empty)
                {
                    UpdatePerson(exportconfigParameters, outXml);
                }
                
            }
            
            PutExportEntriesResults exportEntriesResults = new PutExportEntriesResults();

            return exportEntriesResults;
        }

        public void CloseExportConnection(CloseExportConnectionRunStep exportRunStep)
        {
            Logger.Log.Info("Eksport Ferdig");
        }
        #endregion
    };
}
