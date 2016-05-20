using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SourceCode.SmartObjects.Services.ServiceSDK.Objects;
using Attributes = SourceCode.SmartObjects.Services.ServiceSDK.Attributes;
using SourceCode.SmartObjects.Services.ServiceSDK.Types;
using System.Net;
using System.Globalization;
using System.IO;

namespace K2Field.SmartObjects.Services.MicrosoftGraphExample
{
    [Attributes.ServiceObject("MicrosoftGraphUser", "Microsoft Graph User", "Microsoft Graph User")]
    public class GraphUser
    {
        // This is an example service. 
        // It doesn't implement all User properties

        public ServiceConfiguration ServiceConfiguration { get; set; }

        public string odatacontext { get; set; }


        [Attributes.Property("Id", SoType.Text, "Id", "Id")]
        public string id { get; set; }

        [Attributes.Property("AccountEnabled", SoType.YesNo, "Account Enabled", "Account Enabled")]
        public bool accountEnabled { get; set; }

        public Assignedlicense[] assignedLicenses { get; set; }
        public Assignedplan[] assignedPlans { get; set; }
        public string[] businessPhones { get; set; }

        [Attributes.Property("City", SoType.Text, "City", "City")]
        public string city { get; set; }

        [Attributes.Property("CompanyName", SoType.Text, "Company Name", "Company Name")]
        public string companyName { get; set; }

        [Attributes.Property("Country", SoType.Text, "Country", "Country")]
        public string country { get; set; }

        [Attributes.Property("Department", SoType.Text, "Department", "Department")]
        public string department { get; set; }

        [Attributes.Property("DisplayName", SoType.Text, "Display Name", "Display Name")]
        public string displayName { get; set; }

        [Attributes.Property("GivenName", SoType.Text, "Given Name", "Given Name")]
        public string givenName { get; set; }

        [Attributes.Property("JobTitle", SoType.Text, "Job Title", "Job Title")]
        public string jobTitle { get; set; }

        [Attributes.Property("Mail", SoType.Text, "Mail", "Mail")]
        public string mail { get; set; }

        [Attributes.Property("MailNickname", SoType.Text, "Mail Nickname", "Mail Nickname")]
        public string mailNickname { get; set; }

        [Attributes.Property("MobilePhone", SoType.Text, "MobilePhone", "MobilePhone")]
        public string mobilePhone { get; set; }

        public string onPremisesImmutableId { get; set; }
        public string onPremisesLastSyncDateTime { get; set; }
        public string onPremisesSecurityIdentifier { get; set; }
        public string onPremisesSyncEnabled { get; set; }
        public string passwordPolicies { get; set; }
        public string passwordProfile { get; set; }

        [Attributes.Property("OfficeLocation", SoType.Text, "Office Location", "Office Location")]
        public string officeLocation { get; set; }

        [Attributes.Property("PostalCode", SoType.Text, "Postal Code", "Postal Code")]
        public string postalCode { get; set; }

        [Attributes.Property("PreferredLanguage", SoType.Text, "Preferred Language", "Preferred Language")]
        public string preferredLanguage { get; set; }

        public Provisionedplan[] provisionedPlans { get; set; }
        public string[] proxyAddresses { get; set; }

        public DateTime refreshTokensValidFromDateTime { get; set; }

        [Attributes.Property("State", SoType.Text, "State", "State")]
        public string state { get; set; }

        [Attributes.Property("StreetAddress", SoType.Text, "Street Address", "Street Address")]
        public string streetAddress { get; set; }

        [Attributes.Property("Surname", SoType.Text, "Surname", "Surname")]
        public string surname { get; set; }

        [Attributes.Property("UsageLocation", SoType.Text, "Usage Location", "Usage Location")]
        public string usageLocation { get; set; }

        [Attributes.Property("UserPrincipalName", SoType.Text, "User Principal Name", "User Principal Name")]
        public string userPrincipalName { get; set; }

        [Attributes.Property("UserType", SoType.Text, "User Type", "User Type")]
        public string userType { get; set; }

        [Attributes.Property("AboutMe", SoType.Text, "About Me", "About Me")]
        public string aboutMe { get; set; }

        [Attributes.Property("Birthday", SoType.DateTime, "Birthday", "Birthday")]
        public DateTime birthday { get; set; }

        [Attributes.Property("HireDate", SoType.DateTime, "Hire Date", "Hire Date")]
        public DateTime HireDate { get; set; }

        public string[] interests { get; set; }

        [Attributes.Property("MySite", SoType.Text, "My Site", "My Site")]
        public string mySite { get; set; }

        public string[] pastProjects { get; set; }

        [Attributes.Property("PreferredName", SoType.Text, "Preferred Name", "Preferred Name")]
        public string preferredName { get; set; }

        public string[] responsibilities { get; set; }
        public string[] schools { get; set; }
        public string[] skills { get; set; }


        [Attributes.Property("HttpResponseCode", SoType.Text, "HttpResponseCode", "HttpResponseCode")]
        public string HttpResponseCode { get; set; }


        [Attributes.Method("Me", SourceCode.SmartObjects.Services.ServiceSDK.Types.MethodType.Read, "Me", "Me",
        new string[] { "" }, //required property array (no required properties for this sample)
        new string[] { "", "" }, //input property array (no optional input properties for this sample)
        new string[] { "Id", "AccountEnabled", "City", "CompanyName", "Country", "Department", "DisplayName", "GivenName", "JobTitle", 
            "Mail", "MailNickname", "MobilePhone", "OfficeLocation", "PostalCode", "PreferredLanguage", "State", 
            "StreetAddress", "Surname", "UsageLocation", "UserPrincipalName", "UserType", "AboutMe", "Birthday", "HireDate", "MySite", "PreferredName"})]
        public GraphUser ExecuteMe()
        {
            GraphUser graphUser = new GraphUser();

            HttpWebRequest request = GetHttpWebRequest("https://graph.microsoft.com/beta/me", "GET");
            string result = string.Empty;

            using (HttpWebResponse Response = (HttpWebResponse)request.GetResponse())
            {
                using (Stream st = Response.GetResponseStream())
                {
                    using (StreamReader sr = new StreamReader(st))
                    {
                        result = sr.ReadToEnd();
                    }

                    graphUser = Newtonsoft.Json.JsonConvert.DeserializeObject<GraphUser>(result);
                }
            }


            return graphUser;
        }


        // Utilities
        private HttpWebRequest GetHttpWebRequest(string url, string method)
        {
            HttpWebRequest request = null;

            request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = method;

            if (ServiceConfiguration.ServiceAuthentication.AuthenticationMode == AuthenticationMode.Impersonate || ServiceConfiguration.ServiceAuthentication.AuthenticationMode == AuthenticationMode.ServiceAccount)
            {
                request.UseDefaultCredentials = true;
            }
            if (ServiceConfiguration.ServiceAuthentication.AuthenticationMode == AuthenticationMode.OAuth)
            {
                string accessToken = ServiceConfiguration.ServiceAuthentication.OAuthToken;
                string headerBearer = String.Format(CultureInfo.InvariantCulture, "Bearer {0}", accessToken);

                request.Headers.Add("Authorization", headerBearer.ToString());
            }
            if (ServiceConfiguration.ServiceAuthentication.AuthenticationMode == AuthenticationMode.Static)
            {
                char[] sp = { '\\' };
                string[] user = ServiceConfiguration.ServiceAuthentication.UserName.Split(sp);
                if (user.Length > 1)
                {
                    request.Credentials = new NetworkCredential(user[1], ServiceConfiguration.ServiceAuthentication.Password, user[0]);
                }
                else
                {
                    request.Credentials = new NetworkCredential(ServiceConfiguration.ServiceAuthentication.UserName, ServiceConfiguration.ServiceAuthentication.Password);
                }

            }

            //request.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;

            return request;
        }


    }

    public class Assignedlicense
    {
        public object[] disabledPlans { get; set; }
        public string skuId { get; set; }
    }

    public class Assignedplan
    {
        public DateTime assignedDateTime { get; set; }
        public string capabilityStatus { get; set; }
        public string service { get; set; }
        public string servicePlanId { get; set; }
    }

    public class Provisionedplan
    {
        public string capabilityStatus { get; set; }
        public string provisioningStatus { get; set; }
        public string service { get; set; }
    }


}
