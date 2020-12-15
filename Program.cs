using Ice.Core;
using Ice.Lib.Framework;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;

using Ice.BO;
using System.Net.Http;
using System.Net;
using System.Net.Http.Headers;
using Newtonsoft.Json;

namespace SConnectTemplate
{
    class Program
    {


        static UtilityProcesses oUtility = new UtilityProcesses();
        static string sPathToWatch = "";

        static void Main(string[] args)
        {
            //  *** Start How do you want to start the process? ****//
            //  1. WatchFilePath() - Run process when a file is dropped in a specified file path.
            //  2. ResetTimer() - Run process at a specified time of day.
            //  3. RunAtInterval() - Run at a specified interval.
            //  Uncomment the process yuo want to use.

            //oUtility.WatchFilePath();
            //oUtility.ResetTimer();
            //oUtility.RunAtInterval();
        }


    }

    class RunProcess
    {
        SalesforceClient oSFClient = new SalesforceClient();

        UtilityProcesses oUtility = new UtilityProcesses();

        internal void ProcessData(object stateData)
        {
            DateTime dtNow = DateTime.Now;

            //  Run by Day of Week
            string sDaysOfWeekToRun = ConfigurationManager.AppSettings["DaysOfWeekToRun"];
            if (sDaysOfWeekToRun != "")
            {
                string[] sDaysOfWeekToRunArray = sDaysOfWeekToRun.Split(',');
                for (int x = 0; x < sDaysOfWeekToRunArray.Length; x++)
                {
                    if (dtNow.DayOfWeek.ToString() == sDaysOfWeekToRunArray[x])
                    {
                        ExecuteProcess(stateData);
                    }
                }
            }
            else
            {
                ExecuteProcess(stateData);
            }

            //  Run by Day in Month
            string sDaysOfMonthToRun = ConfigurationManager.AppSettings["DaysOfMonthToRun"];
            if (sDaysOfMonthToRun != "")
            {
                string[] sDaysOfMonthToRunArray = sDaysOfMonthToRun.Split(',');
                for (int x = 0; x < sDaysOfMonthToRunArray.Length; x++)
                {
                    if (dtNow.Day.ToString() == sDaysOfMonthToRunArray[x])
                    {
                        ExecuteProcess(stateData);
                    }
                }
            }
            else
            {
                ExecuteProcess(stateData);
            }
        }

        private void ExecuteProcess(object stateData)
        {
            string sConfigFilePath = "";
            string sCSVFilePath = "";

            string sUserID = "";
            string sPassword = "";

            Session epiSession = null;

            try
            {

                oUtility.LogData("Start Connection to Epicor.");

                //  Use ConfigurationManager object to get values from App.config file.
                sConfigFilePath = ConfigurationManager.AppSettings["configFilePath"];

                sCSVFilePath = ConfigurationManager.AppSettings["CSVFilePath"];
                sCSVFilePath = stateData.ToString();

                sUserID = ConfigurationManager.AppSettings["UserID"];
                sPassword = ConfigurationManager.AppSettings["Password"];

                //  Connect to Epicor
                epiSession = new Session(sUserID, sPassword, Session.LicenseType.Default, sConfigFilePath);
                oUtility.LogData("Connection to Epicor successful.");

                var client = oSFClient.CreateClient();

                //  Get UD25 Business Object
                Ice.Proxy.BO.UD25Impl oUD25 = WCFServiceSupport.CreateImpl<Ice.Proxy.BO.UD25Impl>((Ice.Core.Session)epiSession, Epicor.ServiceModel.Channels.ImplBase<Ice.Contracts.UD25SvcContract>.UriPath);

                //  ***************************************************************
                //  *** Start Get data from a CSV and populate an Epicor table. ***
                //  ***************************************************************
                //  Open and read a CSV file (sample).
                string sCSVFileData = oUtility.OpenAndReadCSVFile(sCSVFilePath);

                string[] sFileDataArray = sCSVFileData.Split('\n');     //  Split the CSV file data rows using the \n new line character.

                //  Check to set if there are any records in the CSV and loop thru the rows.
                if (sFileDataArray.Length > 0)
                {
                    //  Loop thru each row of the CSV file
                    for (int x = 0; x < sFileDataArray.Length - 1; x++)
                    {
                        if (x > 0)  // x > 0 to skip 1st row.
                        {
                            //  Initialize variables to hold field data.
                            string sColumn1 = "";
                            string sColumn2 = "";
                            string sColumn3 = "";
                            //  Get a single row from the CSV 
                            string sFileDataLine = sFileDataArray[x].Replace("\r", ""); ;
                            //  Split the row on the "," comma to get the field values.
                            string[] sFileDataLineArray = sFileDataLine.Split(',');

                            //  Get columns data from CSV file.
                            sColumn1 = sFileDataLineArray[0];
                            sColumn2 = sFileDataLineArray[1];
                            sColumn3 = sFileDataLineArray[2];

                            //  Create Epicor UD25 dataset to store data.
                            UD25DataSet dsUD25 = new UD25DataSet();

                            //  Create a new UD25 record.
                            oUD25.GetaNewUD25(dsUD25);                                  //

                            //  Populate UD25 columns
                            dsUD25.Tables[0].Rows[0]["Key1"] = sColumn1;                //  UD25 Key1 field
                            dsUD25.Tables[0].Rows[0]["Key2"] = sColumn2;                //  UD25 Key2 field
                            Random r = new Random();                                    //  Random number generator
                            dsUD25.Tables[0].Rows[0]["Number01"] = r.NextDouble();      //  Number01 field - Populate field with random number
                            dsUD25.Tables[0].Rows[0]["ShortChar01"] = sColumn3;         //  UD25DataSet Shortchar01 field
                            dsUD25.Tables[0].Rows[0]["Date01"] = DateTime.Now;          //  Date01 field - Populate with current date

                            oUD25.Update(dsUD25);                                       //  Update UD25 with record data.

                            dsUD25.Dispose();
                        }
                    }
                }

                oUD25.Dispose();
                //  ***************************************************************
                //  *** End Get data from a CSV and populate an Epicor table. ***
                //  ***************************************************************


            }
            catch (Exception ex)
            {

            }

            //  Send an email.
            string sEmailFrom = ConfigurationManager.AppSettings["EmailFrom"];
            string sEmailTo = "scoach@dukemfg.com";
            string sSubject = "Program execution complete.";
            string sBody = "Program execution complete.";
            oUtility.SendEmail(sEmailTo, sEmailFrom, sSubject, sBody);

            oUtility.ResetTimer();

        }
    }

    class UtilityProcesses
    {
        SalesforceClient oSFClient = new SalesforceClient();

        static FileSystemWatcher myWatcher = null;
        System.Threading.Timer oTimer;

        //  *** Start How do you want to start the process? ****//
        //  1. WatchFilePath() - Run process when a file is dropped in a specified file path.
        //  2. ResetTimer() - Run process at a specified time of day.
        //  3. RunAtInterval() - Run at a specified interval.
        public void WatchFilePath()
        {
            string sPathToWatch = ConfigurationManager.AppSettings["PathToWatch"];
            myWatcher = new FileSystemWatcher();
            myWatcher.Path = sPathToWatch;
            myWatcher.Filter = "*.CSV";
            myWatcher.EnableRaisingEvents = true;
            myWatcher.IncludeSubdirectories = true;

            myWatcher.Created += new FileSystemEventHandler(RunForWatchFilePath);

            myWatcher.EnableRaisingEvents = true;
        }

        private static void RunForWatchFilePath(object source, FileSystemEventArgs e)
        {
            RunProcess oProcess = new RunProcess();
            oProcess.ProcessData(e.FullPath);
        }
        public void ResetTimer()
        {
            //  Main functiion is used to start timer.
            double dblTimeToRun = 1.0;

            //  Get the time to run from the app.config file.
            string sTimeToRun = ConfigurationManager.AppSettings["TimeToRun24Decimal"];
            if (sTimeToRun != "")
            {
                dblTimeToRun = Convert.ToDouble(sTimeToRun);
            }

            //  Create the .NET object used to run our process object.
            RunProcess oProcess = new RunProcess();

            //  Create the timer object used to run our process object at the specificied time.
            var t = new System.Threading.Timer(new System.Threading.TimerCallback(oProcess.ProcessData));

            // Figure how much time until 1:00 AM
            DateTime now = DateTime.Now;
            DateTime oneAM = DateTime.Today.AddHours(dblTimeToRun);

            // If it's already past 4:00, wait until 4:00 tomorrow 
            if (now > oneAM)
            {
                oneAM = oneAM.AddDays(1.0);
            }

            int msUntilOneAM = (int)((oneAM - now).TotalMilliseconds);

            // Set the timer to elapse only once, at 1:00. 
            t.Change(msUntilOneAM, System.Threading.Timeout.Infinite);

            LogData("XXXX - The Name or Description of the Program - XXXX : set to run at : " + oneAM.ToString());
            Console.WriteLine("XXXX - The Name or Description of the Program - XXXX : set to run at : " + oneAM.ToString());
            Console.ReadLine();
        }

        public void RunAtInterval()
        {
            System.Timers.Timer oTimer = new System.Timers.Timer();

            string sTimerInterval = ConfigurationManager.AppSettings["TimerInterval"];
            string sTimerIntervalUnits = ConfigurationManager.AppSettings["TimerIntervalUnits"];

            int iTimerIntervalUnits = 0;
            if (sTimerIntervalUnits == "Seconds")
            {
                iTimerIntervalUnits = 1000;
            }
            if (sTimerIntervalUnits == "Minutes")
            {
                iTimerIntervalUnits = 60000;
            }
            if (sTimerIntervalUnits == "Hours")
            {
                iTimerIntervalUnits = 360000000;
            }

            int iTimerInterval = Convert.ToInt32(sTimerInterval);
            oTimer.Interval = iTimerInterval * iTimerIntervalUnits;

            oTimer.Enabled = true;
            oTimer.Elapsed += OTimer_Elapsed;
            oTimer.Start();
        }

        private void OTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            RunProcess oProcess = new RunProcess();
            oProcess.ProcessData("");
        }

        //  *** End How do you want to start the process? ****//


        public string OpenAndReadCSVFile(string sCSVFilePath)
        {
            string sFileData = "";
            System.Threading.Thread.Sleep(1000);
            try
            {
                System.IO.StreamReader oReader = new System.IO.StreamReader(sCSVFilePath);
                sFileData = oReader.ReadToEnd();
                oReader.Close();

                LogData("File '" + sCSVFilePath + "' opened and read.");
            }
            catch (Exception ex)
            {

            }

            return sFileData;
        }

        public void LogData(string sDataToLog)
        {
            System.IO.StreamWriter oWriter = new System.IO.StreamWriter("LogFile.txt", true);
            oWriter.WriteLine(DateTime.Now.ToString() + ", " + sDataToLog);
            oWriter.Close();
        }

        public void SendEmail(string sTo, string sFrom, string sSubject, string sBody)
        {
            string sSMTPAddress = ConfigurationManager.AppSettings["SMTPAddress"];

            try
            {
                System.Net.Mail.MailMessage oMail = new System.Net.Mail.MailMessage(sFrom, sTo);

                oMail.Subject = sSubject;

                oMail.Body = sBody;

                oMail.IsBodyHtml = true;

                System.Net.Mail.SmtpClient oSmtp = new System.Net.Mail.SmtpClient(sSMTPAddress);
                oSmtp.Port = 25;

                oSmtp.Send(oMail);

                LogData("Email sent succefully to : " + sTo);
            }
            catch (Exception ex)
            {
                LogData("Error sending email : " + ex.Message);
            }

        }

        public void Login()
        {
            String jsonResponse;
            using (var client = new HttpClient())
            {
                    var request = new FormUrlEncodedContent(new Dictionary<string, string>
                    {
                        {"grant_type", "password"},
                        {"client_id", oSFClient.ClientId},
                        {"client_secret", oSFClient.ClientSecret},
                        {"username", oSFClient.Username},
                        {"password", oSFClient.Password + oSFClient.Token}
                    });
                    request.Headers.Add("X-PrettyPrint", "1");
                    var response = client.PostAsync(oSFClient.LOGIN_ENDPOINT, request).Result;
                    jsonResponse = response.Content.ReadAsStringAsync().Result;
            }
            var values = JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonResponse);
            oSFClient.AuthToken = values["access_token"];
            oSFClient.InstanceUrl = values["instance_url"];
        }


        public string Query(string soqlQuery)
        {
            using (var client = new HttpClient())
            {
                string restRequest = oSFClient.InstanceUrl + oSFClient.API_ENDPOINT + "query/?q=" + soqlQuery;
                var request = new HttpRequestMessage(HttpMethod.Get, restRequest);
                request.Headers.Add("Authorization", "Bearer " + oSFClient.AuthToken);
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                request.Headers.Add("X-PrettyPrint", "1");
                var response = client.SendAsync(request).Result;
                return response.Content.ReadAsStringAsync().Result;
            }
        }

        private string CreateRecord(HttpClient client, string createMessage, string recordType)
        {
            HttpContent contentCreate = new StringContent(createMessage, Encoding.UTF8, "application/xml");
            string uri = oSFClient.InstanceUrl + oSFClient.API_ENDPOINT + "/sobjects/" + recordType;

            HttpRequestMessage requestCreate = new HttpRequestMessage(HttpMethod.Post, uri);
            requestCreate.Headers.Add("Authorization", "Bearer " + oSFClient.AuthToken);
            requestCreate.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/xml"));
            requestCreate.Content = contentCreate;

            HttpResponseMessage response = client.SendAsync(requestCreate).Result;
            return response.Content.ReadAsStringAsync().Result;
        }



    }



    public class SalesforceClient
    {

        static SalesforceClient()
        {
            // SF requires TLS 1.1 or 1.2
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11;
        }

        public string LOGIN_ENDPOINT = "https://login.salesforce.com/services/oauth2/token";
        public string API_ENDPOINT = "/services/data/v36.0/";

        public string Username { get; set; }
        public string Password { get; set; }
        public string Token { get; set; }
        public string ClientId { get; set; }
        public string ClientSecret { get; set; }
        public string AuthToken { get; set; }
        public string InstanceUrl { get; set; }

        public SalesforceClient CreateClient()
        {
            return new SalesforceClient
            {
                Username = ConfigurationManager.AppSettings["username"],
                Password = ConfigurationManager.AppSettings["password"],
                Token = ConfigurationManager.AppSettings["token"],
                ClientId = ConfigurationManager.AppSettings["clientId"],
                ClientSecret = ConfigurationManager.AppSettings["clientSecret"]
            };
        }
    }
}

