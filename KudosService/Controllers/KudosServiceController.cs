using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Data.SqlClient;
using System.Data;
using System.Threading.Tasks;
using Microsoft.Graph;
using System.IO;
using System.Net.Http.Headers;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.Globalization;
using System.Web;

namespace KudosService.Controllers
{
    public class SendKudosJsonReceiver
    {
        public string KudosSender { get; set; }
        public string KudosSenderName { get; set; }
        public string KudosReceiver { get; set; }
        public string KudosReceiverName { get; set; }
        public string ItemID { get; set; }
        public string Subject { get; set; }
        public string AdditionalMessage { get; set; }
        public string SenderEmailAddress { get; set; }
    }

    public class QueryThreadResult
    {
        public string[] senders;
        public string[] senderNames;
        public string[] sentTime;
        public string[] thumbNails;
    }

    public class KudosInfo : IComparable
    {
        public string sender;
        public string senderName;
        public string itemID;
        public string subject;
        public DateTime sentTime;
        public string sentDate;
        public string additionalMessage;

        public KudosInfo()
        {
            sender = "";
            itemID = "";
            sentDate = "";
            sentTime = new DateTime();
            additionalMessage = "";
        }

        public int CompareTo(Object obj)
        {
            KudosInfo t = obj as KudosInfo;
            return sentTime.CompareTo(t.sentTime);
        }
    }

    public class QueryReceiverResult
    {
        public KudosInfo[] kudosInfos;
        public string[] months;
        public int[] kudosPerMonth;
        public int totalKudos;
    }

    class FileTokenCache : TokenCache
    {
        public string CacheFilePath;
        private static readonly object FileLock = new object();

        // Initializes the cache against a local file.
        // If the file is already present, it loads its content in the ADAL cache

        //public FileTokenCache(string filePath = @".\TokenCache.dat")
        public FileTokenCache(string filePath = @"C:/Projects/TokenCache.dat")
        {
            string s = System.Environment.CurrentDirectory;
            CacheFilePath = filePath;
            this.AfterAccess = AfterAccessNotification;
            this.BeforeAccess = BeforeAccessNotification;
            lock (FileLock)
            {
                this.Deserialize(System.IO.File.Exists(CacheFilePath) ? System.IO.File.ReadAllBytes(CacheFilePath) : null);
            }
        }

        // Empties the persistent store.
        public override void Clear()
        {
            base.Clear();
            System.IO.File.Delete(CacheFilePath);
        }

        // Triggered right before ADAL needs to access the cache.
        // Reload the cache from the persistent store in case it changed since the last access.
        void BeforeAccessNotification(TokenCacheNotificationArgs args)
        {
            lock (FileLock)
            {
                //this.Deserialize(File.Exists(CacheFilePath) ? ProtectedData.Unprotect(File.ReadAllBytes(CacheFilePath), null, DataProtectionScope.CurrentUser) : null);
                this.Deserialize(System.IO.File.Exists(CacheFilePath) ? System.IO.File.ReadAllBytes(CacheFilePath) : null);
            }
        }

        // Triggered right after ADAL accessed the cache.
        void AfterAccessNotification(TokenCacheNotificationArgs args)
        {
            // if the access operation resulted in a cache update
            if (this.HasStateChanged)
            {
                lock (FileLock)
                {
                    // reflect changes in the persistent store
                    System.IO.File.WriteAllBytes(CacheFilePath, this.Serialize());
                    // once the write operation took place, restore the HasStateChanged bit to false
                    this.HasStateChanged = false;
                }
            }
        }
    }

    class ThumbnailFetcher
    {
        private const string aadInstance = "https://login.microsoftonline.com/{0}";
        private const string tenant = "microsoft.com";
        private static string authority = String.Format(CultureInfo.InvariantCulture, aadInstance, tenant);

        private const string clientId = "c78c620d-e479-464a-9fe3-8cdac0823ae6";
        private const string resourceId = "https://graph.microsoft.com";

        private static readonly Uri returnUrl = new Uri("urn:ietf:wg:oauth:2.0:oob");

        private static async Task<string> GetAppTokenAsync()
        {
            var authContext = new AuthenticationContext(authority, new FileTokenCache());

            AuthenticationResult result = null;
            // first, try to get a token silently
            try
            {
                result = await authContext.AcquireTokenSilentAsync(resourceId, clientId);
            }
            catch (AdalException ex)
            {
                // There is no token in the cache; prompt the user to sign-in.
                if (ex.ErrorCode == "failed_to_acquire_token_silently")
                {
                }
                else
                {
                    throw;
                }
            }

            if (result == null)
            {
                //UserCredential uc = new UserCredential(user);
                // if you want to use Windows integrated auth, comment the line above and uncomment the one below
                // UserCredential uc = new UserCredential();
                try
                {
                    //result = await authContext.AcquireTokenAsync(resourceId, clientId, uc);
                    result = authContext.AcquireToken(resourceId, clientId, returnUrl, PromptBehavior.Always);
                }
                catch (Exception ee)
                {
                    throw;
                }
            }
            return result.AccessToken;
        }

        public static async Task<string> FetchAsync(string query)
        {
            var accessToken = await GetAppTokenAsync();

            var graphserviceClient = new GraphServiceClient(
                new DelegateAuthenticationProvider((requestMessage) =>
                {
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", accessToken);
                    return Task.FromResult(0);
                })
            );

            string result;
            var requestUrl = graphserviceClient.Users[query].Photo.AppendSegmentToRequestUrl("$value");
            //var requestUrl = graphserviceClient.Users["junxw@microsoft.com"].Photo.AppendSegmentToRequestUrl("$value");
            var requestBuilder = new ProfilePhotoContentRequestBuilder(requestUrl, graphserviceClient);

            try
            {
                var response = await requestBuilder.Request().GetAsync();
                using (var stream = new BinaryReader(response))
                {
                    var bytes = stream.ReadBytes((int)response.Length);
                    var base64 = Convert.ToBase64String(bytes);
                    result = base64;
                }
            }
            catch (Exception)
            {
                result = "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";
                //throw;
            }
            return result;
        }
    }

    public class KudosServiceController : ApiController
    {
        // GET: api/KudosService
        //public IEnumerable<string> Get()
        //{
        //    return new string[] { "value1", "value2" };
        //}

        //GET: api/KudosService

        public static string[] monthsRef = { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };

        public QueryThreadResult Get(string ItemID)
        {
            String connectionString = "Data Source=tcp:q4j05d8bmm.database.windows.net;Initial Catalog=Kudos;User ID=kudoweb;Password=User@123";
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            String commandText = "SELECT Sender, SenderName, SentTime, Thumbnail FROM KudosTable WHERE ItemID = '" + ItemID + "'";
            SqlCommand selectCommand = new SqlCommand(commandText, connection);
            SqlDataAdapter selectAdapter = new SqlDataAdapter();
            selectAdapter.SelectCommand = selectCommand;
            DataSet dataSet = new DataSet();
            selectAdapter.SelectCommand.ExecuteNonQuery();
            selectAdapter.Fill(dataSet);
            connection.Close();

            int totalSenders = dataSet.Tables[0].Rows.Count;
            QueryThreadResult result = new QueryThreadResult();
            result.senders = new string[totalSenders];
            result.senderNames = new string[totalSenders];
            result.sentTime = new string[totalSenders];
            result.thumbNails = new string[totalSenders];
            for (int i = 0; i < totalSenders; ++i)
            {
                result.senders[i] = (string)dataSet.Tables[0].Rows[i].ItemArray[0];
                result.senderNames[i] = (string)dataSet.Tables[0].Rows[i].ItemArray[1];
                result.sentTime[i] = ((DateTime)dataSet.Tables[0].Rows[i].ItemArray[2]).ToString("yyyy-MM-dd HH:mm:ss");
                result.thumbNails[i] = (string)dataSet.Tables[0].Rows[i].ItemArray[3];
            }
            return result;
        }

        //GET: api/KudosService/id
        public QueryReceiverResult Get(int id, string KudosReceiver)
        {
            String connectionString = "Data Source=tcp:q4j05d8bmm.database.windows.net;Initial Catalog=Kudos;User ID=kudoweb;Password=User@123";
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            String commandText = "SELECT Sender, SenderName, Subject, ItemID, SentTime, AdditionalMessage FROM KudosTable WHERE Receiver = '" + KudosReceiver + "'";
            SqlCommand selectCommand = new SqlCommand(commandText, connection);
            SqlDataAdapter selectAdapter = new SqlDataAdapter();
            selectAdapter.SelectCommand = selectCommand;
            DataSet dataSet = new DataSet();
            selectAdapter.SelectCommand.ExecuteNonQuery();
            selectAdapter.Fill(dataSet);

            String countCommandText = "SELECT COUNT(ID) FROM KudosTable";
            SqlCommand countCommand = new SqlCommand(countCommandText, connection);
            int count = (int)countCommand.ExecuteScalar();
            connection.Close();

            int totalSenders = dataSet.Tables[0].Rows.Count;
            QueryReceiverResult result = new QueryReceiverResult();
            result.totalKudos = count;
            result.kudosInfos = new KudosInfo[totalSenders];
            for (int i = 0; i < totalSenders; ++i)
            {
                result.kudosInfos[i] = new KudosInfo();
                result.kudosInfos[i].sender = (string)dataSet.Tables[0].Rows[i].ItemArray[0];
                result.kudosInfos[i].senderName = (string)dataSet.Tables[0].Rows[i].ItemArray[1];
                result.kudosInfos[i].subject = (string)dataSet.Tables[0].Rows[i].ItemArray[2];
                result.kudosInfos[i].itemID = (string)dataSet.Tables[0].Rows[i].ItemArray[3];
                result.kudosInfos[i].sentTime = (DateTime)dataSet.Tables[0].Rows[i].ItemArray[4];
                result.kudosInfos[i].sentDate = result.kudosInfos[i].sentTime.ToShortDateString();
                result.kudosInfos[i].additionalMessage = (string)dataSet.Tables[0].Rows[i].ItemArray[5];
            }
            Array.Sort(result.kudosInfos);
            result.months = GenerateFiveMonths(DateTime.Now);
            result.kudosPerMonth = CountKudosBefore(DateTime.Now, result.kudosInfos);
            return result;
        }

        int[] CountKudosBefore(DateTime currentTime, KudosInfo[] kudosInfo)
        {
            int[] result = new int[5];
            DateTime startTime = currentTime;
            startTime.AddHours(-startTime.Hour);
            startTime.AddMinutes(-startTime.Minute);
            startTime.AddSeconds(-startTime.Second);
            startTime.AddMilliseconds(-startTime.Millisecond);

            startTime = startTime.AddDays(1 - startTime.Day);
            startTime = startTime.AddMonths(-4);
            foreach (KudosInfo info in kudosInfo)
            {
                DateTime time = info.sentTime;
                if ((time.CompareTo(startTime) >= 0) && (time.CompareTo(currentTime) <= 0))
                {
                    int t = time.Month - startTime.Month;
                    if (t < 0)
                    {
                        t += 12;
                    }
                    ++result[t];
                }
            }
            return result;
        }

        string[] GenerateFiveMonths(DateTime currentTime)
        {
            string[] result = new string[5];
            int month = currentTime.Month;
            for (int i = 0; i < 5; ++i)
            {
                result[4 - i] = monthsRef[month - 1];
                --month;
                if (month == 0)
                {
                    month = 12;
                }
            }
            return result;
        }

        // POST: api/KudosService
        public async Task<int> Post([FromBody]SendKudosJsonReceiver value)
        {
            String connectionString = "Data Source=tcp:q4j05d8bmm.database.windows.net;Initial Catalog=Kudos;User ID=kudoweb;Password=User@123";
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            string currentTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            string base64 = await ThumbnailFetcher.FetchAsync(value.SenderEmailAddress);
            String commandText =
                "INSERT INTO KudosTable (Sender, SenderName, Receiver, ReceiverName, Subject, ItemID, AdditionalMessage, SentTime, Thumbnail) VALUES ('"
                + value.KudosSender
                + "', '" + value.KudosSenderName
                + "', '" + value.KudosReceiver
                + "', '" + value.KudosReceiverName
                + "', '" + value.Subject
                + "', '" + value.ItemID
                + "', N'" + value.AdditionalMessage
                + "', '" + currentTime
                + "', '" + base64
                + "')";
            SqlCommand insertCommand = new SqlCommand(commandText, connection);
            insertCommand.ExecuteNonQuery();
            connection.Close();
            return 3154;
        }

        // PUT: api/KudosService/5
        public void Put(int id, [FromBody]string value)
        {
        }

        // DELETE: api/KudosService/5
        public void Delete(int id)
        {
        }
    }
}
