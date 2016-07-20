using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Data.SqlClient;
using System.Data;

namespace KudosService.Controllers
{
    public class SendKudosJsonReceiver
    {
        public string KudosSender { get; set; }
        public string KudosReceiver { get; set; }
        public string InternetMessageID { get; set; }
        public string AdditionalMessage { get; set; }
    }

    public class QueryResult
    {
        public string[] senders;
        public string[] sentTime;
    }

    public class KudosServiceController : ApiController
    {
        // GET: api/KudosService
        //public IEnumerable<string> Get()
        //{
        //    return new string[] { "value1", "value2" };
        //}

        //GET: api/KudosService/5
        public QueryResult Get(string InternetMessageID)
        {
            String connectionString = "Data Source=tcp:q4j05d8bmm.database.windows.net;Initial Catalog=Kudos;User ID=kudoweb;Password=User@123";
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            String commandText = "SELECT Sender, SentTime FROM KudosTable WHERE InternetMessageID = '" + InternetMessageID + "'";
            //String commandText = "SELECT Receiver FROM KudosTable WHERE Sender = 'Zhaohua Feng'";
            SqlCommand selectCommand = new SqlCommand(commandText, connection);
            SqlDataAdapter selectAdapter = new SqlDataAdapter();
            selectAdapter.SelectCommand = selectCommand;
            DataSet dataSet = new DataSet();
            selectAdapter.SelectCommand.ExecuteNonQuery();
            selectAdapter.Fill(dataSet);
            connection.Close();

            int totalSenders = dataSet.Tables[0].Rows.Count;
            QueryResult result = new QueryResult();
            result.senders = new string[totalSenders];
            result.sentTime = new string[totalSenders];
            for (int i = 0; i < totalSenders; ++i)
            {
                result.senders[i] = (string)dataSet.Tables[0].Rows[i].ItemArray[0];
                result.sentTime[i] = ((DateTime)dataSet.Tables[0].Rows[i].ItemArray[1]).ToString("yyyy-MM-dd HH:mm:ss");
            }
            return result;
        }

        // POST: api/KudosService
        public int Post([FromBody]SendKudosJsonReceiver value)
        {
            String connectionString = "Data Source=tcp:q4j05d8bmm.database.windows.net;Initial Catalog=Kudos;User ID=kudoweb;Password=User@123";
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            string currentTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            String commandText = "INSERT INTO KudosTable (Sender, Receiver, InternetMessageID, AdditionalMessage, SentTime) VALUES ('" + value.KudosSender + "', '" + value.KudosReceiver +
                "', '" + value.InternetMessageID + "', '" + value.AdditionalMessage
                + "', '" + currentTime + "')";
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
