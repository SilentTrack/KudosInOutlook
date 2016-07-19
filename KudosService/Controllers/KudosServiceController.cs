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

    public class QueryKudosJsonReceiver
    {
        public string InternetMessageID { get; set; }
    }


    public class KudosServiceController : ApiController
    {
        // GET: api/KudosService
        //public IEnumerable<string> Get()
        //{
        //    return new string[] { "value1", "value2" };
        //}

        //GET: api/KudosService/5
        public string[] Get(string InternetMessageID)
        {
            String connectionString = "Data Source=tcp:q4j05d8bmm.database.windows.net;Initial Catalog=Kudos;User ID=kudoweb;Password=User@123";
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            String commandText = "SELECT Sender FROM KudosTable WHERE InternetMessageID = '" + InternetMessageID + "'";
            //String commandText = "SELECT Receiver FROM KudosTable WHERE Sender = 'Zhaohua Feng'";
            SqlCommand selectCommand = new SqlCommand(commandText, connection);
            SqlDataAdapter selectAdapter = new SqlDataAdapter();
            selectAdapter.SelectCommand = selectCommand;
            DataSet dataSet = new DataSet();
            selectAdapter.SelectCommand.ExecuteNonQuery();
            selectAdapter.Fill(dataSet);
            connection.Close();

            int totalSenders = dataSet.Tables[0].Rows.Count;
            string[] senders = new string[totalSenders];
            for (int i = 0; i < totalSenders; ++i)
            {
                senders[i] = (string)dataSet.Tables[0].Rows[i].ItemArray[0];
            }
            return senders;
        }

        // POST: api/KudosService
        public int Post([FromBody]SendKudosJsonReceiver value)
        {
            String connectionString = "Data Source=tcp:q4j05d8bmm.database.windows.net;Initial Catalog=Kudos;User ID=kudoweb;Password=User@123";
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            String commandText = "INSERT INTO KudosTable (Sender, Receiver, InternetMessageID, AdditionalMessage) VALUES ('" + value.KudosSender + "', '" + value.KudosReceiver + "', '" + value.InternetMessageID + "', '" + value.AdditionalMessage + "')";
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
