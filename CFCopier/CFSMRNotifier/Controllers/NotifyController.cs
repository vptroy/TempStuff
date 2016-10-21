using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using Microsoft.Bot.Connector;

namespace CFSMRNotifier.Controllers
{
    public class NotifyController : ApiController
    {
        [HttpGet]
        public string Notify(string from, string recipient, string conversation, string message)
        {
            ConnectorClient connector = new ConnectorClient(new Uri("https://skype.botframework.com"));
            IMessageActivity newMessage = Activity.CreateMessageActivity();
            newMessage.Type = ActivityTypes.Message;
            //newMessage.From = botAccount;
            //newMessage.Conversation = conversation;
            //newMessage.Recipient = userAccount;
            newMessage.From = new ChannelAccount() {Id = from };
            newMessage.Recipient = new ChannelAccount() {Id = recipient };
            newMessage.Conversation = new ConversationAccount() {Id = conversation };
            newMessage.Text = message;

            connector.Conversations.SendToConversationAsync((Activity)newMessage);
            //connector.Conversations.SendToConversation(activity.CreateReply("asdsds"), "8a684db8");
            


            return "value1";
        }

        //[HttpGet]
        //public string NotifyTest()
        //{
        //    ConnectorClient connector = new ConnectorClient(new Uri("http://localhost:9000/"));
        //    IMessageActivity newMessage = Activity.CreateMessageActivity();
        //    newMessage.Type = ActivityTypes.Message;
        //    //newMessage.From = botAccount;
        //    //newMessage.Conversation = conversation;
        //    //newMessage.Recipient = userAccount;
        //    newMessage.From = new ChannelAccount() { Id = "56800324" };
        //    newMessage.Recipient = new ChannelAccount() { Id = "2c1c7fa3" };
        //    newMessage.Conversation = new ConversationAccount() { Id = "8a684db8" };
        //    newMessage.Text = "Yo yo yo!";

        //    connector.Conversations.SendToConversationAsync((Activity)newMessage);
        //    //connector.Conversations.SendToConversation(activity.CreateReply("asdsds"), "8a684db8");



        //    return "value2";
        //}

        //// POST: api/Default
        //public void Post(string from, string recipient, string conversation)
        //{
            
        //}
    }
}
