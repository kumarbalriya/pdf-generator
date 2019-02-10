using Microsoft.Office.Interop.Outlook;

namespace LetterRollout
{
    public class SendMail
    {
        Application application;
        public SendMail()
        {
            application = new Application();
        }

        public void Main(Model model, string pdfFilePath)
        {
            string body = "Dear " + model.name + ", Test Mail";
            string subject = "Test Mail";
            //string attachmentFileName = "";

            MailItem item = application.CreateItem(OlItemType.olMailItem);
            item.HTMLBody = body;
            item.Subject = subject;
            int position = item.Body.Length + 1;
            item.Attachments.Add(pdfFilePath, 1, position, model.firstName + ".pdf");
            foreach (Account account in application.Session.Accounts)
            {
                item.SendUsingAccount = account;
            }

            // TO)
            Recipient recipientTo = item.Recipients.Add(model.emailId);
            recipientTo.Type = (int)OlMailRecipientType.olTo;

            // CC
            Recipient recipientCC = item.Recipients.Add(model.cc);
            recipientCC.Type = (int)OlMailRecipientType.olCC;

            item.Recipients.ResolveAll();

            //item.Send();
            item = null;
            //application = null;
        }
    }
}
