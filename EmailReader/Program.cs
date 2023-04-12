using MailKit.Net.Pop3;
using MailKit.Security;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace EmailReader
{
    internal class Program
    {

        static void Main(string[] args)
        {
            Console.WriteLine(EmailActivationCodeReader());
            Console.ReadLine();
        }

        public static string EmailActivationCodeReader()
        {
            using (var client = new Pop3Client())
            {
                client.Connect("pop.gmail.com", 995, SecureSocketOptions.SslOnConnect);

                #region your mail adress and your key
                client.Authenticate("enter your mail", "enter your key");
                #endregion

                #region get last massage
                var lastmassage = client.GetMessage(client.Count - 1);
                #endregion

                string ActivationCode = lastmassage.HtmlBody;

                string theBody = "";

                RegexOptions options = RegexOptions.IgnoreCase | RegexOptions.Singleline;

                #region Which HTML tags do you want to read the value between?    
                Regex regx = new Regex("</strong>(?<theBody>.*)</p>", options);
                #endregion

                Match match = regx.Match(ActivationCode);

                if (match.Success)
                {
                    theBody = match.Groups["theBody"].Value;

                }
                client.Disconnect(true);
                return theBody;

            }
        }
    }
}
