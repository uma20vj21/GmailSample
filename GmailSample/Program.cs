using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.PeopleService.v1;
using Google.Apis.PeopleService.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using MimeKit;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace GmailSample
{
    class Program
    {
        private static string? _homeDirectory;
        private static bool isReceivingEmails = false;
        private static Task? emailReceivingTask;

        [Obsolete]
        static void Main(string[] args)
        {
            InitializeGmailService(out var gmailService, out var peopleService);

            while (true)
            {
                Console.WriteLine("次のアクションを選択してください:");
                Console.WriteLine("1: メール送信");
                Console.WriteLine("2: メール受信開始");
                Console.WriteLine("3: メール一覧表示");
                Console.WriteLine("4: システム終了");

                string action = Console.ReadLine();

                switch (action)
                {
                    case "1":
                        SendEmail(gmailService, peopleService);
                        break;
                    case "2":
                        if (!isReceivingEmails)
                        {
                            Console.WriteLine("メールを受信状態にしますか？(y/n)");
                            if (Console.ReadLine()?.ToLower() == "y")
                            {
                                isReceivingEmails = true;
                                emailReceivingTask = Task.Run(() => ReceiveEmails(gmailService));
                            }
                        }
                        else
                        {
                            isReceivingEmails = false;
                            emailReceivingTask?.Wait();
                            Console.WriteLine("メール受信を終了しました。");
                        }
                        break;
                    case "3":
                        DisplayEmailList(gmailService);
                        break;
                    case "4":
                        if (isReceivingEmails)
                        {
                            isReceivingEmails = false;
                            emailReceivingTask?.Wait();
                        }
                        Console.WriteLine("システムを終了します...");
                        return;
                    default:
                        Console.WriteLine("無効な選択です。もう一度入力してください。");
                        break;
                }
            }
        }

        [Obsolete]
        private static void InitializeGmailService(out GmailService gmailService, out PeopleServiceService peopleService)
        {
            string[] scopes = {
                GmailService.Scope.MailGoogleCom,
                PeopleServiceService.Scope.UserinfoProfile,
                PeopleServiceService.Scope.UserinfoEmail,
                PeopleServiceService.Scope.ContactsReadonly
            };

            string appName = "Google.Apis.Gmail.v1 Sample";

            UserCredential credential;
            _homeDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string clientSecretPath = Path.Combine(_homeDirectory, "Downloads", "clientsecret.json");

            if (!File.Exists(clientSecretPath))
            {
                Console.WriteLine("クライアントシークレットファイルが見つかりません。");
                throw new FileNotFoundException("クライアントシークレットファイルが見つかりません。");
            }

            string credPath = Path.Combine(_homeDirectory, ".credentials/gmail-dotnet-quickstart.json");

            using (var stream = new FileStream(clientSecretPath, FileMode.Open, FileAccess.Read))
            {
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.Load(stream).Secrets,
                        scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
            }

            gmailService = new GmailService(new BaseClientService.Initializer()
            {
                ApplicationName = appName,
                HttpClientInitializer = credential
            });

            peopleService = new PeopleServiceService(new BaseClientService.Initializer()
            {
                ApplicationName = appName,
                HttpClientInitializer = credential
            });
        }

        private static void SendEmail(GmailService gmailService, PeopleServiceService peopleService)
        {
            try
            {
                var profileRequest = peopleService.People.Get("people/me");
                profileRequest.PersonFields = "names,emailAddresses";
                var profile = profileRequest.Execute();

                string mailFromAddress = profile.EmailAddresses[0].Value;
                string mailFromName = profile.Names[0].DisplayName;

                Console.WriteLine("送信先の名前を入力して下さい");
                string mailToName = Console.ReadLine();

                Console.WriteLine("送信先のメールアドレスを入力して下さい");
                string mailToAddress = Console.ReadLine();

                Console.WriteLine("メールの件名を入力してください");
                string mailSubject = Console.ReadLine();

                Console.WriteLine("メールの本文を入力してください: ");
                string mailBody = Console.ReadLine();

                Console.WriteLine("メールを送信しますか？(y/n)：");
                if (Console.ReadLine()?.ToLower() == "y")
                {
                    var mimeMessage = new MimeMessage();
                    mimeMessage.From.Add(new MailboxAddress(mailFromName, mailFromAddress));
                    mimeMessage.To.Add(new MailboxAddress(mailToName, mailToAddress));
                    mimeMessage.Subject = mailSubject;
                    mimeMessage.Body = new TextPart(MimeKit.Text.TextFormat.Plain) { Text = mailBody };

                    try
                    {
                        var rawMessage = EncodeMessage(mimeMessage);
                        var result = gmailService.Users.Messages.Send(new Message { Raw = rawMessage }, "me").Execute();

                        Console.WriteLine("送信完了しました。");
                        Console.WriteLine("========================");
                        Console.WriteLine("送信したメールの内容:");
                        Console.WriteLine($"From: {mailFromName} <{mailFromAddress}>");
                        Console.WriteLine($"To: {mailToName} <{mailToAddress}>");
                        Console.WriteLine($"Subject: {mailSubject}");
                        Console.WriteLine($"Body: {mailBody}");
                        Console.WriteLine("========================");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"メールの送信中にエラーが発生しました: {ex.Message}");
                    }
                }
            }
            catch (Google.GoogleApiException e)
            {
                Console.WriteLine($"プロフィール情報の取得中にエラーが発生しました: {e.Message}");
                Console.WriteLine("認証スコープが不足している可能性があります。スコープを確認してください。");
            }
        }

        private static async void ReceiveEmails(GmailService gmailService)
        {
            UsersResource.MessagesResource.ListRequest request = gmailService.Users.Messages.List("me");
            request.LabelIds = "INBOX";
            request.IncludeSpamTrash = false;
            string? lastMessageId = null;

            while (isReceivingEmails)
            {
                try
                {
                    ListMessagesResponse response = await request.ExecuteAsync();
                    if (response.Messages != null)
                    {
                        foreach (var message in response.Messages)
                        {
                            if (message.Id == lastMessageId)
                                break;


                            var emailInfoReq = gmailService.Users.Messages.Get("me", message.Id);
                            var emailInfoResponse = await emailInfoReq.ExecuteAsync();
                            if (emailInfoResponse != null)
                            {
                                string from = "";
                                string date = "";
                                string subject = "";
                                string body = "";

                                foreach (var mParts in emailInfoResponse.Payload.Headers)
                                {
                                    if (mParts.Name == "Date")
                                    {
                                        date = mParts.Value;
                                    }
                                    else if (mParts.Name == "From")
                                    {
                                        from = mParts.Value;
                                    }
                                    else if (mParts.Name == "Subject")
                                    {
                                        subject = mParts.Value;
                                    }
                                }

                                body = GetMessageBody(emailInfoResponse.Payload);

                                while (true)
                                {
                                    Console.WriteLine("新しいメールを受信しました。確認しますか？(y/n)");
                                    string? input = Console.ReadLine()?.ToLower();
                                    if (input == "y")
                                    {
                                        Console.WriteLine($"日付: {date}");
                                        Console.WriteLine($"差出人: {from}");
                                        Console.WriteLine($"件名: {subject}");
                                        Console.WriteLine("本文:");
                                        Console.WriteLine(body);
                                        Console.WriteLine("-----------------------------------------------------");
                                        break;
                                    }
                                    else if (input == "n")
                                    {
                                        break;
                                    }
                                    else
                                    {
                                        Console.WriteLine("無効な選択です。もう一度入力してください。");
                                    }
                                }

                            }
                            lastMessageId = message.Id;
                        }
                    }

                    await Task.Delay(10000);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"メールの受信中にエラーが発生しました: {ex.Message}");
                }
            }
        }

        private static string GetMessageBody(Google.Apis.Gmail.v1.Data.MessagePart payload)
        {
            string body = "";
            if (payload.Parts == null && payload.Body != null && payload.Body.Data != null)
            {
                body = Base64UrlDecode(payload.Body.Data);
            }
            else
            {
                foreach (var part in payload.Parts)
                {
                    if (part.MimeType == "text/plain")
                    {
                        body += Base64UrlDecode(part.Body.Data);
                        break;
                    }
                    else if (part.Parts != null)
                    {
                        body += GetMessageBody(part);
                    }
                }
            }
            return body;
        }

        private static string Base64UrlDecode(string input)
        {
            string s = input.Replace('-', '+').Replace('_', '/');
            switch (input.Length % 4)
            {
                case 0: break;
                case 2: s += "=="; break;
                case 3: s += "="; break;
                default: throw new System.Exception("Illegal base64url string!");
            }
            byte[] bytes = Convert.FromBase64String(s);
            return System.Text.Encoding.UTF8.GetString(bytes);
        }

        private static string EncodeMessage(MimeMessage mimeMessage)
        {
            using (var memoryStream = new MemoryStream())
            {
                mimeMessage.WriteTo(memoryStream);
                return Convert.ToBase64String(memoryStream.ToArray())
                    .Replace('+', '-')
                    .Replace('/', '_')
                    .Replace("=", "");
            }
        }

        private static void DisplayEmailList(GmailService gmailService)
        {
            Console.WriteLine("検索したい文言はありますか？");
            string searchTerm = Console.ReadLine();

            UsersResource.MessagesResource.ListRequest request = gmailService.Users.Messages.List("me");
            request.LabelIds = "INBOX";
            request.IncludeSpamTrash = false;
            if (!string.IsNullOrEmpty(searchTerm))
            {
                request.Q = searchTerm;
            }

            try
            {
                ListMessagesResponse response = request.Execute();
                if (response.Messages != null && response.Messages.Count > 0)
                {
                    int displayCount;
                    if (!string.IsNullOrEmpty(searchTerm))
                    {
                        Console.WriteLine($"{response.Messages.Count}件の一致するメールがあります。何件分表示しますか？");
                    }
                    else
                    {
                        Console.WriteLine("検索文言が入力されませんでした。何件分表示しますか？");
                    }

                    while (true)
                    {
                        if (int.TryParse(Console.ReadLine(), out displayCount) && displayCount > 0 && displayCount <= response.Messages.Count)
                        {
                            break;
                        }
                        Console.WriteLine("適切な値を入力してください。");
                    }

                    Console.WriteLine("受信メール一覧:");
                    for (int i = 0; i < displayCount; i++)
                    {
                        var emailInfoReq = gmailService.Users.Messages.Get("me", response.Messages[i].Id);
                        var emailInfoResponse = emailInfoReq.Execute();
                        if (emailInfoResponse != null)
                        {
                            string from = "";
                            string date = "";
                            string subject = "";

                            foreach (var mParts in emailInfoResponse.Payload.Headers)
                            {
                                if (mParts.Name == "Date")
                                {
                                    date = mParts.Value;
                                }
                                else if (mParts.Name == "From")
                                {
                                    from = mParts.Value;
                                }
                                else if (mParts.Name == "Subject")
                                {
                                    subject = mParts.Value;
                                }
                            }

                            Console.WriteLine($"日付: {date}");
                            Console.WriteLine($"差出人: {from}");
                            Console.WriteLine($"件名: {subject}");
                            Console.WriteLine("-----------------------------------------------------");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("受信メールはありません。");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"メール一覧の表示中にエラーが発生しました: {ex.Message}");
            }
        }
    }
}