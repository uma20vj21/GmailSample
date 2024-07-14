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
using System.Net.Mail;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace GmailSample
{

    /// <summary>
    /// メール送信プログラムこれをベースにまずはオブジェクト指向プログラミングを行い次にメール一覧機能の作成
    /// </summary>
    class Program
    {
        private static string? _homeDirectory; // ユーザープロファイルのホームディレクトリ
        private static bool isReceivingEmails = false; // メール受信中かどうかを示すフラグ
        private static Task? emailRecivingTask; // メール受信タスクの参照

        /// <summary>
        /// メール受信プログラム
        /// </summary>
        /// <param name="args"></param>
        [Obsolete]
        static void Main(string[] args)
        {
            // GmailとPeople APIサービスの初期化
            InitializeGmailService(out var gmailService, out var peopleService);

            // ユーザーインターフェースのループ処理
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
                        // メール送信
                        SendEmail(gmailService, peopleService);
                        break;
                    case "2":
                        // メール受信開始と停止
                        if (isReceivingEmails)
                        {
                            Console.WriteLine("メールを受信状態にしますか？(y/n)");
                            if (Console.ReadLine()?.ToLower() == "y")
                            {
                                isReceivingEmails = true;
                                emailRecivingTask = Task.Run(() => ReceiveEmails(gmailService));
                            }
                        }
                        else
                        {
                            isReceivingEmails = false;
                            emailRecivingTask?.Wait();
                            Console.WriteLine("メール受信を終了しました。");
                        }
                        break;
                    case "3":
                        // メール一覧表示
                        DisplayEmailList(gmailService);
                        break;
                    case "4":
                        // システム終了
                        if (isReceivingEmails)
                        {
                            isReceivingEmails = false;
                            emailRecivingTask?.Wait();
                        }
                        Console.WriteLine("システムを終了します...");
                        return;
                    default:
                        Console.WriteLine("無効な選択です。もう一度入力してください。");
                        break;

                }
            }

        }

        /// <summary>
        /// GmailおよびPeople APIのサービスを初期化
        /// </summary>
        [Obsolete]
        private static void InitializeGmailService(out GmailService gmailService, out PeopleServiceService peopleService)
        {
            // GmailおよびPeople APIのスコープ
            string[] scopes = { GmailService.Scope.MailGoogleCom,
                                PeopleServiceService.Scope.UserinfoProfile,
                                PeopleServiceService.Scope.UserinfoEmail,
                                PeopleServiceService.Scope.ContactsReadonly,
                                "https://www.googleapis.com/auth/user.emails.read",
                                "https://www.googleapis.com/auth/userinfo.email",
                                "https://www.googleapis.com/auth/userinfo.profile" };
            string appName = "Google.Apis.Gmail.v1 Sample";

            // ユーザーの資格情報を取得
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

            // Gmail APIサービスを初期化
            gmailService = new GmailService(new BaseClientService.Initializer()
            {
                ApplicationName = appName,
                HttpClientInitializer = credential
            });

            // People APIサービスを初期化
            peopleService = new PeopleServiceService(new BaseClientService.Initializer()
            {
                ApplicationName = appName,
                HttpClientInitializer = credential
            });
        }

        /// <summary>
        /// ユーザーからの入力を受け取り、Gmail APIを使用してメールを送信します
        /// </summary>
        private static void SendEmail(GmailService gmailService, PeopleServiceService peopleService)
        {
            try
            {
                var profileRequest = peopleService.People.Get("people/me");
                profileRequest.PersonFields = "names,emailAddress";
                var profile = profileRequest.Execute();

                string mailFromAddress = profile.EmailAddresses[0].Value;
                string mailFromName = profile.Names[0].DisplayName;

                Console.WriteLine("送信先の名前を入力して下さい");
                string mailToName = Console.ReadLine();

                Console.WriteLine("送信先のメールアドレスを入力して下さい");
                string mailToAddress = Console.ReadLine();

                Console.WriteLine("メールの件名を入力して下さい");
                string mailSubject = Console.ReadLine();

                Console.WriteLine("メールの本文を入力して下さい");
                string mailBody = Console.ReadLine();

                // メール送信確認
                Console.WriteLine("メールを送信しますか？(y/n) : ");
                if (Console.ReadLine()?.ToLower() == "y")
                {
                    var mimeMessage = new MimeMessage();
                    mimeMessage.From.Add(new MailboxAddress(mailFromName, mailFromAddress));
                    mimeMessage.To.Add(new MailboxAddress(mailToName, mailToAddress));
                    mimeMessage.Subject = mailSubject;
                    mimeMessage.Body = new TextPart(MimeKit.Text.TextFormat.Plain) { Text = mailBody };

                    try
                    {
                        // メールをエンコードして送信
                        var rawMessage = EncodeMessage(mimeMessage);
                        var result = gmailService.Users.Messages.Send(new Message { Raw = rawMessage }, "me").Execute();

                        // 送信完了メッセージを表示
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

        /// <summary>
        /// Gmail APIを使用して定期的にメールを受信します
        /// </summary>
        private static async Task ReceiveEmails(GmailService gmailService)
        {
            Console.WriteLine("メール受信を開始します...");
            while (isReceivingEmails)
            {
                try
                {
                    // 最も最近のメールを取得
                    UsersResource.MessagesResource.ListRequest request = gmailService.Users.Messages.List("me");
                    request.MaxResults = 1;

                    IList<Message> messages = request.Execute().Messages;
                    if (messages != null && messages.Count > 0)
                    {
                        var message = gmailService.Users.Messages.Get("me", messages[0].Id).Execute();
                        Console.WriteLine("メールを受信しました。");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"メール受信中にエラーが発生しました: {ex.Message}");
                }

                await Task.Delay(30000);
            }
        }

        /// <summary>
        /// ユーザーにメール一覧を表示します
        /// </summary>
        private static void DisplayEmailList(GmailService gmailService)
        {
            // ユーザーが添付ファイル付きのメールのみを表示するかどうかを決定
            bool showOnlyWithAttachments = GetUserInput();
            if (showOnlyWithAttachments)
            {
                ListMessageWithAttachment(gmailService, "me");
            }
            else
            {
                ListMessageWithoutAttachment(gmailService, "me");
            }
        }

        /// <summary>
        /// MimeMessageをBase64URLエンコードしてGmail APIで送信可能な形式に変換します
        /// </summary>
        /// <param name="message">MimeMessageオブジェクト</param>
        /// <returns>エンコードされたメールの文字列</returns>
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

        /// <summary>
        /// 添付ファイル付きのメールのみを表示するかどうかのユーザー入力を取得します
        /// </summary>
        /// <returns>添付ファイル付きのメールのみを表示するかどうかのフラグ</returns>
        private static bool GetUserInput()
        {
            while (true)
            {
                Console.WriteLine("添付ファイルのあるメールのみを表示しますか？ (y/n): ");
                var input = Console.ReadLine();
                if (input?.ToLower() == "y")
                {
                    return true;
                }
                else if (input?.ToLower() == "n")
                {
                    return false;
                }
                else
                {
                    Console.WriteLine("無効な入力です。もう一度入力してください。");
                }
            }
        }

        /// <summary>
        /// Gmail APIを使用して添付ファイルがあるメールリストを表示します
        /// </summary>
        /// <param name="service">Gmailサービス</param>
        /// <param name="userId">ユーザーID</param>
        private static void ListMessageWithAttachment(GmailService service, string userId)
        {
            try
            {
                UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List(userId);
                request.MaxResults = 10;
                IList<Message> messages = request.Execute().Messages;

                if (messages != null && messages.Count > 0)
                {
                    Console.WriteLine("添付ファイルのあるメールタイトル、本文、添付ファイル：");
                    foreach (var message in messages)
                    {
                        var msg = service.Users.Messages.Get(userId, message.Id).Execute();
                        string subject = "";
                        foreach (var header in msg.Payload.Headers)
                        {
                            if (header.Name == "Subject")
                            {
                                subject = header.Value;
                                break;
                            }
                        }

                        string body = GetMessageBody(msg.Payload);
                        IList<string> attachmentTitles = GetAttachmentTitles(service, msg, userId);
                        if (attachmentTitles.Count > 0)
                        {
                            Console.WriteLine($"Subject: {subject}");
                            Console.WriteLine($"Body: {body}");
                            foreach (var attachmentTitle in attachmentTitles)
                            {
                                Console.WriteLine($"Attachment: {attachmentTitle}");
                            }
                            Console.WriteLine("========================");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("添付ファイルのあるメールが見つかりませんでした。");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"メールリスト取得中にエラーが発生しました: {e.Message}");
            }
        }

        /// <summary>
        /// Gmail APIを使用して添付ファイルのないメールリストを表示します
        /// </summary>
        /// <param name="service">Gmailサービス</param>
        /// <param name="userId">ユーザーID</param>
        private static void ListMessageWithoutAttachment(GmailService service, string userId)
        {
            try
            {
                UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List(userId);
                request.MaxResults = 10;
                IList<Message> messages = request.Execute().Messages;

                if (messages != null && messages.Count > 0)
                {
                    Console.WriteLine("メールタイトル、本文：");
                    foreach (var message in messages)
                    {
                        var msg = service.Users.Messages.Get(userId, message.Id).Execute();
                        string subject = "";
                        foreach (var header in msg.Payload.Headers)
                        {
                            if (header.Name == "Subject")
                            {
                                subject = header.Value;
                                break;
                            }
                        }

                        string body = GetMessageBody(msg.Payload);
                        Console.WriteLine($"Subject: {subject}");
                        Console.WriteLine($"Body: {body}");
                        Console.WriteLine("========================");
                    }
                }
                else
                {
                    Console.WriteLine("メールが見つかりませんでした。");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"メールリスト取得中にエラーが発生しました: {e.Message}");
            }
        }

        /// <summary>
        /// メール本文を取得します
        /// </summary>
        /// <param name="payload">メールのペイロード</param>
        /// <returns>メールの本文</returns>
        private static string GetMessageBody(Google.Apis.Gmail.v1.Data.MessagePart payload)
        {
            if (payload.Parts == null)
            {
                // メール本文をデコードして返す
                return payload.Body?.Data == null ? "" : DecodeBase64Url(payload.Body.Data);
            }
            else
            {
                // メールパートをチェックしてテキスト本文を取得
                foreach (var part in payload.Parts)
                {
                    if (part.MimeType == "text/plain")
                    {
                        return part.Body?.Data == null ? "" : DecodeBase64Url(part.Body.Data);
                    }
                }
            }
            return "";
        }

        /// <summary>
        /// メールの添付ファイルのタイトルを取得します
        /// </summary>
        /// <param name="service">Gmailサービス</param>
        /// <param name="msg">メールメッセージ</param>
        /// <param name="userId">ユーザーID</param>
        /// <returns>添付ファイルのタイトルリスト</returns>
        private static IList<string> GetAttachmentTitles(GmailService service, Message msg, string userId)
        {
            List<string> attachmentTitles = new List<string>();
            if (msg.Payload.Parts != null)
            {
                foreach (var part in msg.Payload.Parts)
                {
                    // 添付ファイルのタイトルをリストに追加
                    if (!string.IsNullOrEmpty(part.Filename) && part.Body != null && !string.IsNullOrEmpty(part.Body.AttachmentId))
                    {
                        attachmentTitles.Add(part.Filename);
                    }
                }
            }
            return attachmentTitles;
        }

        /// <summary>
        /// Base64URLエンコードされた文字列をBase64形式にデコードします
        /// </summary>
        /// <param name="input">エンコードされた文字列</param>
        /// <returns>デコードされた文字列</returns>
        private static string DecodeBase64Url(string input)
        {
            // Base64URLの文字列を標準のBase64の形式に変換します。
            // '-' を '+' に、'_' を '/' に置換します。
            string result = input.Replace('-', '+').Replace('_', '/');
            // Base64エンコードされた文字列は、長さが4の倍数である必要なためパディング精査
            switch (result.Length % 4)
            {
                case 2: result += "=="; break;
                case 3: result += "="; break;
            }
            // Base64の文字列をバイト配列に変換し、UTF-8文字列にデコードします。
            var bytes = Convert.FromBase64String(result);
            return System.Text.Encoding.UTF8.GetString(bytes);
        }

    }
}
