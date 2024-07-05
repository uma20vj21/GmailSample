using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.PeopleService.v1;
using Google.Apis.PeopleService.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util;
using Google.Apis.Util.Store;
using MimeKit;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Mail;
using System.Text;
using System.Threading;

namespace GmailSample
{

    /// <summary>
    /// メール送信プログラムこれをベースにまずはオブジェクト指向プログラミングを行い次にメール一覧機能の作成
    /// </summary>
    class Program
    {
        private static string? _homeDirectory;

        [Obsolete]
        static void Main(string[] args)
        {
            //認証情報(20240704_Google People APIのスコープ追加)
            string[] scopes = { GmailService.Scope.MailGoogleCom,
                                PeopleServiceService.Scope.UserinfoProfile,
                                PeopleServiceService.Scope.UserinfoEmail,
                                PeopleServiceService.Scope.ContactsReadonly,
                                "https://www.googleapis.com/auth/user.emails.read",
                                "https://www.googleapis.com/auth/userinfo.email",
                                "https://www.googleapis.com/auth/userinfo.profile"
            };
            string appName = "Google.Apis.Gmail.v1 Sample";

            //認証
            UserCredential credential;

            // ホームディレクトリのダウンロードにファイルが置いてあるか確認
            _homeDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string clientSecretPath = Path.Combine(_homeDirectory, "Downloads", "clientsecret.json");
            if (!File.Exists(clientSecretPath))
            {
                // macOSのホームディレクトリのダウンロードフォルダから参照
                clientSecretPath = Path.Combine(_homeDirectory, "Downloads", "clientsecret.json");
                if (!File.Exists(clientSecretPath))
                {
                    Console.WriteLine("クライアントシークレットファイルが見つかりません。");
                    return;
                }
            }

            // 古い認証情報の削除
            string credPath = Path.Combine(_homeDirectory, ".credentials/gmail-dotnet-quickstart.json");
            if (File.Exists(credPath))
            {
                File.Delete(credPath);
            }

            // クライアントシークレットファイルを読み込んで認証を取得
            using (var stream = new FileStream(clientSecretPath, FileMode.Open, FileAccess.Read))
            {
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.Load(stream).Secrets,
                        scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
            }
            // GmailServiceの初期化
            var gmailService = new GmailService(new BaseClientService.Initializer()
            {
                ApplicationName = appName,
                HttpClientInitializer = credential
            });

            // PeopleServiceの初期化
            var peopleService = new PeopleServiceService(new BaseClientService.Initializer()
            {
                ApplicationName = appName,
                HttpClientInitializer = credential
            });

            // 認証されたアカウントのプロフィール情報を取得
            try
            {
                var profileRequest = peopleService.People.Get("people/me");
                profileRequest.PersonFields = "names,emailAddresses";
                var profile = profileRequest.Execute();

                // ユーザーのメールアドレスと表示名を取得
                string mailFromAddress = profile.EmailAddresses[0].Value;
                string mailFromName = profile.Names[0].DisplayName;

                // ユーザーに送信先のメールアドレスとメールの内容を入力してもらう
                Console.WriteLine("送信先の名前を入力して下さい");
                string mailToName = Console.ReadLine();

                Console.WriteLine("送信先のメールアドレスを入力して下さい");
                string mailToAddress = Console.ReadLine();

                Console.WriteLine("メールの件名を入力してください");
                string mailSubject = Console.ReadLine();

                Console.WriteLine("メールの本文を入力してください: ");
                string mailBody = Console.ReadLine();

                // メール送信確認
                Console.WriteLine("メールを送信しますか？(y/n)：");
                string sendMailResponse = Console.ReadLine().ToLower();

                if (sendMailResponse == "y")
                {
                    // MimeMessageを作成してメールを送信
                    var mime_message = new MimeMessage();
                    mime_message.From.Add(new MailboxAddress(mailFromName, mailFromAddress));
                    mime_message.To.Add(new MailboxAddress(mailToName, mailToAddress));
                    mime_message.Subject = mailSubject;
                    mime_message.Body = new TextPart(MimeKit.Text.TextFormat.Plain)
                    {
                        Text = mailBody
                    }; // プレーンテキスト形式でメール本文を作成

                    try
                    {
                        //メール送信
                        var rawMessage = EncodeMessage(mime_message); // メール本文を適切な形式にエンコーディング
                        var result = gmailService.Users.Messages.Send(
                        new Message()
                        {
                            Raw = rawMessage
                        },
                        "me"
                        ).Execute();

                        // メール送信中の表示
                        Console.WriteLine("送信中です。");

                        // 送信完了後の表示
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
                else
                {
                    // ユーザー入力を促す
                    bool showOnlyWithAttachments = GetUserInput();

                    // メールの一覧を取得して表示
                    if (showOnlyWithAttachments)
                    {
                        ListMessageWithAttachment(gmailService, "me");
                    }
                    else
                    {
                        ListMessageWithoutAttachment(gmailService, "me");
                    }
                }

            }
            catch (Google.GoogleApiException e)
            {
                Console.WriteLine($"プロフィール情報の取得中にエラーが発生しました: {e.Message}");
                Console.WriteLine("認証スコープが不足している可能性があります。スコープを確認してください。");
            }

            Console.WriteLine("プログラムを終了するにはEnterキーを押してください...");
            Console.ReadLine(); // ここでEnterキーの入力を待つようにする
        }
        // MimeMessageをBase64エンコードして文字列として取得するメソッド
        private static string EncodeMessage(MimeMessage mimeMessage)
        {
            using (var memoryStream = new MemoryStream())
            {
                mimeMessage.WriteTo(memoryStream); // MimeMessageをMemoryStreamに書き込む
                return Convert.ToBase64String(memoryStream.ToArray()) // MemoryStreamをバイト配列に変換
                    .Replace('+', '-')
                    .Replace('/', '_')
                    .Replace("=", "");
            }
        }
        // ユーザー入力を取得するメソッド
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

        // 添付ファイルがあるメールのみを表示させるメソッド
        private static void ListMessageWithAttachment(GmailService service, string userId)
        {
            try
            {
                // userIdは"me"と表示される
                Console.WriteLine(userId);
                // メールリストのリクエスト
                UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List(userId);
                request.MaxResults = 10; // 最大10件のメールを取得

                // メールリストの取得
                IList<Message> messages = request.Execute().Messages;

                if (messages != null && messages.Count > 0)
                {
                    Console.WriteLine("添付ファイルのあるメールのタイトルと本文");
                    foreach (Message message in messages)
                    {
                        // メールの取得
                        var msg = service.Users.Messages.Get(userId, message.Id).Execute();
                        // メールのタイトルを取得
                        string subject = "";
                        foreach (var header in msg.Payload.Headers)
                        {
                            if (header.Name == "Subject")
                            {
                                subject = header.Value;
                                break;
                            }
                        }
                        // メール本文を取得
                        string body = GetMessageBody(msg.Payload);

                        // 添付ファイルのタイトルを取得
                        IList<string> attachmentTitles = GetMessageAttachments(msg.Payload);

                        if (attachmentTitles.Count > 0)
                        {
                            Console.WriteLine($"タイトル：{subject}");
                            Console.WriteLine("添付ファイルのタイトル：");
                            foreach (var title in attachmentTitles)
                            {
                                Console.WriteLine(title);
                            }
                            Console.WriteLine($"本文：{body}");
                            Console.WriteLine();
                        }
                    }
                }
                else
                {
                    Console.WriteLine("メールが見つかりませんでした");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"メールの取得中にエラーが発生しました: {ex.Message}");
            }
        }

        // 添付ファイルのないメールのみを表示するメソッド
        private static void ListMessageWithoutAttachment(GmailService service, string userId)
        {
            try
            {
                // メールリストのリクエスト
                UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List(userId);
                request.MaxResults = 10; // 最大10件のメールを取得

                // メールリストの取得(上記のMaxResultsで設定した件数分のリストを取得してくる)
                IList<Message> messages = request.Execute().Messages;
                // 取得したメールリストがnullでなく最低1件以上含まれているかチェック
                if (messages != null && messages.Count > 0)
                {
                    Console.WriteLine("添付ファイルのないメールタイトル、本文：");
                    foreach (var message in messages)
                    {
                        // メールの取得
                        var msg = service.Users.Messages.Get(userId, message.Id).Execute();
                        // メールのタイトルを取得
                        string subject = "";
                        foreach (var header in msg.Payload.Headers)
                        {
                            if (header.Name == "Subject")
                            {
                                subject = header.Value;
                                break;
                            }
                        }
                        // メールの本文を取得
                        string body = GetMessageBody(msg.Payload);

                        // 添付ファイルのタイトルを取得
                        IList<string> attachmentTitles = GetMessageAttachments(msg.Payload);
                        // 添付ファイルが含まれていないかのチェック
                        if (attachmentTitles.Count == 0)
                        {
                            Console.WriteLine($"タイトル：{subject}");
                            Console.WriteLine($"本文：{body}");
                            Console.WriteLine();
                        }
                    }

                }
                else
                {
                    Console.Write("メールが見つかりませんでした");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"メールの取得中にエラーが発生しました：{ex.Message}");
            }
        }

        // メールの本文を取得するメソッド(20240629_再記述）
        // MessagePartの型が曖昧な参照になっていたため明示的に指定
        private static string GetMessageBody(Google.Apis.Gmail.v1.Data.MessagePart payload)
        {
            try
            {
                if (payload.Parts == null)
                {
                    return DecodeBase64Url(payload.Body.Data);
                }
                foreach (var part in payload.Parts)
                {
                    if (part.MimeType == "text/plain" && part.Body != null && part.Body.Data != null)
                    {
                        return DecodeBase64Url(part.Body.Data);
                    }
                    else if (part.MimeType == "text/html" && part.Body != null && part.Body.Data != null)
                    {
                        return DecodeBase64Url(part.Body.Data);
                    }
                    else if (part.Parts != null)
                    {
                        string result = GetMessageBody(part); // 再帰的に検索
                        if (!string.IsNullOrEmpty(result))
                        {
                            return result;
                        }
                    }

                }
                return "";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"メールの本文の取得中にエラーが発生しました: {ex.Message}");
                return "";
            }
        }

        // Base64デコードを行うメソッド
        public static string DecodeBase64Url(string input)
        {
            string s = input.Replace('-', '+').Replace('_', '/');
            switch (s.Length % 4)
            {
                case 2: s += "=="; break;
                case 3: s += "="; break;
            }
            byte[] bytes = Convert.FromBase64String(s);
            return Encoding.UTF8.GetString(bytes); // UTF8エンコード
        }

        // メールの添付ファイルのタイトルを取得するメソッド
        // MessagePartの型が曖昧な参照になっていたため明示的に指定

        private static IList<string> GetMessageAttachments(Google.Apis.Gmail.v1.Data.MessagePart payload)
        {
            IList<string> attachmentTitles = new List<string>();
            try
            {
                if (payload.Parts == null)
                {
                    if (!string.IsNullOrEmpty(payload.Filename))
                    {
                        attachmentTitles.Add(payload.Filename);
                    }
                }
                else
                {
                    foreach (var part in payload.Parts)
                    {
                        if (!string.IsNullOrEmpty(part.Filename))
                        {
                            attachmentTitles.Add(part.Filename);
                        }
                        else if (part.Parts != null)
                        {
                            foreach (var subPart in part.Parts)
                            {
                                if (!string.IsNullOrEmpty(subPart.Filename))
                                {
                                    attachmentTitles.Add(subPart.Filename);
                                }
                            }
                        }
                    }
                }
                return attachmentTitles;

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return attachmentTitles;
            }

        }

    }

}
