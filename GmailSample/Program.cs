using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using MimeKit;
using System;
using System.Collections.Generic;
using System.IO;
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
            //認証情報
            string[] scopes = { GmailService.Scope.MailGoogleCom };
            string app_name = "Google.Apis.Gmail.v1 Sample";

            //メール情報
            string mail_from_name = "大嶋由真";
            string mail_from_address = "yumastudy.0201@gmail.com";
            string mail_subject = "テストメール";
            string mail_body = @"テストメールを送信します。
               受信確認できましたら、返事をお願いします。
               ちなみにこれってスニペットで本文全て取得できているんですかね？
               あと、メールの本文取得メソッドも追加してみました。
               確認お願いします。
 
               テスト太郎より";

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
            using (var stream = new FileStream(Path.Combine(_homeDirectory, "Downloads", "clientsecret.json"), FileMode.Open, FileAccess.Read))
            {
                string credpath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.Load(stream).Secrets,
                        new[] { GmailService.Scope.MailGoogleCom },
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credpath, true)).Result;

                // GmailServiceの初期化
                var service = new GmailService(new BaseClientService.Initializer()
                {
                    ApplicationName = app_name,
                    HttpClientInitializer = credential
                });

                // ユーザー入力で送信先メールアドレスと名前を取得
                Console.WriteLine("送信先の名前を入力して下さい");
                string mailToName = Console.ReadLine();
                Console.WriteLine("送信先のメールアドレスを入力して下さい");
                string mailToAddress = Console.ReadLine();

                // メール送信確認
                Console.WriteLine("メールを送信しますか？(y/n)：　");
                string sendMailResponse = Console.ReadLine().ToLower();

                if (sendMailResponse == "y")
                {
                    // メール作成
                    var mime_message = new MimeMessage();
                    mime_message.From.Add(new MailboxAddress(mail_from_name, mail_from_address));
                    mime_message.To.Add(new MailboxAddress(mailToName, mailToAddress));
                    mime_message.Subject = mail_subject;
                    var text_part = new TextPart(MimeKit.Text.TextFormat.Plain);
                    text_part.SetText(Encoding.UTF8, mail_body); // UTF-8エンコーディングを使用
                    mime_message.Body = text_part;

                    byte[] bytes;
                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        mime_message.WriteTo(memoryStream); // MimeMessageをMemoryStreamに書き込む
                        bytes = memoryStream.ToArray(); // MemoryStreamをバイト配列に変換
                    }

                    string raw_message = Convert.ToBase64String(bytes)
                        .Replace('+', '-')
                        .Replace('/', '_')
                        .Replace("=", "");

                    //メール送信
                    var result = service.Users.Messages.Send(
                    new Message()
                    {
                        Raw = raw_message
                    },
                    "me"
                    ).Execute();

                    Console.WriteLine("送信完了しました。");
                    Console.WriteLine("Message ID: {0}", result.Id);

                    // 送信したメールの内容を表示
                    Console.WriteLine("送信したメールの内容を表示します。");
                    Console.WriteLine("========================");

                    // メールの取得
                    var message = service.Users.Messages.Get("me", result.Id).Execute();

                    // メールの送信者とタイトルを取得
                    string from = "";
                    string subject = "";

                    //　取得するためにMessagePartHeaderクラスに
                    //　Name プロパティが存在するかどうか確認をする
                    foreach (var header in message.Payload.Headers)
                    {
                        if (header.Name == "From")
                        {
                            from = header.Value;
                        }
                        else if (header.Name == "Subject")
                        {
                            subject = header.Value;
                        }
                    }

                    // メールの内容表示
                    Console.WriteLine("from: " + from);
                    Console.WriteLine("subject: " + subject);
                    // 
                    Console.WriteLine("Message snippet: " + message.Snippet);
                    Console.WriteLine("Message payload: " + message.Payload);
                    Console.WriteLine("========================");
                }

                

                // ユーザー入力を促す
                bool showOnlyWithAttachments = GetUserInput();

                // メールの一覧を取得して表示
                if (showOnlyWithAttachments)
                {
                    ListMessageWithAttachment(service, "me");
                }
                else 
                {
                    ListMessageWithoutAttachment(service, "me");
                }
            }
            Console.WriteLine("プログラムを終了するにはEnterキーを押してください...");
            Console.ReadLine(); // ここでEnterキーの入力を待つようにする
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
        private static void ListMessageWithAttachment(GmailService service,string userId) 
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
        private static string GetMessageBody(Google.Apis.Gmail.v1.Data.MessagePart payload) {
            try {
                if(payload.Parts == null) 
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

                }catch (Exception ex)
                {
                    Console.WriteLine(ex);
                    return attachmentTitles;
                }
            
        }
    }
}