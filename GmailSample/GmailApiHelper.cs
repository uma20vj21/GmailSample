//using System.Collections.Generic;
//using System.IO;

//using System.Threading;
//using Google.Apis.Auth.OAuth2;
//using Google.Apis.Gmail.v1;
//using Google.Apis.Gmail.v1.Data;
//using Google.Apis.Services;
//using Google.Apis.Util.Store;

//namespace GmailSample;

//public class GmailApiHelper
//{
//    // 外部からのアクセスできないことを示すため命名の最初にアンダースコアを入れている
//    private readonly GmailService _service;
//    //private readonly string _homeDirectory;
//    //private const string ApplicationName = "Gmail API Console";
//    //private const string User = "user";

//    // メール情報
////    string mail_from_name = "大嶋由真";
////    string mail_from_address = "(差出人アドレス)";
////    string mail_to_name = "宛先";
////    string mail_to_address = "(宛先アドレス)";
////    string mail_subject = "テストメール";
////    string mail_body = @"テストメールを送信します。
////受信確認できましたら、返信をお願いします。

////テスト太郎";

//    public GmailApiHelper(UserCredential credential, string appName)
//    {
//        _service = new GmailService(new BaseClientService.Initializer()
//        {
//            ApplicationName = appName,
//            HttpClientInitializer = credential
//        });;

//        //_homeDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
//        //UserCredential credential;

//        //using (var stream = new FileStream(Path.Combine(_homeDirectory, "Downloads", "clientsecret.json"), FileMode.Open, FileAccess.Read))
//        //{
//        //    string credpath = "token.json";
//        //    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
//        //            GoogleClientSecrets.Load(stream).Secrets,
//        //            new[] { GmailService.Scope.MailGoogleCom },
//        //            "user",
//        //            CancellationToken.None,
//        //            new FileDataStore(credpath, true)).Result;
//        //}

//        // GmailServiceの初期化
//        //_service = new GmailService(new BaseClientService.Initializer()
//        //{
//        //    HttpClientInitializer = credential,
//        //    ApplicationName = ApplicationName
//        //});
//    }



//    public IList<Message> GetMessages(int maxResults = 10)
//    {
//        UsersResource.MessagesResource.ListRequest request = _service.Users.Messages.List(User);
//        request.MaxResults = maxResults;

//        return request.Execute().Messages;
//    }

//}
 