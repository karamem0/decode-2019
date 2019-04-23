//
// Copyright (c) 2019 karamem0
//
// This software is released under the MIT License.
//
// https://github.com/karamem0/decode2019/blob/master/LICENSE
//

using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;

namespace UserProvisioningSample
{

    public static class UserProvisioningFunction
    {

        // Microsoft Teams のサービス プラン ID
        private static readonly Guid TeamsServicePlanId = new Guid("57ff2da0-773e-42df-b2af-ffb7a2317929");

        // Exchange Online (Plan1) のサービス プラン ID
        private static readonly Guid ExchangeStandardServicePlanId = new Guid("9aaf7827-d63c-4b61-89c3-182f06f82e5c");

        // Exchange Online (Plan2) のサービス プラン ID
        private static readonly Guid ExchangeEnterpriseServicePlanId = new Guid("efb87545-963c-4e0d-99df-69c6916d9eb0");

        [FunctionName("UserProvisioning")]
        public static async Task Run([TimerTrigger("0 */10 * * * *")]TimerInfo timer, ILogger log)
        {
            // アプリケーション設定を取得します
            var appConfig = new ConfigurationBuilder().AddEnvironmentVariables().Build();
            var graphTenantId = appConfig.GetValue<string>("GraphTenantId");
            var graphAuthority = appConfig.GetValue<string>("GraphAuthority");
            var graphClientId = appConfig.GetValue<string>("GraphClientId");
            var graphClientSecret = appConfig.GetValue<string>("GraphClientSecret");
            var graphRedirectUrl = appConfig.GetValue<string>("GraphRedirectUrl");
            var graphScope = appConfig.GetValue<string>("GraphScope");
            var blobStorage = appConfig.GetValue<string>("BlobStorage");
            var blobContainerName = appConfig.GetValue<string>("BlobContainerName");
            var blobFileName = appConfig.GetValue<string>("BlobFileName");
            var teamsGroupId = appConfig.GetValue<string>("TeamsGroupId");
            var eventSenderId = appConfig.GetValue<string>("EventSenderId");

            try
            {
                // OAuth の Access Token を取得します
                var clientCredential = new ClientCredential(graphClientSecret);
                var clientApplication = new ConfidentialClientApplication(
                    graphClientId,
                    graphAuthority + "/" + graphTenantId,
                    graphRedirectUrl,
                    clientCredential,
                    null,
                    null);
                var authenticationResult = await clientApplication.AcquireTokenForClientAsync(graphScope.Split(", "));

                // Microsoft Graph Service Client を初期化します
                var graphClient = new GraphServiceClient(new DelegateAuthenticationProvider(async msg =>
                {
                    await Task.Run(() =>
                    {
                        msg.Headers.Authorization = new AuthenticationHeaderValue("Bearer", authenticationResult.AccessToken);
                    });
                }));

                // BLOB ストレージにファイルが存在するかどうかを確認します
                var storageAccount = CloudStorageAccount.Parse(blobStorage);
                var blobClient = storageAccount.CreateCloudBlobClient();
                var blobContainer = blobClient.GetContainerReference(blobContainerName);
                var blobFile = blobContainer.GetBlockBlobReference(blobFileName);

                var userDeltaCollectionPage = default(IUserDeltaCollectionPage);
                if (await blobFile.ExistsAsync())
                {
                    // ファイルがある場合、前回からの差分のユーザーを取得します
                    userDeltaCollectionPage = new UserDeltaCollectionPage();
                    userDeltaCollectionPage.InitializeNextPageRequest(graphClient, await blobFile.DownloadTextAsync());
                    userDeltaCollectionPage = await userDeltaCollectionPage.NextPageRequest.GetAsync();
                }
                else
                {
                    // ファイルがない場合、すべての差分を取得します
                    userDeltaCollectionPage = await graphClient.Users.Delta().Request().GetAsync();
                }

                // ページングして対象のユーザーのみをフィルターします
                var targetUsers = new List<User>();
                while (true)
                {
                    foreach (var user in userDeltaCollectionPage)
                    {
                        // 削除済みユーザーを除外します
                        if (user.AdditionalData?.ContainsKey("@removed") != true)
                        {
                            // ライセンスのないユーザーを除外します
                            var licenseDetails = await graphClient.Users[user.Id].LicenseDetails.Request().GetAsync();
                            var servicePlans = licenseDetails.SelectMany(ld => ld.ServicePlans).ToList();
                            if (servicePlans.Any(sp => sp.ServicePlanId == TeamsServicePlanId) &&
                                servicePlans.Any(sp => sp.ServicePlanId == ExchangeStandardServicePlanId ||
                                                       sp.ServicePlanId == ExchangeEnterpriseServicePlanId))
                            {
                                targetUsers.Add(user);
                            }
                        }
                    }

                    // 次のページがない場合は処理を抜けます
                    if (userDeltaCollectionPage.NextPageRequest == null)
                    {
                        break;
                    }
                    // 次のページを取得します
                    userDeltaCollectionPage = await userDeltaCollectionPage.NextPageRequest.GetAsync();
                }

                // ユーザーを Teams のメンバーとして追加します
                foreach (var user in targetUsers)
                {
                    try
                    {
                        await graphClient.Groups[teamsGroupId].Members.References.Request().AddAsync(user);
                    }
                    catch (Exception)
                    {
                        // 追加済みの場合があるためエラーは無視します
                    }
                }

                // ユーザーに会議出席依頼を送付します
                if (targetUsers.Any())
                {
                    await graphClient.Users[eventSenderId].Events.Request().AddAsync(new Event()
                    {
                        Attendees = targetUsers.Select(user => new Attendee()
                        {
                            EmailAddress = new EmailAddress()
                            {
                                Address = user.Mail,
                                Name = user.DisplayName
                            },
                            Type = AttendeeType.Required
                        }),
                        Subject = "新入社員向け説明会",
                        Body = new ItemBody()
                        {
                            ContentType = BodyType.Html,
                            Content = "<p>お疲れさまです。</p>"
                                    + "<p>新入社員向け説明会を開催しますのでご参集ください。</p>"
                                    + "<p>よろしくお願いします。</p>",
                        },
                        Start = new DateTimeTimeZone()
                        {
                            DateTime = DateTime.Today.AddDays(2).AddHours(9).ToString("s"),
                            TimeZone = "Tokyo Standard Time"
                        },
                        End = new DateTimeTimeZone()
                        {
                            DateTime = DateTime.Today.AddDays(2).AddHours(18).ToString("s"),
                            TimeZone = "Tokyo Standard Time"
                        },
                    });
                }

                // deltaLink を BLOB ストレージに保存します
                if (userDeltaCollectionPage.AdditionalData.TryGetValue("@odata.deltaLink", out var deltaLink))
                {
                    await blobContainer.CreateIfNotExistsAsync();
                    await blobFile.UploadTextAsync(deltaLink.ToString());
                }
            }
            catch (Exception ex)
            {
                log.LogError(ex.ToString());
            }
        }

    }

}
