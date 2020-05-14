using Microsoft.Identity.Client;
using System;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;
using Microsoft.Graph.Auth;
using Microsoft.Graph;
using System.Net.Http.Headers;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Navigation;
using System.Diagnostics;
using System.Windows.Markup;

namespace ZoomURLGetter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>

    public partial class MainWindow : Window
    {
        //Set the API Endpoint to Graph 'me' endpoint. 
        // To change from Microsoft public cloud to a national cloud, use another value of graphAPIEndpoint.
        // Reference with Graph endpoints here: https://docs.microsoft.com/graph/deployments#microsoft-graph-and-graph-explorer-service-root-endpoints

        //Set the scope for API call to user.read
        string[] scopes = new string[] { "Mail.Read" };
        GraphServiceClient graphClient = null;
        string content = "";


        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Call AcquireToken - to acquire a token requiring user to sign-in
        /// </summary>
        private async void CallGraphButton_Click(object sender, RoutedEventArgs e)
        {
            graphClient = GetServiceClient();
            if (graphClient != null)
            {
                //var user = graphClient.Me.Request().GetAsync().Result;
                var address_filter = "from/emailAddress/address eq 'no-reply@kibaco.tmu.ac.jp'";
                string enddate = DateTime.Now.ToString("yyyy-MM-dd");
                //string startDate = DateTime.Now.AddDays(-14).ToString("yyyy-MM-dd");
                string startDate = "2020-05-01";
                var date_filter = "ReceivedDateTime ge " + startDate + " and receivedDateTime lt " + enddate;
                var messages = await graphClient.Me.Messages
                .Request()
                .Filter(date_filter + " and " + address_filter)
                .Select(o => new
                {
                    o.Id,
                    o.Subject,
                    o.WebLink,
                    o.Body,
                    o.ToRecipients

                })
                .OrderBy("receivedDateTime desc")
                .GetAsync();
                int c = 0;
                Paragraph parx = new Paragraph();
                var pageIteretor = PageIterator<Message>
                    .CreatePageIterator(graphClient, messages, (m) =>
                    {
                        if (IsContainingZoomURL(m.Body.Content))
                        {
                            content = m.Body.Content;
                            Run r1 = new Run(m.Subject + "\n");
                            Run r2 = new Run("メールを表示" + "\n");
                            Hyperlink hl = new Hyperlink(r2);
                            hl.NavigateUri = new Uri(m.WebLink);
                            hl.RequestNavigate += new RequestNavigateEventHandler(link_RequestNavigate);
                            parx.Inlines.Add(r1);
                            parx.Inlines.Add(hl);
                        }

                        return true;

                    });
                await pageIteretor.IterateAsync();
                FlowDocument document = new FlowDocument();
                document.Blocks.Add(parx);
                richTB.Document = document;
            }
            this.SignOutButton.Visibility = Visibility.Visible;
        }
        private void link_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            var sub = new SubWindow(content);
            sub.ShowDialog();
            //System.Diagnostics.Process.Start(e.Uri.AbsoluteUri.ToString());
            e.Handled = true;
        }

        public bool IsContainingZoomURL(string body)
        {
            return body.Contains("https://zoom.us/j");
        }
        public GraphServiceClient GetServiceClient()
        {
            try
            {
                graphClient = new GraphServiceClient(
                    "https://graph.microsoft.com/v1.0",
                    new DelegateAuthenticationProvider(
                        async (requestMessage) =>
                        {
                            var token = await GetTokenForUserAsync();
                            requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);
                        }
                    )
                );
                return graphClient;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Could not create a graph client: " + ex.Message);
            }
            return graphClient;
        }

        public async Task<string> GetTokenForUserAsync()
        {
            AuthenticationResult authResult = null;
            var app = App.PublicClientApp;
            var accounts = await app.GetAccountsAsync();
            var firstAccount = accounts.FirstOrDefault();

            try
            {
                authResult = await app.AcquireTokenSilent(scopes, firstAccount)
                    .ExecuteAsync();
            }
            catch (MsalUiRequiredException ex)
            {
                // A MsalUiRequiredException happened on AcquireTokenSilent. 
                // This indicates you need to call AcquireTokenInteractive to acquire a token
                System.Diagnostics.Debug.WriteLine($"MsalUiRequiredException: {ex.Message}");

                try
                {
                    authResult = await app.AcquireTokenInteractive(scopes)
                        .WithAccount(accounts.FirstOrDefault())
                        .WithParentActivityOrWindow(new WindowInteropHelper(this).Handle) // optional, used to center the browser on the window
                        .WithPrompt(Microsoft.Identity.Client.Prompt.SelectAccount)
                        .ExecuteAsync();
                }
                catch (MsalException msalex)
                {
                }
            }
            var usertoken = authResult.AccessToken;
            return usertoken;
        }


        /// <summary>
        /// Sign out the current user
        /// </summary>
        private async void SignOutButton_Click(object sender, RoutedEventArgs e)
        {
            var accounts = await App.PublicClientApp.GetAccountsAsync();
            if (accounts.Any())
            {
                try
                {
                    await App.PublicClientApp.RemoveAsync(accounts.FirstOrDefault());
                    this.CallGraphButton.Visibility = Visibility.Visible;
                    this.SignOutButton.Visibility = Visibility.Collapsed;
                }
                catch (MsalException ex)
                {

                }
            }
        }
    }
}
