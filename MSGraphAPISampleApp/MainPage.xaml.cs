using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.Diagnostics;
using System.Linq;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;

namespace MSGraphAPISampleApp
{
    public sealed partial class MainPage : Page
    {
        public MainPage()
        {
            InitializeComponent();
        }

        private async void SignInButton_Click(object sender, RoutedEventArgs e)
        {
            var account = (await App.PublicClientApplication.GetAccountsAsync())?.FirstOrDefault();
            try
            {
                AuthenticationResult authResult;
                if (account == null)
                {
                    authResult = await App.PublicClientApplication.AcquireTokenAsync(Consts.Scopes);
                }
                else
                {
                    authResult = await App.PublicClientApplication.AcquireTokenSilentAsync(Consts.Scopes, account);
                }

                Debug.WriteLine(authResult.AccessToken);

                // Get user profile.
                var client = new GraphServiceClient(new DelegateAuthenticationProvider(x =>
                {
                    x.Headers.Authorization = new AuthenticationHeaderValue("Bearer", authResult.AccessToken);
                    return Task.CompletedTask;
                }));
                var user = await client.Me.Request().GetAsync();
                textBlockProfile.Text = $"{user.DisplayName}, {user.UserPrincipalName}";
            }
            catch (MsalClientException ex)
            {
                Debug.WriteLine(ex);
            }
        }
    }
}
