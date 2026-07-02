using CostAnalysis.Services;
using System.Windows;

namespace CostAnalysis.UI
{
    public partial class LoginWindow : Window
    {
        private readonly AuthService _auth;

        public string EnteredUsername { get; private set; }

        public LoginWindow(AuthService authService)
        {
            InitializeComponent();
            _auth = authService;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void SignIn_Click(object sender, RoutedEventArgs e)
        {
            var username = TB_Username.Text?.Trim();
            var password = PB_Password.Password;

            if (_auth.ValidateCredentials(username, password, out var matchedUser, out var error))
            {
                _auth.CurrentUser = matchedUser;
                EnteredUsername = matchedUser.Username;
                DialogResult = true;
                Close();
            }
            else
            {
                TB_Error.Text = error;
            }
        }

        private void SignUp_RequestNavigate(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = e.Uri.AbsoluteUri,
                    UseShellExecute = true
                });
                e.Handled = true;
            }
            catch { }
        }
    }
}
