using System.Windows;
using System.Windows.Input;

namespace PrimeInsulationBilling.Views
{
    /// <summary>
    /// Interaction logic for LoginWindow.xaml
    /// </summary>
    public partial class LoginWindow : Window
    {
        public LoginWindow()
        {
            InitializeComponent();
        }

        // Handles the click event for the Login button
        private void LoginButton_Click(object sender, RoutedEventArgs e)
        {
            // In a real application, you would replace this with a call to a DatabaseService
            // to securely check the username and hashed password.
            string username = txtUsername.Text;
            string password = txtPassword.Password;

            if (username.Equals("admin", System.StringComparison.OrdinalIgnoreCase) && password == "password123")
            {
                // If login is successful, open the main application window
                MainWindow mainWindow = new MainWindow();

                mainWindow.Show();


                // Close the current login window
                this.Close();
            }
            else
            {
                MessageBox.Show("Invalid username or password. Please try again.", "Login Failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Handles the click event for the custom close button
        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            // Shuts down the entire application
            Application.Current.Shutdown();
        }

        // Allows the user to drag the borderless window by clicking and holding anywhere on it
        protected override void OnMouseLeftButtonDown(MouseButtonEventArgs e)
        {
            base.OnMouseLeftButtonDown(e);
            this.DragMove();
        }
    }
}
