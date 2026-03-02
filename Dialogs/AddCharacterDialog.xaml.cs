using System.Windows;
using System.Windows.Input;

namespace Drops_Tracker
{
    public partial class AddCharacterDialog : Window
    {
        public string CharacterName { get; private set; } = string.Empty;
        public string CharacterClass { get; private set; } = string.Empty;
        public int CharacterLevel { get; private set; }

        public AddCharacterDialog(
            string characterName = "",
            string characterClass = "",
            int characterLevel = 0,
            bool isEditMode = false)
        {
            InitializeComponent();
            NameTextBox.Text = characterName;
            ClassTextBox.Text = characterClass;
            LevelTextBox.Text = characterLevel > 0 ? characterLevel.ToString() : string.Empty;

            if (isEditMode)
            {
                Title = "Edit Character";
                HeaderIconTextBlock.Text = "\u270E";
                HeaderTitleTextBlock.Text = "Edit Character";
                AddButton.Content = "Save Changes";
            }

            NameTextBox.Focus();
            NameTextBox.CaretIndex = NameTextBox.Text.Length;
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                Add_Click(sender, e);
            }
            else if (e.Key == Key.Escape)
            {
                Cancel_Click(sender, e);
            }
        }

        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ButtonState == MouseButtonState.Pressed)
            {
                DragMove();
            }
        }

        private void Add_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(NameTextBox.Text))
            {
                MessageBox.Show("Please enter a character name.", "Validation Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                NameTextBox.Focus();
                return;
            }

            CharacterName = NameTextBox.Text.Trim();
            CharacterClass = ClassTextBox.Text.Trim();

            if (int.TryParse(LevelTextBox.Text.Trim(), out int level))
            {
                CharacterLevel = level;
            }

            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
