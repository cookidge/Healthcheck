using System.Windows;

namespace Healthcheck_Plus_2
{
    /// <summary>
    /// Provides a dialog enabling the user to select the forename matching
    /// levels that were used for both deduping and suppressing within a dedupe module.
    /// </summary>
    public partial class Window_ForenameMatching : Window
    {

        /// <summary>
        /// Constructor for the class.
        /// </summary>
        public Window_ForenameMatching()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Updates the dialog with the name of the dedupe.
        /// </summary>
        /// <param name="text">name of the dedupe as a string.</param>
        public void SetWindowText(string text)
        {
            ForenameMatchingText.Content = text;
        }

        /// <summary>
        /// When called it returns the selected dedupe level for the current dedupe.
        /// </summary>
        /// <returns>dedupe forename matching level as string.</returns>
        public string GetDedupingLevel()
        {
            return DedupingLevelText.SelectedValue.ToString();
        }

        /// <summary>
        /// When called it returns the selected suppression level for the current dedupe.
        /// </summary>
        /// <returns>suppression forename matching level as string.</returns>
        public string GetSuppressionLevel()
        {
            return SuppressingLevelText.SelectedValue.ToString();
        }

        /// <summary>
        /// Ends the dialogs operation if both drop down boxes have level selected.
        /// </summary>
        /// <param name="sender">Not required</param>
        /// <param name="e">Not required</param>
        private void ConfirmMatchLevel_Click(object sender, RoutedEventArgs e)
        {
            if ((DedupingLevelText.SelectedItem == null) || (SuppressingLevelText.SelectedItem == null))
            {
                MessageBox.Show("Please complete the form.");
            }
            else
            {
                ForenameMatchingForm.Close();
            }
        }

      
    }
}
