using System.Windows;

namespace Healthcheck_Plus_2
{
    /// <summary>
    /// Provides a dialog enabling the user to select the forename matching
    /// levels that were used for industry suppressions.
    /// </summary>
    public partial class Window_IndustryMatching : Window
    {
        /// <summary>
        /// Constructor for the class.
        /// </summary>
        public Window_IndustryMatching()
        {
            InitializeComponent();
        }

        /// <summary>
        /// When called it returns the selected forename matching level for industry suppressions.
        /// </summary>
        /// <returns>industry suppression forename matching level as string.</returns>
        public string GetIndustryLevel()
        {
            return IndustryLevelText.SelectedValue.ToString();
        }

        /// <summary>
        /// Ends the dialogs operation if a forename matching level has been selected.
        /// </summary>
        /// <param name="sender">Not required</param>
        /// <param name="e">Not required</param>
        private void ConfirmIndustryLevel_Click(object sender, RoutedEventArgs e)
        {
            if (IndustryLevelText.SelectedItem == null)
            {
                MessageBox.Show("Please complete the form.");
            }
            else
            {
                IndustryMatchingForm.Close();
            }
        }
    }
}
