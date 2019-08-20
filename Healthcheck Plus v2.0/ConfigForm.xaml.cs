using System;
using System.IO;
using System.Windows;
using FolderBrowserDialog = System.Windows.Forms.FolderBrowserDialog;

namespace Healthcheck_Plus_2
{
    /// <summary>
    /// Interaction logic for ConfigForm.xaml
    /// </summary>
    public partial class ConfigForm : Window
    {
        /// <summary>
        /// 
        /// </summary>
        public ConfigForm()
        {
            InitializeComponent();

            // Display directories
            Window_Main mainForm = new Window_Main();
            txt_ClientDataDir.Text = mainForm.ReadInDirectories()[0];
            txt_CygnusJobDir.Text = mainForm.ReadInDirectories()[1];
            txt_WfdDir.Text = mainForm.ReadInDirectories()[2];
        }

        /// <summary>
        /// 
        /// </summary>
        public void UpdateDirectories()
        {
            string dirFileLocation = @"P:\Data Services\HealthCheckPlus\Development\HealthCheck2\directories.txt";
            string dirChangeLog = @"P:\Data Services\HealthCheckPlus\Development\HealthCheck2\directory_change_log.txt";

            // Update directories
            using (StreamWriter dirWriter = new StreamWriter(dirFileLocation))
            {
                dirWriter.WriteLine("client data > \"" + txt_ClientDataDir.Text + "\"");
                dirWriter.WriteLine("client cygnus > \"" + txt_CygnusJobDir.Text + "\"");
                dirWriter.WriteLine("wfds > \"" + txt_WfdDir.Text + "\"");
            }

            // Append time of change and user to log
            using (StreamWriter dirLogWriter = new StreamWriter(dirChangeLog, true))
            {
                dirLogWriter.WriteLine("Updated " + DateTime.Now + " by " + Environment.UserName);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Btn_SaveConfig_Click(object sender, RoutedEventArgs e)
        {
            // Save directories to file from form
            UpdateDirectories();

            MessageBox.Show("Directories updated.");

            // restart application
            System.Diagnostics.Process.Start(Application.ResourceAssembly.Location);
            Application.Current.Shutdown();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_ClientDataSelect_Click(object sender, RoutedEventArgs e)
        {
            // Instantiate a folder browser
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            System.Windows.Forms.DialogResult result = dialog.ShowDialog();

            // Add path if user confirms folder
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                txt_ClientDataDir.Text = dialog.SelectedPath + "\\";
            }
        }
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_CygnusJobSelect_Click(object sender, RoutedEventArgs e)
        {
            // Instantiate a folder browser
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            System.Windows.Forms.DialogResult result = dialog.ShowDialog();

            // Add path if user confirms folder
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                txt_CygnusJobDir.Text = dialog.SelectedPath + "\\";
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_WfdDirSelect_Click(object sender, RoutedEventArgs e)
        {
            // Instantiate a folder browser
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            System.Windows.Forms.DialogResult result = dialog.ShowDialog();

            // Add path if user confirms folder
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                txt_WfdDir.Text = dialog.SelectedPath + "\\";
            }
        }
    }
}
