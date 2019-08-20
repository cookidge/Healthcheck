using CDO;
using Ionic.Zip;
using Microsoft.Win32;
using System;
using System.Collections;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Animation;
using System.Xml;
using UsingForms = System.Windows.Forms;
using VBInteraction = Microsoft.VisualBasic.Interaction;

namespace Healthcheck_Plus_2
{
    /// <summary>
    /// Contains the operations and all application methods.
    /// </summary>
    public partial class Window_Main : Window
    {
        /// <summary>
        /// a private property that facilitates the fading of form controls in the class animation methods.
        /// </summary>
        private Storyboard fading;
        /// <summary>
        /// a private property that facilitates the moving of form controls in the class animation methods.
        /// </summary>
        private Storyboard moving;
        /// <summary>
        /// a private property that facilitates the appearing of form controls in the class animation methods.
        /// </summary>
        private Storyboard appearing;
        /// <summary>
        /// a private property that holds the number of download files selected by the user.
        /// </summary>
        private int fileCount;

        // Class constructor

        /// <summary>
        /// The constructor for the class, it performs the following operations:
        /// 
        /// <para>
        /// <list class="bullet">
		///	    <listItem>
		///		    <para> - Sets the window properties of the form.</para>
		///		</listItem>
        ///     <listItem>
        ///         <para> - Checks the users installation files for GMC</para>
        ///     </listItem>
        ///     <listItem>
        ///         <para> - Populates drop down with list of jobs</para>
        ///     </listItem>
        ///     <listItem>
        ///         <para> - Populates list of recently reported jobs</para>
        ///     </listItem>
		///	</list>
        ///	</para>
        /// </summary>
        /// <example>
        /// <code>
        ///     // Initialise
        ///     InitializeComponent();
        ///
        ///     // Set window properties
        ///     JobDescription.Visibility = Visibility.Hidden;
        ///     MailingFileList.Opacity = 0.0;
        ///     FilesConfirm.Opacity = 0.0;
        ///     CreateReport.Opacity = 0.0;
        ///
        ///     // Check Installation
        ///     if (!CheckInstallations())
        ///     {
        ///        MessageBoxResult result = new MessageBoxResult();
        ///        result = MessageBoxResult.No;
        ///
        ///        while (result == MessageBoxResult.No)
        ///        {
        ///            result = MessageBox.Show("In order to run Healthcheck Plus, you need GMC Inspire Designer installed on your system. Closing...", "Exiting Application", MessageBoxButton.OK);
        ///        }
        ///
        ///         Application.Current.Shutdown();
        ///      }
        ///
        ///      // Populate job list
        ///      RefreshJobList();
        ///
        ///      // Recently reported jobs
        ///      RecentlyReportedJobs();
        /// </code>
        /// </example>
        public Window_Main()
        {
            // Initialise
            InitializeComponent();

            // Set window properties
            JobDescription.Visibility = Visibility.Hidden;
            MailingFileList.Opacity = 0.0;
            FilesConfirm.Opacity = 0.0;
            CreateReport.Opacity = 0.0;

            // Check Installation
            if (!CheckInstallations())
            {
                MessageBoxResult result = new MessageBoxResult();
                result = MessageBoxResult.No;

                while (result == MessageBoxResult.No)
                {
                    result = MessageBox.Show("In order to run Healthcheck Plus, you need GMC Inspire Designer installed on your system. Closing...", "Exiting Application", MessageBoxButton.OK);
                }

                Application.Current.Shutdown();
            }

            // Populate job list
            RefreshJobList();

            // Recently reported jobs
            RecentlyReportedJobs();

            
        }

        // Operations

        /// <summary>
        /// Accepts the selected job from the drop down list on the form, and removes unwanted information
        /// from the _Merged.xml file to prevent any issues with characters. If industry suppressions exist then it
        /// starts the processing methods to create both the count and cost CSVs. The downloaded files that
        /// exist on the Cygnus workflow are then displayed to the user on the form.
        /// </summary>
        /// <param name="sender">not required</param>
        /// <param name="e">not required</param>
        /// <example><code>
        ///     //Check that a job has been selected
        ///     if (JobSelect.SelectionBoxItem.ToString() == "")
        ///     { 
        ///         UserInformation.Content = "Please select a job from the list.";
        ///     }
        ///     else
        ///     {
        ///         UserInformation.Content = "";
        ///     
        ///         // Clean XML File  
        ///         MergedXMLCleaner();
        ///     
        ///         // Check files exist
        ///         if (CheckDownloadsExist())
        ///         {
        ///             // Process industry suppression information [if applicable]
        ///             if (CheckISExist())
        ///         {
        ///             // Process suppression information
        ///             ProcessIndustrySuppressions();
        ///         }
        ///     
        ///         // Retrieve and display downloaded files on the Cygnus workflow
        ///         FilesOnWorkflow();
        ///     
        ///         } else
        ///         {
        ///             UserInformation.Content = "There are no downloaded files on the Cygnus workflow.";
        ///         }
        ///     
        ///     }
        ///     
        ///     // Re-populate job list
        ///     RefreshJobList();
        /// </code></example>
        private void JobConfirm_Click(object sender, RoutedEventArgs e)
        {
            //Check that a job has been selected
            if (JobSelect.SelectionBoxItem.ToString() == "")
            { 
                UserInformation.Content = "Please select a job from the list.";
            }
            else
            {
                UserInformation.Content = "";
                Menu_Config.Visibility = Visibility.Hidden;
                MenuItem_Main.IsEnabled = false;
                MenuItem_Main.Visibility = Visibility.Hidden;

                // Clean XML File  
                MergedXMLCleaner();

                // Check files exist
                if (CheckDownloadsExist())
                {
                    // Process industry suppression information [if applicable]
                    if (CheckISExist())
                    {
                        // Process suppression information
                        ProcessIndustrySuppressions();
                    }

                    // Retrieve and display downloaded files on the Cygnus workflow
                    FilesOnWorkflow();

                } else
                {
                    UserInformation.Content = "There are no downloaded files on the Cygnus workflow.";
                }

            }

            // Re-populate job list
            RefreshJobList();

        }

        /// <summary>
        /// Checks that at least one file has been selected, then kicks off the main processing
        /// </summary>
        /// <param name="sender">not required</param>
        /// <param name="e">not required</param>
        /// <example>
        /// <code>
        ///     if (MailingFileList.SelectedItems.Count == 0) 
        ///     {
        ///         UserInformation.Content = "There are no files selected.";
        ///     }
        ///     else
        ///     {
        ///         UserInformation.Content = "";
        ///     
        ///         // Run file processing
        ///         MainFileProcessing();
        ///     }
        /// </code>
        /// </example>
        private void FilesConfirm_Click(object sender, RoutedEventArgs e)
        {
            if (MailingFileList.SelectedItems.Count == 0) 
            {
                UserInformation.Content = "There are no files selected.";
            }
            else
            {
                UserInformation.Content = "";
                
                // Run file processing
                MainFileProcessing();
            }
        }

        /// <summary>
        /// Checks that at least one dedupe module has been selected, loops through each dedupe selected and
        /// retrieves the information for it from the _Merged.xml file produced by Cygnus, 
        /// then calls the methods that export the information, statistics, and drops matrix for 
        /// each dedupe that was selected. The total methods for each dedupe, and the files overall are then executed.
        /// </summary>
        /// <param name="sender">not required</param>
        /// <param name="e">not required</param>
        /// <example>
        /// <code>
        ///     if (MailingFileList.SelectedItems.Count == 0)
        ///     {
        ///         UserInformation.Content = "There are no modules selected.";
        ///     }
        ///     else
        ///     {
        ///         UserInformation.Content = "";
        ///     
        ///         // Process each dedupe
        ///         foreach (string module in MailingFileList.SelectedItems)
        ///         {
        ///             // Open _Merged.XML
        ///             XmlDocument xml_document = new XmlDocument();
        ///             xml_document.Load(MergedXMLPath());
        ///     
        ///             // Retrieve process node
        ///             XmlNodeList process_list = xml_document.DocumentElement.SelectNodes("/cygnus_reports/process");
        ///     
        ///             // Loop through each process
        ///             foreach (XmlNode process in process_list)
        ///             {
        ///                 // Store process attributes
        ///                 XmlAttributeCollection process_attributes = process.Attributes;
        ///                 string process_type = "";
        ///                 string process_name = "";
        ///     
        ///                 // loop through attributes
        ///                 foreach (XmlAttribute attribute in process_attributes)
        ///                 {
        ///                     // Check attribute field
        ///                     switch (attribute.Name)
        ///                     {
        ///                         // if 'type' store contents (i.e. download)
        ///                         case "type":
        ///                             process_type = attribute.Value;
        ///                             break;
        ///                         // If 'name' store contents (i.e. the filename)
        ///                         case "name":
        ///                             process_name = attribute.Value;
        ///     
        ///                             if (process_name == module)
        ///                             {
        ///                                 // Export Dedupe Information [Settings]
        ///                                 ExportDedupeInformation(process);
        ///                                 // Export Dedupe Statistics [File Totals]
        ///                                 ExportDedupeStatistics(process);
        ///                                 // Export Dedupe Examples [Selected Groups]
        ///                                 ExportDedupeExamples(process);
        ///                                 // Export Drops Matrix [Single]
        ///                                 ExportDropsMatrix(process);
        ///                             }
        ///                             break;
        ///                     }
        ///                 }
        ///             }
        ///     }
        ///     
        ///     // Calculate dedupe totals
        ///     CalculateDedupeFileTotals();
        ///     
        ///     // Calculate dedupe totals
        ///     CalculateFinalDedupeTotals();
        ///     
        ///     // End Dedupe Processing Animations
        ///     EndDedupeAnimations();
        ///     
        ///     // Produce summary animations
        ///     SummaryAnimations();
        ///     
        /// </code>
        /// </example>
        private void RunDedupe_Click(object sender, RoutedEventArgs e)
        {
            if (MailingFileList.SelectedItems.Count == 0)
            {
                UserInformation.Content = "There are no modules selected.";
            }
            else
            {
                UserInformation.Content = "";

                // Process each dedupe
                foreach (string module in MailingFileList.SelectedItems)
                {
                   
                    // Open _Merged.XML
                    XmlDocument xml_document = new XmlDocument();
                    xml_document.Load(MergedXMLPath());

                    // Retrieve process node
                    XmlNodeList process_list = xml_document.DocumentElement.SelectNodes("/cygnus_reports/process");

                    // Loop through each process
                    foreach (XmlNode process in process_list)
                    {
                        // Store process attributes
                        XmlAttributeCollection process_attributes = process.Attributes;
                        string process_type = "";
                        string process_name = "";

                        // loop through attributes
                        foreach (XmlAttribute attribute in process_attributes)
                        {
                            // Check attribute field
                            switch (attribute.Name)
                            {
                                // if 'type' store contents (i.e. download)
                                case "type":
                                    process_type = attribute.Value;
                                    break;
                                // If 'name' store contents (i.e. the filename)
                                case "name":
                                    process_name = attribute.Value;

                                    if (process_name == module)
                                    {
                                        // Export Dedupe Information [Settings]
                                        ExportDedupeInformation(process);
                                        // Export Dedupe Statistics [File Totals]
                                        ExportDedupeStatistics(process);
                                        // Export Dedupe Examples [Selected Groups]
                                        ExportDedupeExamples(process);
                                        // Export Drops Matrix [Single]
                                        ExportDropsMatrix(process);
                                    }

                                    break;
                            }
                        }
                    }
                }

                // Calculate dedupe totals
                CalculateDedupeFileTotals();

                // Calculate dedupe totals
                CalculateFinalDedupeTotals();

                // End Dedupe Processing Animations
                EndDedupeAnimations();

                // Produce summary animations
                SummaryAnimations();

            }          
        }

        /// <summary>
        /// The final button that the user uses on the form, it calls the methods that export the basic job information,
        /// create and compress the report before sending to the project manager.
        /// </summary>
        /// <param name="sender">not required</param>
        /// <param name="e">not required</param>
        /// <example>
        /// <code>
        ///     // Create Job Information CSV
        ///     JobInformation();
        ///     
        ///     // Create Healthcheck Report
        ///     CreateHealthcheckReport();
        ///     
        ///     // Compress Healthcheck Report
        ///     CompressHealthcheckReport();
        ///     
        ///     // Send to PM
        ///     SendReportEmail();
        /// </code>
        /// </example>
        private void CreateReport_Click(object sender, RoutedEventArgs e)
        {
            // Create Job Information CSV
            JobInformation();

            // Create Healthcheck Report
            CreateHealthcheckReport();

            // Compress Healthcheck Report
            CompressHealthcheckReport();

            // Send to PM
            SendReportEmail();
        }

        /// <summary>
        /// Enables selection of all or none of the downloaded files when the user interacts with the form.
        /// </summary>
        /// <param name="sender">not required</param>
        /// <param name="e">not required</param>
        /// <example>
        /// <code>
        ///     // Number of items in mailing file list
        ///     int files = MailingFileList.Items.Count;
        ///     int files_selected = MailingFileList.SelectedItems.Count;
        ///       
        ///     // Check if all items selected
        ///     if (files_selected == files)
        ///     {
        ///         // Deselect all files
        ///         MailingFileList.SelectedIndex = -1;
        ///     }
        ///     else
        ///     {
        ///         // Select all files
        ///         MailingFileList.SelectAll();
        ///     }
        /// </code>
        /// </example>
        private void MailingFileList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            // Number of items in mailing file list
            int files = MailingFileList.Items.Count;
            int files_selected = MailingFileList.SelectedItems.Count;
          
            // Check if all items selected
            if (files_selected == files)
            {
                // Deselect all files
                MailingFileList.SelectedIndex = -1;
            }
            else
            {
                // Select all files
                MailingFileList.SelectAll();
            }
        }

        /// <summary>
        /// Removes any characters from a filename that a user may have added during the healthcheck process in Cygnus,
        /// to provide more effective matching when creating dedupe totals.
        /// </summary>
        /// <param name="filename">accepts a filename as a string</param>
        /// <returns>the filename as a string with superfluous characters removed</returns>
        /// <example>
        /// <code>
        ///     // Trim whitespace from either side of filename
        ///     filename = filename.Trim();
        ///     
        ///     // Continue if not blank
        ///     if (String.IsNullOrWhiteSpace(filename) == false)
        ///     {
        ///         // Set loop state
        ///         bool finished = false;
        ///         int count = 0;
        ///     
        ///         // Format the filename until process is finished
        ///         do
        ///         {
        ///             // Store character positions
        ///             string first_position = filename.Substring(1, 1);
        ///             string last_position = filename.Substring((filename.Length - 1), 1);
        ///             int new_length = (filename.Length) - 1;
        ///     
        ///             // Check first position
        ///             switch (first_position)
        ///             {
        ///                 case "_":
        ///                     filename = filename.Substring(2, new_length);
        ///                     break;
        ///                 case ".":
        ///                     filename = filename.Substring(2, new_length);
        ///                     break;
        ///                 default:
        ///                     // Do nothing
        ///                     break;
        ///              }    
        ///     
        ///             // Check last position
        ///             switch (last_position)
        ///             {
        ///                 case "_":
        ///                     filename = filename.Substring(1, (new_length - 1));
        ///                     break;
        ///                 case ".":
        ///                     filename = filename.Substring(1, (new_length - 1));
        ///                     break;
        ///                 default:
        ///                     // Increment count
        ///                     count++;
        ///                     break;
        ///             }
        ///     
        ///             // Finish loop if count is 5
        ///             if (count == 5)
        ///             {
        ///                 finished = true;
        ///             }
        ///         } while (finished == false);
        ///     
        ///         // Remove unwanted characters
        ///         int new_last_position = filename.Length - 1;
        ///         string upper_filename = filename.ToUpper();
        ///         int parse_result = 0;
        ///     
        ///         // Remove cygnus splerge definition
        ///         if ((upper_filename.Substring(1,1)=="S") :amp;:amp; (upper_filename.Substring(2,1)=="P") :amp;:amp; (upper_filename.Substring(5,1)=="_")) 
        ///         {
        ///             filename = filename.Substring(6, new_last_position);
        ///         }
        ///     
        ///         // Remove w_ or w.
        ///         if ((upper_filename.Substring(1, 1) == "W") :amp;:amp; ((upper_filename.Substring(2, 1) == "_") || (upper_filename.Substring(2, 1) == ".")))
        ///         {
        ///             filename = filename.Substring(3, new_last_position);
        ///         }
        ///     
        ///         // Remove c_ or c.
        ///         if ((upper_filename.Substring(1, 1) == "C") :amp;:amp; ((upper_filename.Substring(2, 1) == "_") || (upper_filename.Substring(2, 1) == ".")))
        ///         {
        ///             filename = filename.Substring(3, new_last_position);
        ///         }
        ///     
        ///         // Remove n_ or n.
        ///         if ((int.TryParse(filename.Substring(1,1), out parse_result)) :amp;:amp; ((upper_filename.Substring(2, 1) == "_") || (upper_filename.Substring(2, 1) == ".")))
        ///         {
        ///             filename = filename.Substring(3, new_last_position);
        ///         }
        ///     
        ///         // Remove nn_ or nn.
        ///         if ((int.TryParse(filename.Substring(1, 2), out parse_result)) :amp;:amp; ((upper_filename.Substring(3, 1) == "_") || (upper_filename.Substring(3, 1) == ".")))
        ///         {
        ///             filename = filename.Substring(4, new_last_position);
        ///         }
        ///     
        ///         // Remove #n
        ///         if ((int.TryParse(filename.Substring(new_last_position, 1), out parse_result)) :amp;:amp; (upper_filename.Substring(new_last_position-1,1)=="#"))
        ///         {
        ///          filename = filename.Substring(4, new_last_position);
        ///         }
        ///     }
        ///        
        ///     return filename;
        /// </code>
        /// </example>
        private string FormatName(string filename)
        {
            // Trim whitespace from either side of filename
            filename = filename.Trim();

            // Continue if not blank
            if (String.IsNullOrWhiteSpace(filename) == false)
            {
                // Set loop state
                bool finished = false;
                int count = 0;

                // Format the filename until process is finished
                do
                {
                    // Store character positions
                    string first_position = filename.Substring(1, 1);                  
                    string last_position = filename.Substring((filename.Length - 1), 1);
                    int new_length = (filename.Length) - 1;

                    // Check first position
                    switch (first_position)
                    {
                        case "_":
                            filename = filename.Substring(2, new_length);
                            break;
                        case ".":
                            filename = filename.Substring(2, new_length);
                            break;
                        default:
                            // Do nothing
                            break;
                    }

                    // Check last position
                    switch (last_position)
                    {
                        case "_":
                            filename = filename.Substring(1, (new_length - 1));
                            break;
                        case ".":
                            filename = filename.Substring(1, (new_length - 1));
                            break;
                        default:
                            // Increment count
                            count++;
                            break;
                    }

                    // Finish loop if count is 5
                    if (count == 5)
                    {
                        finished = true;
                    }

                } while (finished == false);

                // Remove unwanted characters
                int new_last_position = filename.Length - 1;
                string upper_filename = filename.ToUpper();
                int parse_result = 0;

                // Remove cygnus splerge definition
                if ((upper_filename.Substring(1,1)=="S") && (upper_filename.Substring(2,1)=="P") && (upper_filename.Substring(5,1)=="_")) 
                {
                    filename = filename.Substring(6, new_last_position);
                }

                // Remove w_ or w.
                if ((upper_filename.Substring(1, 1) == "W") && ((upper_filename.Substring(2, 1) == "_") || (upper_filename.Substring(2, 1) == ".")))
                {
                    filename = filename.Substring(3, new_last_position);
                }

                // Remove c_ or c.
                if ((upper_filename.Substring(1, 1) == "C") && ((upper_filename.Substring(2, 1) == "_") || (upper_filename.Substring(2, 1) == ".")))
                {
                    filename = filename.Substring(3, new_last_position);
                }

                // Remove n_ or n.
                if ((int.TryParse(filename.Substring(1,1), out parse_result)) && ((upper_filename.Substring(2, 1) == "_") || (upper_filename.Substring(2, 1) == ".")))
                {
                    filename = filename.Substring(3, new_last_position);
                }

                // Remove nn_ or nn.
                if ((int.TryParse(filename.Substring(1, 2), out parse_result)) && ((upper_filename.Substring(3, 1) == "_") || (upper_filename.Substring(3, 1) == ".")))
                {
                    filename = filename.Substring(4, new_last_position);
                }

                // Remove #n
                if ((int.TryParse(filename.Substring(new_last_position, 1), out parse_result)) && (upper_filename.Substring(new_last_position-1,1)=="#"))
                {
                    filename = filename.Substring(4, new_last_position);
                }
            }
            
            return filename;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuItem_Main_Click(object sender, RoutedEventArgs e)
        {
            // Invoke Config form
            ConfigForm config = new ConfigForm();
            config.Show();
        }

        // Main Functions

        /// <summary>
        /// Calls various methods to process files for output. The download information and overseas
        /// redirects methods are called. Should dedupes exist then those methods are called here also.
        /// </summary>
        /// <example>
        /// <code>
        ///     // Animate upon confirmation
        ///     ConfirmAnimations();
        ///     
        ///     // Process download information
        ///     ProcessDownloadInformation();
        ///     
        ///     // Process overseas redirects information
        ///     OverseasRedirects();
        ///     
        ///     // Process Download Information Totals
        ///     CalculateDownloadTotals();
        ///     
        ///     // Process dedupe information [if applicable]
        ///     if (CheckDedupesExist())
        ///     {
        ///         // Clear MailingFileList
        ///         MailingFileList.Items.Clear();
        ///     
        ///         // Retrieve dedupe information
        ///         RetrieveDedupeInformation();
        ///     
        ///         // Dedupe Animations
        ///         DedupeAnimations();
        ///     
        ///     }
        ///     else
        ///     {
        ///         // Clear MailingFileList
        ///         MailingFileList.Items.Clear();
        ///     
        ///         // Summary Animations
        ///         SummaryAnimations();
        ///     }
        /// </code>
        /// </example>
        private void MainFileProcessing()
        {
            // Animate upon confirmation
            ConfirmAnimations();

            // Process download information
            ProcessDownloadInformation();

            // Process overseas redirects information
            OverseasRedirects();

            // Process Download Information Totals
            CalculateDownloadTotals();

            // Process dedupe information [if applicable]
            if (CheckDedupesExist())
            {
                // Clear MailingFileList
                MailingFileList.Items.Clear();

                // Retrieve dedupe information
                RetrieveDedupeInformation();

                // Dedupe Animations
                DedupeAnimations();

            }
            else
            {
                // Clear MailingFileList
                MailingFileList.Items.Clear();

                // Summary Animations
                SummaryAnimations();
            }
        }

        /// <summary>
        /// Updates the selected job if the job has been selected from the recently reported
        /// job list.
        /// </summary>
        /// <param name="sender">not required</param>
        /// <param name="e">not required</param>
        /// <example>
        /// <code>
        ///     JobSelect.SelectedItem = RecentJobList.SelectedItem;
        /// </code>
        /// </example>
        private void RecentJobList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            JobSelect.SelectedItem = RecentJobList.SelectedItem;
        }

        /// <summary>
        /// Updates the form with a list of available jobs, and adds them to the drop down
        /// on the form
        /// </summary>
        /// <example>
        /// <code>
        ///     // Refresh the list of available jobs on W and adds them to the job list
        ///     string job_directory;
        ///     string[] jobs;
        ///     
        ///     // Initialise job directory
        ///     job_directory = "Q:\\Cygnus DP\\Jobs\\";
        ///     
        ///     // Retrieve list of directories
        ///     jobs = Directory.GetDirectories(job_directory);
        ///     Array.Sort(jobs);
        ///     Array.Reverse(jobs);
        ///     
        ///     // For each job folder in Q populate list
        ///     foreach (var job in jobs)
        ///     {
        ///         string job_name = job.Replace("Q:\\Cygnus DP\\Jobs\\", "");
        ///         double result;
        ///     
        ///         // Check the first five characters are numerical nnnnn
        ///         if (Double.TryParse(job_name.Substring(0,5), out result))
        ///         {
        ///             JobSelect.Items.Add(job_name);
        ///         }
        ///     }
        /// </code>
        /// </example>
        private void RefreshJobList()
        {
            // Refresh the list of available jobs and adds them to the job list
            string job_directory;
            string[] jobs;

            // Initialise job directory
            job_directory = ReadInDirectories()[1];

            // Retrieve list of directories
            jobs = Directory.GetDirectories(job_directory);
            Array.Sort(jobs);
            Array.Reverse(jobs);

            // For each job folder in Q populate list
            foreach (var job in jobs)
            {
                string job_name = job.Replace(ReadInDirectories()[1], "");
                double result;

                // Check the first five characters are numerical nnnnn
                try
                {
                    if (Double.TryParse(job_name.Substring(0, 5), out result))
                    {
                        if (job_name.Substring(0, 4) != "9999")
                        {
                            JobSelect.Items.Add(job_name);
                        }
                        
                    }
                }
                catch (Exception ex)
                {

                }          
            }
        }

        /// <summary>
        /// Retrieves the ten most recent jobs that have had the healthcheck report created. And adds
        /// to the form.
        /// </summary>
        /// <example>
        /// <code>
        ///     // Refreshes the list of recently reported jobs
        ///     string job_directory;
        ///     string[] jobs;
        ///     int job_count = 0;
        ///     
        ///     // Initialise job directory
        ///     job_directory = "W:\\Client Data\\";
        ///     
        ///     // Retrieve list of directories
        ///     jobs = Directory.GetDirectories(job_directory);
        ///     Array.Sort(jobs);
        ///     Array.Reverse(jobs);
        ///     
        ///     // For each job folder in W populate list
        ///     foreach (var job in jobs)
        ///     {
        ///         string job_name = job.Replace("W:\\Client Data\\", "");
        ///         string file_string;
        ///         double result;
        ///     
        ///         // Validate job name
        ///         if (job_name.Length > 4)
        ///         {
        ///             if (Double.TryParse(job_name.Substring(0, 5), out result))
        ///             {
        ///                 file_string = job + "\\Project Management\\Data Health Check\\Report\\Healthcheck Report.pdf";
        ///     
        ///                 // Increment count if it exists
        ///                 if (File.Exists(file_string))
        ///                 {
        ///                     job_count++;
        ///                 }
        ///             }
        ///         }
        ///     }
        ///     
        ///     // Initialise array
        ///     string[] reported_jobs = new string[job_count];
        ///     int count = 0;
        ///     
        ///     // Store in array
        ///     // For each job folder in W populate list
        ///     foreach (var job in jobs)
        ///     {
        ///         string job_name = job.Replace("W:\\Client Data\\", "");
        ///         string file_string;
        ///         double result;
        ///     
        ///         // Perform validation
        ///         if (job_name.Length > 4)
        ///         {
        ///             if (Double.TryParse(job_name.Substring(0, 5), out result))
        ///             {
        ///                 file_string = job + "\\Project Management\\Data Health Check\\Report\\Healthcheck Report.pdf";
        ///     
        ///                 if (File.Exists(file_string))
        ///                 {
        ///                     // Add file string to reported jobs
        ///                     reported_jobs[count] = file_string;
        ///                     count++;
        ///                 }
        ///             }
        ///         }
        ///     }
        ///     
        ///     // Create array to store time pdf report was created
        ///     DateTime[] creation_times = new DateTime[reported_jobs.Length];
        ///     
        ///     for (int i = 0; i &lt; reported_jobs.Length; i++)
        ///     {
        ///         // Store current creation time in array
        ///         creation_times[i] = Directory.GetLastWriteTime(reported_jobs[i]);
        ///     }
        ///     
        ///     // Sort and then reverse order
        ///     Array.Sort(creation_times, reported_jobs);
        ///     Array.Reverse(reported_jobs);
        ///     
        ///     // Loop through reported jobs
        ///     // Max of ten
        ///     int max = 10;
        ///     int iter = 0;
        ///     
        ///     foreach (var name in reported_jobs)
        ///     {
        ///         // Add to recently reported 
        ///         if (iter &lt;= max)
        ///         {
        ///             string job = name.Replace("W:\\Client Data\\", "");
        ///             job = job.Replace("\\Project Management\\Data Health Check\\Report\\Healthcheck Report.pdf", "");
        ///     
        ///             RecentJobList.Items.Add(job);
        ///         }     
        ///         iter++;    
        ///     }
        /// </code>
        /// </example>
        private void RecentlyReportedJobs()
        {
            // Refreshes the list of recently reported jobs
            string job_directory;
            string[] jobs;
            int job_count = 0;

            // Initialise job directory
            job_directory = ReadInDirectories()[0];

            // Retrieve list of directories
            jobs = Directory.GetDirectories(job_directory);
            Array.Sort(jobs);
            Array.Reverse(jobs);

            // For each job folder in W populate list
            foreach (var job in jobs)
            {
                string job_name = job.Replace(ReadInDirectories()[0], "");
                string file_string;
                double result;

                // Validate job name
                if (job_name.Length > 4)
                {
                    if (Double.TryParse(job_name.Substring(0, 5), out result))
                    {

                        file_string = job + "\\Project Management\\Data Health Check\\Report\\Healthcheck Report.pdf";

                        // Increment count if it exists
                        if (File.Exists(file_string))
                        {
                            job_count++;
                        }
                    }   
                }             
            }

            // Initialise array
            string[] reported_jobs = new string[job_count];
            int count = 0;

            // Store in array
            // For each job folder in W populate list
            foreach (var job in jobs)
            {
                string job_name = job.Replace(ReadInDirectories()[0], "");
                string file_string;
                double result;

                // Perform validation
                if (job_name.Length > 4)
                {
                    if (Double.TryParse(job_name.Substring(0, 5), out result))
                    {

                        file_string = job + "\\Project Management\\Data Health Check\\Report\\Healthcheck Report.pdf";

                        if (File.Exists(file_string))
                        {
                            // Add file string to reported jobs
                            reported_jobs[count] = file_string;
                            count++;
                        }
                    }
                } 
            }

            // Create array to store time pdf report was created
            DateTime[] creation_times = new DateTime[reported_jobs.Length];

            for (int i = 0; i < reported_jobs.Length; i++)
            {
                // Store current creation time in array
                creation_times[i] = Directory.GetLastWriteTime(reported_jobs[i]);
            }

            // Sort and then reverse order
            Array.Sort(creation_times, reported_jobs);
            Array.Reverse(reported_jobs);

            // Loop through reported jobs
            // Max of ten
            int max = 10;
            int iter = 0;

            foreach (var name in reported_jobs)
            {
                // Add to recently reported 
                if (iter <= max)
                {
                    string job = name.Replace(ReadInDirectories()[0], "");
                    job = job.Replace("\\Project Management\\Data Health Check\\Report\\Healthcheck Report.pdf", "");
                    
                    RecentJobList.Items.Add(job);
                }
                iter++;
            }
        }

        /// <summary>
        /// Reads the _Merged.xml file and displays on the form, and downloaded files that
        /// exist on the Cygnus workflow
        /// </summary>
        /// <example>
        /// <code>
        ///     // Declare variables
        ///     string xml_path;
        ///     bool proceed;
        ///     XmlTextReader xml_reader;
        ///     
        ///     // Intialise variables
        ///     xml_path = MergedXMLPath();
        ///     
        ///     // Clear current information from the list of mailing files
        ///     MailingFileList.Items.Clear();
        ///     
        ///     // Check XML for Errors using XDocument
        ///     try //loading XML
        ///     {
        ///         xml_reader = new XmlTextReader(xml_path);
        ///     
        ///         while (xml_reader.Read())
        ///         {
        ///             // Read through file
        ///         }
        ///     
        ///         xml_reader.Close();
        ///     
        ///         proceed = true;
        ///     }
        ///     catch (Exception e) // Upon error...
        ///     {
        ///         // XML did not load correctly
        ///         MessageBox.Show(e.Message);
        ///         proceed = false;
        ///     }
        ///     
        ///     // If xml can be read then proceed
        ///     if (proceed)
        ///     {
        ///         // Run initial screen animation
        ///         RunAnimations();
        ///     
        ///         // Open xml
        ///         xml_reader = new XmlTextReader(xml_path);
        ///             
        ///         // Read each node in xml
        ///         while(xml_reader.Read())
        ///         {
        ///             // if node is a process node
        ///             if(xml_reader.Name == "process")
        ///             {
        ///                 // look through process attributes
        ///                 while (xml_reader.MoveToNextAttribute())
        ///                 {
        ///                     // if attribute value is download
        ///                     if (xml_reader.Value == "download")
        ///                     {
        ///                         // move to next attribute (next attribute is always name..)
        ///                         xml_reader.MoveToNextAttribute();
        ///     
        ///                         // Print attribute name to GUI
        ///                         MailingFileList.Items.Add(xml_reader.Value);
        ///     
        ///                     }
        ///                 }
        ///              }
        ///          }
        ///     }
        /// </code>
        /// </example>
        private void FilesOnWorkflow()
        {

            // Declare variables
            string xml_path;
            bool proceed;
            XmlTextReader xml_reader;

            // Intialise variables
            xml_path = MergedXMLPath();

            // Clear current information from the list of mailing files
            MailingFileList.Items.Clear();

            // Check XML for Errors using XDocument
            try //loading XML
            {
                xml_reader = new XmlTextReader(xml_path);

                while (xml_reader.Read())
                {
                    // Read through file
                }

                xml_reader.Close();
                            
                proceed = true;
            }
            catch (Exception e) // Upon error...
            {
                // XML did not load correctly
                MessageBox.Show(e.Message);
                proceed = false;
            }

            // If xml can be read then proceed
            if (proceed)
            {
                // Run initial screen animation
                RunAnimations();
                
                // Open xml
                xml_reader = new XmlTextReader(xml_path);
                
                // Read each node in xml
                while(xml_reader.Read())
                {
                    // if node is a process node
                    if(xml_reader.Name == "process")
                    {
                        // look through process attributes
                        while (xml_reader.MoveToNextAttribute())
                        {
                            // if attribute value is download
                            if (xml_reader.Value == "download")
                            {
                                // move to next attribute (next attribute is always name..)
                                xml_reader.MoveToNextAttribute();

                                // Print attribute name to GUI
                                MailingFileList.Items.Add(xml_reader.Value);

                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Creates a copy of the _Merged.xml file, and removes unecessary information from it, like character, translational or
        /// post-suppression dedupe information, that could cause reading problems.
        /// </summary>
        /// <example>
        /// <code>
        ///     // Set paths
        ///     string xmlFile = MergedXMLPath();
        ///     string cleaningFile = CygnusJobDirectory() + "toclean.txt";
        ///     string cleanedFileTrans = CygnusJobDirectory() + "cleanedTrans.txt";
        ///     string cleanedFileSupps = CygnusJobDirectory() + "cleanedSupps.txt";
        ///     string currentLine;
        ///     string deletionFlag = "";
        ///     
        ///     // Inform user of cleaning operation
        ///     UserInformation.Content = "Cleaning _Merged.xml file.";
        ///     
        ///     // Create copy of merged xml as the file to clean, overwrite existing
        ///     File.Copy(xmlFile, cleaningFile, true);
        ///     
        ///     // Remove translational data
        ///     using (StreamReader reader = new StreamReader(cleaningFile))
        ///     {
        ///         using (StreamWriter writer = new StreamWriter(cleanedFileTrans))
        ///         {
        ///             // Read through file to be cleaned
        ///             while (!reader.EndOfStream)
        ///             {
        ///                 // Store current line
        ///                 currentLine = reader.ReadLine();
        ///     
        ///                 // If the line contains the opening translation xml tag
        ///                 if (currentLine.Contains("&lt;translations&gt;"))
        ///                 {
        ///                     // Set deletion to on and ignore (delete) line
        ///                     deletionFlag = "1";
        ///                 }
        ///     
        ///                 // If the line deletion flag is on and the line does not contain the closing tag
        ///                 if (deletionFlag == "1" &amp;&amp; !currentLine.Contains("&lt;/translations&gt;"))
        ///                 {
        ///                     // ignore (delete) line
        ///                 }
        ///     
        ///                 // If the line deletion flag is on and the line contains the closing tag
        ///                 if (deletionFlag == "1" &amp;&amp; currentLine.Contains("&lt;/translations&gt;"))
        ///                 {
        ///                     // ignore (delete) line and reset
        ///                     deletionFlag = "";
        ///                 }
        ///     
        ///                 // If the line deletion flag is off and the line does not contain the closing tag
        ///                 if (deletionFlag == "" &amp;&amp; !currentLine.Contains("&lt;/translations&gt;"))
        ///                 {
        ///                     // Remove erroneous characters
        ///                     currentLine = currentLine.Replace("&amp;amp;","");
        ///                     currentLine = currentLine.Replace("&amp;", "");
        ///                     currentLine = currentLine.Replace("\0", "");
        ///     
        ///                     // Write line to cleaned file
        ///                     writer.WriteLine(currentLine);
        ///                 }
        ///             }
        ///         }
        ///     }
        ///     
        ///     // Reset flags for next run
        ///     deletionFlag = "";
        ///     currentLine = "";
        ///     
        ///     // Remove post-suppression dedupe data
        ///     using (StreamReader reader = new StreamReader(cleanedFileTrans))
        ///     {
        ///         using (StreamWriter writer = new StreamWriter(cleanedFileSupps))
        ///         {
        ///             // Read through recently cleaned file
        ///             while (!reader.EndOfStream)
        ///             {
        ///                 // Retrieve current line, and hold in memory
        ///                 currentLine = reader.ReadLine();
        ///     
        ///                 // If the line contains Trace 9 module info then output
        ///                 if (currentLine.Contains("Trace 9"))
        ///                 {
        ///                     writer.WriteLine(currentLine);
        ///                 }
        ///                 else
        ///                 {
        ///                     // If line is the first in the post suppression dedupe section then ignore (delete) and set deletion flag
        ///                     if (currentLine.Contains("Post-Suppression Dedupe") &amp;&amp; currentLine.Contains("overview"))
        ///                     {
        ///                         deletionFlag = "1";
        ///                     }
        ///     
        ///                     // If deletion flag is set and the current line does not contain the closing tag
        ///                     if (deletionFlag == "1" &amp;&amp; !currentLine.Contains("&lt;/overview&gt;"))
        ///                     {
        ///                         // Ignore line (delete)
        ///                     }
        ///     
        ///                     // If deletion flag is set and the current line contains the closing tag
        ///                     if (deletionFlag == "1" &amp;&amp; currentLine.Contains("&lt;/overview&gt;"))
        ///                     {
        ///                         // Reset the deletion flag
        ///                         deletionFlag = "";
        ///                     }
        ///     
        ///                     // If the deletion flag is not set and the current line does not contain the closing tag
        ///                     if (deletionFlag == "" &amp;&amp; !currentLine.Contains("&lt;/overview&gt;"))
        ///                     {
        ///                         // Write line
        ///                         writer.WriteLine(currentLine);
        ///                     }
        ///                 }
        ///             }
        ///         }
        ///     }
        ///     
        ///     // Perform backup operations and delete files
        ///     File.Copy(xmlFile, CygnusJobDirectory() + "old_Merged.xml", true);
        ///     File.Delete(xmlFile);
        ///     File.Copy(cleanedFileSupps, xmlFile);
        ///     File.Delete(cleanedFileSupps);
        ///     File.Delete(cleanedFileTrans);
        ///     
        ///     UserInformation.Content = "_Merged.xml file cleaned.";
        /// </code>
        /// </example>
        private void MergedXMLCleaner()
        {
            // Set paths
            string xmlFile = MergedXMLPath();
            string cleaningFile = CygnusJobDirectory() + "toclean.txt";
            string cleanedFileTrans = CygnusJobDirectory() + "cleanedTrans.txt";
            string cleanedFileSupps = CygnusJobDirectory() + "cleanedSupps.txt";
            string currentLine;
            string deletionFlag = "";

            // Inform user of cleaning operation
            UserInformation.Content = "Cleaning _Merged.xml file.";

            // Create copy of merged xml as the file to clean, overwrite existing
            File.Copy(xmlFile, cleaningFile, true);

            // Remove translational data
            using (StreamReader reader = new StreamReader(cleaningFile))
            {
                using (StreamWriter writer = new StreamWriter(cleanedFileTrans))
                {
                    // Read through file to be cleaned
                    while (!reader.EndOfStream)
                    {
                        // Store current line
                        currentLine = reader.ReadLine();

                        // If the line contains the opening translation xml tag
                        if (currentLine.Contains("<translations>"))
                        {
                            // Set deletion to on and ignore (delete) line
                            deletionFlag = "1";
                        }

                        // If the line deletion flag is on and the line does not contain the closing tag
                        if (deletionFlag == "1" && !currentLine.Contains("</translations>"))
                        {
                            // ignore (delete) line
                        }

                        // If the line deletion flag is on and the line contains the closing tag
                        if (deletionFlag == "1" && currentLine.Contains("</translations>"))
                        {
                            // ignore (delete) line and reset
                            deletionFlag = "";
                        }

                        // If the line deletion flag is off and the line does not contain the closing tag
                        if (deletionFlag == "" && !currentLine.Contains("</translations>"))
                        {
                            // Remove erroneous characters
                            currentLine = currentLine.Replace("&amp;","");
                            currentLine = currentLine.Replace("&", "");
                            currentLine = currentLine.Replace("\0", "");

                            // Write line to cleaned file
                            writer.WriteLine(currentLine);
                        }
                    }
                }
            }

            // Reset flags for next run
            deletionFlag = "";
            currentLine = "";

            // Remove post-suppression dedupe data
            using (StreamReader reader = new StreamReader(cleanedFileTrans))
            {
                using (StreamWriter writer = new StreamWriter(cleanedFileSupps))
                {
                    // Read through recently cleaned file
                    while (!reader.EndOfStream)
                    {
                        // Retrieve current line, and hold in memory
                        currentLine = reader.ReadLine();

                        // If the line contains Trace 9 module info then output
                        if (currentLine.Contains("Trace 9"))
                        {
                            writer.WriteLine(currentLine);
                        }
                        else
                        {
                            // If line is the first in the post suppression dedupe section then ignore (delete) and set deletion flag
                            if (currentLine.Contains("Post-Suppression Dedupe") && currentLine.Contains("overview"))
                            {
                                deletionFlag = "1";
                            }

                            // If deletion flag is set and the current line does not contain the closing tag
                            if (deletionFlag == "1" && !currentLine.Contains("</overview>"))
                            {
                                // Ignore line (delete)
                            }

                            // If deletion flag is set and the current line contains the closing tag
                            if (deletionFlag == "1" && currentLine.Contains("</overview>"))
                            {
                                // Reset the deletion flag
                                deletionFlag = "";
                            }

                            // If the deletion flag is not set and the current line does not contain the closing tag
                            if (deletionFlag == "" && !currentLine.Contains("</overview>"))
                            {
                                // Write line
                                writer.WriteLine(currentLine);
                            }
                        }
                    }
                }
            }

            // Perform backup operations and delete files
            File.Copy(xmlFile, CygnusJobDirectory() + "old_Merged.xml", true);
            File.Delete(xmlFile);
            File.Copy(cleanedFileSupps, xmlFile);
            File.Delete(cleanedFileSupps);
            File.Delete(cleanedFileTrans);

            UserInformation.Content = "_Merged.xml file cleaned.";
        }

        // Download Functions

        /// <summary>
        /// Reads the _Merged.xml to determine if any downloaded files exist on the Cygnus workflow
        /// </summary>
        /// <returns>true or false depending on the existence of downloaded files</returns>
        /// <example>
        /// <code>
        ///     // Open _Merged.XML
        ///     XmlDocument xml_document = new XmlDocument();
        ///     xml_document.Load(MergedXMLPath());
        ///     bool downloads_exist = false;
        ///     
        ///     // Retrieve process node
        ///     XmlNodeList process_list = xml_document.DocumentElement.SelectNodes("/cygnus_reports/process");
        ///     
        ///     // Loop through each process
        ///     foreach (XmlNode process in process_list)
        ///     {
        ///         // Store process attributes
        ///         XmlAttributeCollection process_attributes = process.Attributes;
        ///         string process_type = "";
        ///         string process_name = "";
        ///     
        ///         // loop through attributes
        ///         foreach (XmlAttribute attribute in process_attributes)
        ///         {
        ///                 // Check attribute field
        ///                 switch (attribute.Name)
        ///                 {
        ///                     // if 'type' store contents (i.e. download)
        ///                     case "type":
        ///     process_type = attribute.Value;
        ///                         break;
        ///                     // If 'name' store contents (i.e. download name)
        ///                     case "name":
        ///     process_name = attribute.Value;
        ///                         break;
        ///     }
        ///     }
        ///     
        ///             // Check process is dedupe and the name is not a suppression module or post-suppression dedupe
        ///             if (process_type.ToUpper() == "DOWNLOAD")
        ///             {
        ///     downloads_exist = true;
        ///     }
        ///     }
        ///     
        ///     // If there are valid dedupes then return true
        ///     return downloads_exist;
        /// </code>
        /// </example>
        private bool CheckDownloadsExist()
        {
            // Open _Merged.XML
            XmlDocument xml_document = new XmlDocument();
            xml_document.Load(MergedXMLPath());
            bool downloads_exist = false;

            // Retrieve process node
            XmlNodeList process_list = xml_document.DocumentElement.SelectNodes("/cygnus_reports/process");

            // Loop through each process
            foreach (XmlNode process in process_list)
            {
                // Store process attributes
                XmlAttributeCollection process_attributes = process.Attributes;
                string process_type = "";
                string process_name = "";

                // loop through attributes
                foreach (XmlAttribute attribute in process_attributes)
                {
                    // Check attribute field
                    switch (attribute.Name)
                    {
                        // if 'type' store contents (i.e. download)
                        case "type":
                            process_type = attribute.Value;
                            break;
                        // If 'name' store contents (i.e. download name)
                        case "name":
                            process_name = attribute.Value;
                            break;
                    }
                }

                // Check process is dedupe and the name is not a suppression module or post-suppression dedupe
                if (process_type.ToUpper() == "DOWNLOAD")
                {
                    downloads_exist = true;
                }
            }

            // If there are valid dedupes then return true
            return downloads_exist;
        }

        /// <summary>
        /// Clears the output directory of any pre-existing download files, and proceeds to retrieve
        /// the download statistics, from the _Merged.xml file, of each file the user has selected on the form.
        /// This is performed by finding the matching process in the _Merged.xml file and passing the process
        /// to two methods that export the statistics and file structures respectively.
        /// </summary>
        /// <example>
        /// <code>
        ///     // Declare variables
        ///     fileCount = MailingFileList.SelectedItems.Count;
        ///     string[] file_list = new string[fileCount];
        ///     int count = 0;
        ///     
        ///     // Delete existing download information - TEST DIRECTORY
        ///     File.Delete(OutputDirectory() + "download_stats.csv");
        ///     File.Delete(OutputDirectory() + "file_structures.csv");
        ///     File.Delete(OutputDirectory() + "overseas_redirects.csv");
        ///     
        ///     // Populate file list array
        ///     foreach (var file in MailingFileList.SelectedItems)
        ///     {
        ///         // Add file to array
        ///         file_list[count] = file.ToString();
        ///     
        ///         // Iterate count
        ///         count++;
        ///     }
        ///     
        ///     // For each file in array, retrieve the matching stats from the XML file
        ///     foreach (var mailing_file in file_list)
        ///     {
        ///         // Open _Merged.XML
        ///         XmlDocument xml_document = new XmlDocument();
        ///         xml_document.Load(MergedXMLPath());
        ///     
        ///         // Retrieve process node
        ///         XmlNodeList process_list = xml_document.DocumentElement.SelectNodes("/cygnus_reports/process");
        ///     
        ///         // Loop through each process
        ///         foreach (XmlNode process in process_list)
        ///         {
        ///             // Store process attributes
        ///             XmlAttributeCollection process_attributes = process.Attributes;
        ///             string process_type = "";
        ///             string process_name = "";
        ///                 
        ///             // loop through attributes
        ///             foreach (XmlAttribute attribute in process_attributes)
        ///             {
        ///                 // Check attribute field
        ///                 switch (attribute.Name)
        ///                 {
        ///                     // if 'type' store contents (i.e. download)
        ///                     case "type":
        ///                         process_type = attribute.Value;
        ///                         break;
        ///                     // If 'name' store contents (i.e. the filename)
        ///                     case "name":
        ///                         process_name = attribute.Value;
        ///                         break;
        ///                 }
        ///             }
        ///     
        ///             // Check process is download and the name matches selected filename
        ///             if ((process_type == "download") &amp;&amp; (process_name == mailing_file))
        ///             {
        ///                 // Export download statistics to CSV
        ///                 ExportDownloadStatistics(mailing_file, process); // selected mailing file, current xml process information
        ///     
        ///                 // Export file structures to CSV
        ///                 ExportFileStructures(mailing_file, process); // selected mailing file, current xml process information
        ///     
        ///             }
        ///         }
        ///     }
        /// </code>
        /// </example>
        private void ProcessDownloadInformation()
        {
            // Declare variables
            fileCount = MailingFileList.SelectedItems.Count;
            string[] file_list = new string[fileCount];
            int count = 0;

            // Delete existing download information - TEST DIRECTORY
            File.Delete(OutputDirectory() + "download_stats.csv");
            File.Delete(OutputDirectory() + "file_structures.csv");
            File.Delete(OutputDirectory() + "overseas_redirects.csv");

            // Populate file list array
            foreach (var file in MailingFileList.SelectedItems)
            {
                // Add file to array
                file_list[count] = file.ToString();

                // Iterate count
                count++;
            }

            // For each file in array, retrieve the matching stats from the XML file
            foreach (var mailing_file in file_list)
            {
                // Open _Merged.XML
                XmlDocument xml_document = new XmlDocument();
                xml_document.Load(MergedXMLPath());
                
                // Retrieve process node
                XmlNodeList process_list = xml_document.DocumentElement.SelectNodes("/cygnus_reports/process");

                // Loop through each process
                foreach (XmlNode process in process_list)
                {
                    // Store process attributes
                    XmlAttributeCollection process_attributes = process.Attributes;
                    string process_type = "";
                    string process_name = "";
                    
                    // loop through attributes
                    foreach (XmlAttribute attribute in process_attributes)
                    {
                        // Check attribute field
                        switch (attribute.Name)
                        {
                            // if 'type' store contents (i.e. download)
                            case "type":
                                process_type = attribute.Value;
                                break;
                            // If 'name' store contents (i.e. the filename)
                            case "name":
                                process_name = attribute.Value;
                                break;
                        }
                    }

                    // Check process is download and the name matches selected filename
                    if ((process_type == "download") && (process_name == mailing_file))
                    {
                        // Export download statistics to CSV
                        ExportDownloadStatistics(mailing_file, process); // selected mailing file, current xml process information

                        // Export file structures to CSV
                        ExportFileStructures(mailing_file, process); // selected mailing file, current xml process information

                    }
                }
            }

        }

        /// <summary>
        /// Retrieves the download statistics for the supplied mailing/download file from the accompanying XmlNode,
        /// and outputs to file.
        /// </summary>
        /// <param name="mailing_file">the name of the download file.</param>
        /// <param name="process">the XmlNode linked to the download file process.</param>
        /// <example><code>
        ///     // Create download information array
        ///     String[] downloadStatistics;
        ///     downloadStatistics = new String[43];
        ///     
        ///     // Store filename in Array
        ///     downloadStatistics[0] = mailing_file;
        ///     
        ///     // Retrieve statistic node
        ///     XmlNode statistic_node = process.SelectSingleNode("stats");
        ///     XmlNodeList statistics_list = statistic_node.ChildNodes;
        ///     
        ///     // Retrieve contained statistic nodes
        ///     XmlNodeList volume = statistics_list.Item(0).ChildNodes;
        ///     XmlNodeList general = statistics_list.Item(1).ChildNodes;
        ///     XmlNodeList name = statistics_list.Item(2).ChildNodes;
        ///     XmlNodeList address = statistics_list.Item(3).ChildNodes;
        ///     
        ///     // Store general statistics
        ///     downloadStatistics[1] = volume.Item(0).InnerText; // Input Records
        ///     downloadStatistics[2] = volume.Item(1).InnerText; // Split Records
        ///     downloadStatistics[3] = volume.Item(2).InnerText; // Output Records
        ///     downloadStatistics[4] = general.Item(0).InnerText; // Obscenities
        ///     downloadStatistics[5] = general.Item(1).InnerText; // Diacritics
        ///     downloadStatistics[6] = general.Item(2).InnerText; // Preparsed
        ///     downloadStatistics[7] = general.Item(3).InnerText; // Jumbled
        ///     downloadStatistics[8] = general.Item(4).InnerText; // Suspect
        ///     downloadStatistics[9] = general.Item(5).InnerText; // Notables
        ///     
        ///     // Declare name children
        ///     XmlNodeList patterns_child = name.Item(0).ChildNodes;
        ///     XmlNodeList defsalutations_child = name.Item(1).ChildNodes;
        ///     XmlNodeList gen_child = name.Item(2).ChildNodes;
        ///     
        ///     XmlNodeList recognised_child = patterns_child.Item(0).ChildNodes;
        ///     XmlNodeList unrecognised_child = patterns_child.Item(1).ChildNodes;
        ///     
        ///     // Populate pattern information
        ///     downloadStatistics[10] = recognised_child.Item(0).InnerText; // male names
        ///     downloadStatistics[11] = recognised_child.Item(1).InnerText; // female names
        ///     downloadStatistics[12] = recognised_child.Item(2).InnerText; // neutral names
        ///     downloadStatistics[13] = recognised_child.Item(3).InnerText; // joint names
        ///     downloadStatistics[14] = recognised_child.Item(4).InnerText; // blank names
        ///     downloadStatistics[15] = recognised_child.Item(5).InnerText; // default names
        ///     
        ///     downloadStatistics[16] = unrecognised_child.Item(0).InnerText; // single names
        ///     downloadStatistics[17] = unrecognised_child.Item(1).InnerText; // other names
        ///     
        ///     // Populate defsalutation information
        ///     downloadStatistics[18] = defsalutations_child.Item(0).InnerText; // male sal
        ///     downloadStatistics[19] = defsalutations_child.Item(1).InnerText; // fem sal
        ///     downloadStatistics[20] = defsalutations_child.Item(2).InnerText; // neutral sal
        ///     downloadStatistics[21] = defsalutations_child.Item(3).InnerText; // joint sal
        ///     downloadStatistics[22] = defsalutations_child.Item(4).InnerText; // blank sal
        ///     downloadStatistics[23] = defsalutations_child.Item(5).InnerText; // def sal
        ///     downloadStatistics[24] = defsalutations_child.Item(6).InnerText; // unrec sal
        ///     
        ///     // Populate general name info
        ///     downloadStatistics[25] = gen_child.Item(0).InnerText; // poss deceased
        ///     downloadStatistics[26] = gen_child.Item(1).InnerText; // has affix
        ///     
        ///     // ------- Address information
        ///     XmlNodeList country_child = address.Item(0).ChildNodes;
        ///     XmlNodeList gen_address_child = address.Item(1).ChildNodes;
        ///     XmlNodeList uk_child = address.Item(2).ChildNodes;
        ///     
        ///     // Children
        ///     XmlNodeList postcode_child = uk_child.Item(0).ChildNodes;
        ///     XmlNodeList selcode_child = uk_child.Item(1).ChildNodes;
        ///     XmlNodeList dps_child = uk_child.Item(2).ChildNodes;
        ///     
        ///     // Populate country information
        ///     downloadStatistics[27] = country_child.Item(0).InnerText; // uk identified
        ///     downloadStatistics[28] = country_child.Item(1).InnerText; // uk assumed
        ///     downloadStatistics[29] = country_child.Item(2).InnerText; // overseas
        ///     downloadStatistics[30] = country_child.Item(3).InnerText; // country blank
        ///     
        ///     // Populate general address information
        ///     downloadStatistics[31] = gen_address_child.Item(0).InnerText; // business address
        ///     downloadStatistics[32] = gen_address_child.Item(1).InnerText; // residential
        ///     downloadStatistics[33] = gen_address_child.Item(2).InnerText; // suspect address
        ///     downloadStatistics[34] = gen_address_child.Item(3).InnerText; // bfpo address
        ///     
        ///     // Populate uk information
        ///     downloadStatistics[35] = postcode_child.Item(0).InnerText; // full zip
        ///     downloadStatistics[36] = postcode_child.Item(1).InnerText; // outbound only
        ///     downloadStatistics[37] = postcode_child.Item(2).InnerText; // no zip
        ///     downloadStatistics[38] = postcode_child.Item(3).InnerText; // invalid zip
        ///     
        ///     downloadStatistics[39] = selcode_child.Item(0).InnerText; // Dps direct
        ///     downloadStatistics[40] = selcode_child.Item(1).InnerText; // dps residue
        ///     
        ///     downloadStatistics[41] = dps_child.Item(0).InnerText; // dps present
        ///     downloadStatistics[42] = dps_child.Item(1).InnerText; // dps default
        ///     
        ///     // Add download statistics to CSV
        ///     StreamWriter csv_output;
        ///     string concat_statistics = "";
        ///     
        ///     if (File.Exists((OutputDirectory() + "download_stats.csv")))
        ///     {
        ///         csv_output = new StreamWriter(new FileStream(OutputDirectory() + "download_stats.csv", FileMode.Append, FileAccess.Write, FileShare.Read));
        ///     
        ///         // Build line from array
        ///         foreach (var index in downloadStatistics)
        ///         {
        ///             concat_statistics = concat_statistics + index + ",";
        ///         }
        ///     
        ///         // Write line to the csv file
        ///         csv_output.WriteLine(concat_statistics);
        ///     
        ///         // Close csv
        ///         csv_output.Close();
        ///     
        ///     }
        ///     else
        ///     {
        ///         csv_output = new StreamWriter(new FileStream(OutputDirectory() + "download_stats.csv", FileMode.Append, FileAccess.Write, FileShare.Read));
        ///     
        ///         // Build line from array
        ///         foreach (var index in downloadStatistics)
        ///         {
        ///             concat_statistics = concat_statistics + index + ",";
        ///         }
        ///     
        ///         // Retrieve headings
        ///         FileHeadings headings = new FileHeadings();
        ///     
        ///         // Write line to the csv file
        ///         csv_output.WriteLine(headings.GetDownloadStatsHeadings()); // Writing field names
        ///         csv_output.WriteLine(concat_statistics); // Writing file information
        ///     
        ///         // Close csv
        ///         csv_output.Close();
        ///     }  
        ///  </code></example>
        private void ExportDownloadStatistics(string mailing_file, XmlNode process)
        {
            // Create download information array
            String[] downloadStatistics;
            downloadStatistics = new String[43];

            // Store filename in Array
            downloadStatistics[0] = mailing_file;

            // Retrieve statistic node
            XmlNode statistic_node = process.SelectSingleNode("stats");
            XmlNodeList statistics_list = statistic_node.ChildNodes;

            // Retrieve contained statistic nodes
            XmlNodeList volume = statistics_list.Item(0).ChildNodes;
            XmlNodeList general = statistics_list.Item(1).ChildNodes;
            XmlNodeList name = statistics_list.Item(2).ChildNodes;
            XmlNodeList address = statistics_list.Item(3).ChildNodes;

            // Store general statistics
            downloadStatistics[1] = volume.Item(0).InnerText; // Input Records
            downloadStatistics[2] = volume.Item(1).InnerText; // Split Records
            downloadStatistics[3] = volume.Item(2).InnerText; // Output Records
            downloadStatistics[4] = general.Item(0).InnerText; // Obscenities
            downloadStatistics[5] = general.Item(1).InnerText; // Diacritics
            downloadStatistics[6] = general.Item(2).InnerText; // Preparsed
            downloadStatistics[7] = general.Item(3).InnerText; // Jumbled
            downloadStatistics[8] = general.Item(4).InnerText; // Suspect
            downloadStatistics[9] = general.Item(5).InnerText; // Notables

            // Declare name children
            XmlNodeList patterns_child = name.Item(0).ChildNodes;
            XmlNodeList defsalutations_child = name.Item(1).ChildNodes;
            XmlNodeList gen_child = name.Item(2).ChildNodes;

            XmlNodeList recognised_child = patterns_child.Item(0).ChildNodes;
            XmlNodeList unrecognised_child = patterns_child.Item(1).ChildNodes;

            // Populate pattern information
            downloadStatistics[10] = recognised_child.Item(0).InnerText; // male names
            downloadStatistics[11] = recognised_child.Item(1).InnerText; // female names
            downloadStatistics[12] = recognised_child.Item(2).InnerText; // neutral names
            downloadStatistics[13] = recognised_child.Item(3).InnerText; // joint names
            downloadStatistics[14] = recognised_child.Item(4).InnerText; // blank names
            downloadStatistics[15] = recognised_child.Item(5).InnerText; // default names

            downloadStatistics[16] = unrecognised_child.Item(0).InnerText; // single names
            downloadStatistics[17] = unrecognised_child.Item(1).InnerText; // other names

            // Populate defsalutation information
            downloadStatistics[18] = defsalutations_child.Item(0).InnerText; // male sal
            downloadStatistics[19] = defsalutations_child.Item(1).InnerText; // fem sal
            downloadStatistics[20] = defsalutations_child.Item(2).InnerText; // neutral sal
            downloadStatistics[21] = defsalutations_child.Item(3).InnerText; // joint sal
            downloadStatistics[22] = defsalutations_child.Item(4).InnerText; // blank sal
            downloadStatistics[23] = defsalutations_child.Item(5).InnerText; // def sal
            downloadStatistics[24] = defsalutations_child.Item(6).InnerText; // unrec sal

            // Populate general name info
            downloadStatistics[25] = gen_child.Item(0).InnerText; // poss deceased
            downloadStatistics[26] = gen_child.Item(1).InnerText; // has affix

            // ------- Address information
            XmlNodeList country_child = address.Item(0).ChildNodes;
            XmlNodeList gen_address_child = address.Item(1).ChildNodes;
            XmlNodeList uk_child = address.Item(2).ChildNodes;

            // Children
            XmlNodeList postcode_child = uk_child.Item(0).ChildNodes;
            XmlNodeList selcode_child = uk_child.Item(1).ChildNodes;
            XmlNodeList dps_child = uk_child.Item(2).ChildNodes;

            // Populate country information
            downloadStatistics[27] = country_child.Item(0).InnerText; // uk identified
            downloadStatistics[28] = country_child.Item(1).InnerText; // uk assumed
            downloadStatistics[29] = country_child.Item(2).InnerText; // overseas
            downloadStatistics[30] = country_child.Item(3).InnerText; // country blank

            // Populate general address information
            downloadStatistics[31] = gen_address_child.Item(0).InnerText; // business address
            downloadStatistics[32] = gen_address_child.Item(1).InnerText; // residential
            downloadStatistics[33] = gen_address_child.Item(2).InnerText; // suspect address
            downloadStatistics[34] = gen_address_child.Item(3).InnerText; // bfpo address

            // Populate uk information
            downloadStatistics[35] = postcode_child.Item(0).InnerText; // full zip
            downloadStatistics[36] = postcode_child.Item(1).InnerText; // outbound only
            downloadStatistics[37] = postcode_child.Item(2).InnerText; // no zip
            downloadStatistics[38] = postcode_child.Item(3).InnerText; // invalid zip

            downloadStatistics[39] = selcode_child.Item(0).InnerText; // Dps direct
            downloadStatistics[40] = selcode_child.Item(1).InnerText; // dps residue

            downloadStatistics[41] = dps_child.Item(0).InnerText; // dps present
            downloadStatistics[42] = dps_child.Item(1).InnerText; // dps default

            // Add download statistics to CSV
            StreamWriter csv_output;
            string concat_statistics = "";

            if (File.Exists((OutputDirectory() + "download_stats.csv")))
            {
                csv_output = new StreamWriter(new FileStream(OutputDirectory() + "download_stats.csv", FileMode.Append, FileAccess.Write, FileShare.Read));

                // Build line from array
                foreach (var index in downloadStatistics)
                {
                    concat_statistics = concat_statistics + index + ",";
                }

                // Write line to the csv file
                csv_output.WriteLine(concat_statistics);

                // Close csv
                csv_output.Close();

            }
            else
            {
                csv_output = new StreamWriter(new FileStream(OutputDirectory() + "download_stats.csv", FileMode.Append, FileAccess.Write, FileShare.Read));

                // Build line from array
                foreach (var index in downloadStatistics)
                {
                    concat_statistics = concat_statistics + index + ",";
                }

                // Retrieve headings
                FileHeadings headings = new FileHeadings();

                // Write line to the csv file
                csv_output.WriteLine(headings.GetDownloadStatsHeadings()); // Writing field names
                csv_output.WriteLine(concat_statistics); // Writing file information

                // Close csv
                csv_output.Close();
            }
        }

        /// <summary>
        /// Retrieves the file structure for the supplied mailing/download file from the accompanying XmlNode,
        /// and outputs to file.
        /// </summary>
        /// <param name="mailing_file">the name of the download file.</param>
        /// <param name="process">the XmlNode linked to the download file process.</param> 
        /// <example>
        /// <code>
        ///     // Prepend file headings if file does not exist
        ///     if (File.Exists(OutputDirectory() + "file_structures.csv") == false)
        ///     {
        ///         // Write file headings
        ///         StreamWriter file_writer = new StreamWriter(new FileStream(OutputDirectory() + "file_structures.csv", FileMode.Append, FileAccess.Write, FileShare.Read));
        ///     
        ///         // Get file headings
        ///         FileHeadings file_headings = new FileHeadings();
        ///     
        ///         // Write line
        ///         file_writer.WriteLine(file_headings.GetStructureHeadings());
        ///     
        ///         // Close file
        ///         file_writer.Dispose();
        ///     }
        ///     
        ///     // Declare variables
        ///     int field_count = 0;
        ///     string[,] field_array;
        ///     
        ///     // Retrieve structure node
        ///     XmlNode fields_node = process.SelectSingleNode("structure").FirstChild;
        ///     
        ///     // count fields
        ///     field_count = fields_node.ChildNodes.Count;
        ///     
        ///     // Initialise array
        ///     field_array = new string[field_count, 10];
        ///     field_array[0, 0] = mailing_file;
        ///     
        ///     // Initialise counter
        ///     int i = 0;
        ///     
        ///     // Loop through fields
        ///     foreach (XmlNode field in fields_node)
        ///     {
        ///         // Loop through each attribute
        ///         foreach (XmlNode attribute in field)
        ///         {
        ///             // Check attribute > retrieve value > place in correct array index
        ///             switch (attribute.Name)
        ///             {
        ///                 case "name":
        ///                     field_array[i, 1] = attribute.InnerText;
        ///                     break;
        ///                 case "start":
        ///                     field_array[i, 2] = attribute.InnerText;
        ///                     break;
        ///                 case "length":
        ///                     field_array[i, 3] = attribute.InnerText;
        ///                     break;
        ///                 case "data":
        ///                     field_array[i, 4] = "\"" + attribute.InnerText + "\""; // Apply quotes in case of commas in the data
        ///                     break;
        ///                 case "constant":
        ///                     field_array[i, 5] = attribute.InnerText;
        ///                     break;
        ///                 case "blanks":
        ///                     field_array[i, 6] = attribute.InnerText;
        ///                     break;
        ///                 case "average":
        ///                     field_array[i, 7] = attribute.InnerText;
        ///                     break;
        ///                 case "min":
        ///                     field_array[i, 8] = attribute.InnerText;
        ///                     break;
        ///                 case "max":
        ///                     field_array[i, 9] = attribute.InnerText;
        ///                     break;
        ///             }
        ///         }
        ///     
        ///     
        ///         // Output field structure to CSV
        ///         StreamWriter csv_output;
        ///         string concat_structures = "";
        ///     
        ///         csv_output = new StreamWriter(new FileStream(OutputDirectory() + "file_structures.csv", FileMode.Append, FileAccess.Write, FileShare.Read));
        ///     
        ///         // Build line from array
        ///         for (int j = 0; j &lt; 10; j++)
        ///     	{
        ///             concat_structures = concat_structures + field_array[i, j] + ",";
        ///         }
        ///     
        ///         // If new file - line break
        ///         if (i == 0)
        ///         {
        ///             csv_output.WriteLine();
        ///         }
        ///     
        ///         // Write line to the csv file - including field number (i)
        ///         csv_output.WriteLine(concat_structures + "" + (i + 1));
        ///     
        ///         // Close csv
        ///         csv_output.Close();
        ///     
        ///         // Iterate field counter
        ///         i++;
        ///     }
        /// </code>
        /// </example>
        private void ExportFileStructures(string mailing_file, XmlNode process)
        {           
            // Prepend file headings if file does not exist
            if (File.Exists(OutputDirectory() + "file_structures.csv") == false)
            {
                // Write file headings
                StreamWriter file_writer = new StreamWriter(new FileStream(OutputDirectory() + "file_structures.csv", FileMode.Append, FileAccess.Write, FileShare.Read));

                // Get file headings
                FileHeadings file_headings = new FileHeadings();

                // Write line
                file_writer.WriteLine(file_headings.GetStructureHeadings());

                // Close file
                file_writer.Dispose();
            }
            
            // Declare variables
            int field_count = 0;
            string[,] field_array;

            // Retrieve structure node
            XmlNode fields_node = process.SelectSingleNode("structure").FirstChild;
 
            // count fields
            field_count = fields_node.ChildNodes.Count;

            // Initialise array
            field_array = new string[field_count, 10];
            field_array[0, 0] = mailing_file;

            // Initialise counter
            int i = 0;

            // Loop through fields
            foreach (XmlNode field in fields_node)
            {
                // Loop through each field
                foreach (XmlNode attribute in field)
                {
                    // Check attribute > retrieve value > place in correct array index
                    switch (attribute.Name)
                    {
                        case "name":
                            field_array[i, 1] = attribute.InnerText;
                            break;
                        case "start":
                            field_array[i, 2] = attribute.InnerText;
                            break;
                        case "length":
                            field_array[i, 3] = attribute.InnerText;
                            break;
                        case "data":
                            field_array[i, 4] = "\"" + attribute.InnerText + "\""; // Apply quotes in case of commas in the data
                            break;
                        case "constant":
                            field_array[i, 5] = attribute.InnerText;
                            break;
                        case "blanks":
                            field_array[i, 6] = attribute.InnerText;
                            break;
                        case "average":
                            field_array[i, 7] = attribute.InnerText;
                            break;
                        case "min":
                            field_array[i, 8] = attribute.InnerText;
                            break;
                        case "max":
                            field_array[i, 9] = attribute.InnerText;
                            break;
                    }
                }


                // Output field structure to CSV
                StreamWriter csv_output;
                string concat_structures = "";

                csv_output = new StreamWriter(new FileStream(OutputDirectory() + "file_structures.csv", FileMode.Append, FileAccess.Write, FileShare.Read));

                // Build line from array
                for (int j = 0; j < 10; j++)
			    {
			        concat_structures = concat_structures + field_array[i,j] + ",";
			    }

                // If new file - line break
                if (i == 0)
                {
                    csv_output.WriteLine();
                }

                // Write line to the csv file - including field number (i)
                csv_output.WriteLine(concat_structures + "" + (i + 1));

                // Close csv
                csv_output.Close();

                // Iterate field counter
                i++;
            }
        }

        /// <summary>
        /// Retrieves and outputs to file, records that have addresses updated to an overseas address.
        /// </summary>
        /// <example>
        /// <code>
        ///     // *** RETRIEVE OVERSEAS REDIRECTS INFORMATION ***
        ///     XmlNode outputs;
        ///     string overseasRedirects = "";
        ///     string csvOutput = OutputDirectory() + "overseas_redirects.csv";
        ///     string header = "OVERSEAS_REDIRECTS";
        ///     
        ///     // Open _Merged.XML
        ///     XmlDocument xml_document = new XmlDocument();
        ///     xml_document.Load(MergedXMLPath());
        ///     
        ///     // Retrieve process node
        ///     XmlNodeList process_list = xml_document.DocumentElement.SelectNodes("/cygnus_reports/process");
        ///     
        ///     // Loop through each process
        ///     foreach (XmlNode process in process_list)
        ///     {
        ///         if (process.Attributes["name"].InnerText == "Trace 9")
        ///         {
        ///     
        ///             // Retrieve outputs node
        ///             outputs = process.SelectSingleNode("overview").SelectSingleNode("outputs");
        ///     
        ///             // Retrieve redirects figures
        ///             foreach (XmlNode child in outputs.ChildNodes)
        ///             {
        ///                 if (child.SelectSingleNode("name").InnerText == "tr9_Traced (O'Seas)")
        ///                 {
        ///                     overseasRedirects = child.SelectSingleNode("datasize").InnerText;
        ///                 }
        ///             }
        ///         }
        ///     }
        ///     
        ///     // *** EXPORT OVERSEAS REDIRECTS INFO TO CSV *** 
        ///     using (StreamWriter writer = new StreamWriter(csvOutput, true))
        ///     {
        ///         writer.WriteLine(header);
        ///         writer.WriteLine(overseasRedirects);
        ///     }
        /// </code>
        /// </example>
        private void OverseasRedirects()
        {
            // *** RETRIEVE OVERSEAS REDIRECTS INFORMATION ***
            XmlNode outputs;
            string overseasRedirects = "";
            string csvOutput = OutputDirectory() + "overseas_redirects.csv";
            string header = "OVERSEAS_REDIRECTS";

            // Open _Merged.XML
            XmlDocument xml_document = new XmlDocument();
            xml_document.Load(MergedXMLPath());

            // Retrieve process node
            XmlNodeList process_list = xml_document.DocumentElement.SelectNodes("/cygnus_reports/process");

            // Loop through each process
            foreach (XmlNode process in process_list)
            {
                if (process.Attributes["name"].InnerText == "Trace 9")
                {
                    // Retrieve outputs node
                    outputs = process.SelectSingleNode("overview").SelectSingleNode("outputs");

                    // Retrieve redirects figures
                    foreach (XmlNode child in outputs.ChildNodes)
                    {
                        if (child.SelectSingleNode("name").InnerText == "tr9_Traced (O'Seas)")
                        {
                            overseasRedirects = child.SelectSingleNode("datasize").InnerText;
                        }
                    }
                }  
            }

            // *** EXPORT OVERSEAS REDIRECTS INFO TO CSV *** 
            using (StreamWriter writer = new StreamWriter(csvOutput, true))
            {
                writer.WriteLine(header);
                writer.WriteLine(overseasRedirects);
            }
        }

        // Dedupe Functions

        /// <summary>
        /// Checks the _Merged.xml file for existence of dedupe modules and returns true or false.
        /// </summary>
        /// <returns>true or false as boolean</returns>
        /// <example>
        /// <code>
        ///     // Open _Merged.XML
        ///     XmlDocument xml_document = new XmlDocument();
        ///     xml_document.Load(MergedXMLPath());
        ///     bool dedupes_exist = false;
        ///     
        ///     // Retrieve process node
        ///     XmlNodeList process_list = xml_document.DocumentElement.SelectNodes("/cygnus_reports/process");
        ///     
        ///     // Loop through each process
        ///     foreach (XmlNode process in process_list)
        ///     {
        ///         // Store process attributes
        ///         XmlAttributeCollection process_attributes = process.Attributes;
        ///         string process_type = "";
        ///         string process_name = "";
        ///     
        ///         // loop through attributes
        ///         foreach (XmlAttribute attribute in process_attributes)
        ///         {
        ///             // Check attribute field
        ///             switch (attribute.Name)
        ///             {
        ///                 // if 'type' store contents (i.e. dedupe)
        ///                 case "type":
        ///                     process_type = attribute.Value;
        ///                     break;
        ///                 // If 'name' store contents (i.e. dedupe name)
        ///                 case "name":
        ///                     process_name = attribute.Value;
        ///                     break;
        ///             }
        ///         }
        ///     
        ///         // Check process is dedupe and the name is not a suppression module or post-suppression dedupe
        ///         if (process_type.ToUpper() == "DEDUPE")
        ///         {
        ///             if (process_name != "Post-Suppression Dedupe")
        ///             {
        ///                 if ((process_name.Contains("Suppressions") != true) &amp;&amp; (process_name.Contains("Locate") != true))
        ///                 {
        ///                     dedupes_exist = true;
        ///     
        ///                 }
        ///             }
        ///         }
        ///     }
        ///     
        ///     // If there are valid dedupes then return true
        ///     return dedupes_exist;
        /// </code>
        /// </example>
        private bool CheckDedupesExist()
        {
            // Open _Merged.XML
            XmlDocument xml_document = new XmlDocument();
            xml_document.Load(MergedXMLPath());
            bool dedupes_exist = false;
               
            // Retrieve process node
            XmlNodeList process_list = xml_document.DocumentElement.SelectNodes("/cygnus_reports/process");

            // Loop through each process
            foreach (XmlNode process in process_list)
            {
                // Store process attributes
                XmlAttributeCollection process_attributes = process.Attributes;
                string process_type = "";
                string process_name = "";

                // loop through attributes
                foreach (XmlAttribute attribute in process_attributes)
                {
                    // Check attribute field
                    switch (attribute.Name)
                    {
                        // if 'type' store contents (i.e. dedupe)
                        case "type":
                            process_type = attribute.Value;
                            break;
                        // If 'name' store contents (i.e. dedupe name)
                        case "name":
                            process_name = attribute.Value;
                            break;
                    }
                }

                // Check process is dedupe and the name is not a suppression module or post-suppression dedupe
                if (process_type.ToUpper() == "DEDUPE")
                {
                    if (process_name != "Post-Suppression Dedupe")
                    {
                        if ((process_name.Contains("Suppressions") != true) && (process_name.Contains("Locate") != true))
                        {
                           
                           dedupes_exist = true;
                           
                        }
                    }
                }
            }

            // If there are valid dedupes then return true
            return dedupes_exist;

        }

        /// <summary>
        /// Deletes any pre-existing output CSVs pertaining to dedupes before building a list
        /// of dedupes that exist in the Cygnus workflow, and displaying them on the form for
        /// the user.
        /// </summary>
        /// <example>
        /// <code>
        ///     // Delete existing dedupe information - TEST DIRECTORY
        ///     File.Delete(OutputDirectory() + "dedupe_information.csv");
        ///     File.Delete(OutputDirectory() + "dedupe_statistics.csv");
        ///     File.Delete(OutputDirectory() + "dedupe_examples.csv");
        ///     File.Delete(OutputDirectory() + "dedupe_totals.csv");
        ///     
        ///     // Retrieve Matrix files (store file paths)
        ///     string[] matrix_paths = Directory.GetFiles(OutputDirectory() + @"Matrices\");
        ///        
        ///     // Loop through each matrix
        ///     foreach (string matrix in matrix_paths)
        ///     {
        ///         // Delete matrix
        ///         File.Delete(matrix);
        ///     }
        ///     
        ///     // Open _Merged.XML
        ///     XmlDocument xml_document = new XmlDocument();
        ///     xml_document.Load(MergedXMLPath());
        ///     
        ///     // Retrieve process node
        ///     XmlNodeList process_list = xml_document.DocumentElement.SelectNodes("/cygnus_reports/process");
        ///     
        ///     // Loop through each process
        ///     foreach (XmlNode process in process_list)
        ///     {
        ///         // Store process attributes
        ///         XmlAttributeCollection process_attributes = process.Attributes;
        ///         string process_type = "";
        ///         string process_name = "";
        ///     
        ///         // loop through attributes
        ///         foreach (XmlAttribute attribute in process_attributes)
        ///         {
        ///             // Check attribute field
        ///             switch (attribute.Name)
        ///             {
        ///                 // if 'type' store contents (i.e. dedupe)
        ///                 case "type":
        ///                     process_type = attribute.Value;
        ///                     break;
        ///                 // If 'name' store contents (i.e. dedupe name)
        ///                 case "name":
        ///                     process_name = attribute.Value;
        ///                     break;
        ///             }
        ///         }
        ///     
        ///         // Check process is dedupe and the name is not a suppression module or post-suppression dedupe
        ///         if (process_type.ToUpper() == "DEDUPE")
        ///         {
        ///             if (process_name != "Post-Suppression Dedupe")
        ///             {
        ///                 if ((process_name.Contains("Suppressions") != true) &amp;&amp; (process_name.Contains("Locate") != true))
        ///                 {
        ///                     // Add process_name to the list box
        ///                     MailingFileList.Items.Add(process_name);  
        ///                 }
        ///             }
        ///         }
        ///     }
        /// </code>
        /// </example>
        private void RetrieveDedupeInformation()
        {
            // Delete existing dedupe information - TEST DIRECTORY
            File.Delete(OutputDirectory() + "dedupe_information.csv");
            File.Delete(OutputDirectory() + "dedupe_statistics.csv");
            File.Delete(OutputDirectory() + "dedupe_examples.csv");
            File.Delete(OutputDirectory() + "dedupe_totals.csv");

            // Retrieve Matrix files (store file paths)
            string[] matrix_paths = Directory.GetFiles(OutputDirectory() + @"Matrices\");
           
            // Loop through each matrix
            foreach (string matrix in matrix_paths)
            {
                // Delete matrix
                File.Delete(matrix);
            }

            // Open _Merged.XML
            XmlDocument xml_document = new XmlDocument();
            xml_document.Load(MergedXMLPath());

            // Retrieve process node
            XmlNodeList process_list = xml_document.DocumentElement.SelectNodes("/cygnus_reports/process");

            // Loop through each process
            foreach (XmlNode process in process_list)
            {
                // Store process attributes
                XmlAttributeCollection process_attributes = process.Attributes;
                string process_type = "";
                string process_name = "";

                // loop through attributes
                foreach (XmlAttribute attribute in process_attributes)
                {
                    // Check attribute field
                    switch (attribute.Name)
                    {
                        // if 'type' store contents (i.e. dedupe)
                        case "type":
                            process_type = attribute.Value;
                            break;
                        // If 'name' store contents (i.e. dedupe name)
                        case "name":
                            process_name = attribute.Value;
                            break;
                    }
                }

                // Check process is dedupe and the name is not a suppression module or post-suppression dedupe
                if (process_type.ToUpper() == "DEDUPE")
                {
                    if (process_name != "Post-Suppression Dedupe")
                    {
                        if ((process_name.Contains("Suppressions") != true) && (process_name.Contains("Locate") != true))
                        {
                            // Add process_name to the list box
                            MailingFileList.Items.Add(process_name);  
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Exports dedupe settings, including rules and criteria for example.
        /// </summary>
        /// <param name="dedupe">the XmlNode containing the information for the current dedupe.</param>
        /// <example>
        /// <code>
        ///     // Create download information array
        ///     String[] dedupeInformation;
        ///     dedupeInformation = new String[10];         
        ///     
        ///     // Retrieve nodes 
        ///     XmlNode dedupe_node = dedupe.SelectSingleNode("properties");
        ///     XmlNode criteria_node = dedupe_node.SelectSingleNode("criteria");
        ///     XmlNode level_node = dedupe_node.SelectSingleNode("level");
        ///     XmlNode rules_node = dedupe_node.SelectSingleNode("rules");
        ///     
        ///     // Store dedupe name
        ///     dedupeInformation[0] = dedupe.Attributes.GetNamedItem("name").Value;
        ///     
        ///     // Retrieve dedupe information
        ///     dedupeInformation[1] = criteria_node.SelectSingleNode("category").InnerText;
        ///     dedupeInformation[2] = criteria_node.SelectSingleNode("basis").InnerText;
        ///     dedupeInformation[3] = criteria_node.SelectSingleNode("criterion").InnerText;
        ///     dedupeInformation[4] = dedupe_node.SelectSingleNode("action").InnerText;
        ///     dedupeInformation[5] = level_node.SelectSingleNode("action").InnerText;
        ///     dedupeInformation[6] = level_node.SelectSingleNode("target").InnerText;
        ///     dedupeInformation[7] = rules_node.SelectSingleNode("mismatch").InnerText;
        ///     
        ///     // Retrieve forename matching
        ///     string[] levels = ForenameMatching(dedupeInformation[0]);
        ///     
        ///     // Populate array with the levels
        ///     dedupeInformation[8] =  levels[0]; // [deduping] Form needs to be added 
        ///     dedupeInformation[9] =  levels[1]; // [Suppressing] Form needs to be added
        ///     
        ///     // Write array information to CSV
        ///     StreamWriter csv_output;
        ///     string concat_information = "";
        ///     
        ///     if (File.Exists((OutputDirectory() + "dedupe_information.csv")))
        ///     {
        ///         csv_output = new StreamWriter(new FileStream(OutputDirectory() + "dedupe_information.csv", FileMode.Append, FileAccess.Write, FileShare.Read));
        ///     
        ///         // Build line from array
        ///         foreach (var index in dedupeInformation)
        ///         {
        ///             concat_information = concat_information + index + ",";
        ///         }
        ///     
        ///         // Write line to the csv file
        ///         csv_output.WriteLine(concat_information);
        ///     
        ///         // Close csv
        ///         csv_output.Close();
        ///     
        ///     }
        ///     else
        ///     {
        ///         csv_output = new StreamWriter(new FileStream(OutputDirectory() + "dedupe_information.csv", FileMode.Append, FileAccess.Write, FileShare.Read));
        ///     
        ///         // Build line from array
        ///         foreach (var index in dedupeInformation)
        ///         {
        ///         concat_information = concat_information + index + ",";
        ///         }
        ///     
        ///         // Retrieve headings
        ///         FileHeadings headings = new FileHeadings();
        ///     
        ///         // Write line to the csv file
        ///         csv_output.WriteLine(headings.GetDedupeInformationHeadings()); // Writing field names
        ///         csv_output.WriteLine(concat_information); // Writing file information
        ///     
        ///         // Close csv
        ///         csv_output.Close();
        ///     }
        /// </code>
        /// </example>
        private void ExportDedupeInformation(XmlNode dedupe)
        {
            // Create download information array
            String[] dedupeInformation;
            dedupeInformation = new String[10];         

            // Retrieve nodes 
            XmlNode dedupe_node = dedupe.SelectSingleNode("properties");
            XmlNode criteria_node = dedupe_node.SelectSingleNode("criteria");
            XmlNode level_node = dedupe_node.SelectSingleNode("level");
            XmlNode rules_node = dedupe_node.SelectSingleNode("rules");

            // Store dedupe name
            dedupeInformation[0] = dedupe.Attributes.GetNamedItem("name").Value;

            // Retrieve dedupe information
            dedupeInformation[1] = criteria_node.SelectSingleNode("category").InnerText;
            dedupeInformation[2] = criteria_node.SelectSingleNode("basis").InnerText;
            dedupeInformation[3] = criteria_node.SelectSingleNode("criterion").InnerText;
            dedupeInformation[4] = dedupe_node.SelectSingleNode("action").InnerText;
            dedupeInformation[5] = level_node.SelectSingleNode("action").InnerText;
            dedupeInformation[6] = level_node.SelectSingleNode("target").InnerText;
            dedupeInformation[7] = rules_node.SelectSingleNode("mismatch").InnerText;

            // Retrieve forename matching
            string[] levels = ForenameMatching(dedupeInformation[0]); 

            // Populate array with the levels
            dedupeInformation[8] =  levels[0]; // [deduping] Form needs to be added 
            dedupeInformation[9] =  levels[1]; // [Suppressing] Form needs to be added

            // Write array information to CSV
            StreamWriter csv_output;
            string concat_information = "";

            if (File.Exists((OutputDirectory() + "dedupe_information.csv")))
            {
                csv_output = new StreamWriter(new FileStream(OutputDirectory() + "dedupe_information.csv", FileMode.Append, FileAccess.Write, FileShare.Read));

                // Build line from array
                foreach (var index in dedupeInformation)
                {
                    concat_information = concat_information + index + ",";
                }

                // Write line to the csv file
                csv_output.WriteLine(concat_information);

                // Close csv
                csv_output.Close();

            }
            else
            {
                csv_output = new StreamWriter(new FileStream(OutputDirectory() + "dedupe_information.csv", FileMode.Append, FileAccess.Write, FileShare.Read));

                // Build line from array
                foreach (var index in dedupeInformation)
                {
                    concat_information = concat_information + index + ",";
                }

                // Retrieve headings
                FileHeadings headings = new FileHeadings();

                // Write line to the csv file
                csv_output.WriteLine(headings.GetDedupeInformationHeadings()); // Writing field names
                csv_output.WriteLine(concat_information); // Writing file information

                // Close csv
                csv_output.Close();
            }

        }

        /// <summary>
        /// Produces the file statistics for the supplied dedupe, inc total drops, inter, intra etc.
        /// </summary>
        /// <param name="dedupe">the XmlNode containing the dedupe statistics</param>
        /// <example>
        /// <code>
        ///     // Write headings to dedupe statistics 
        ///     
        ///     // If file exists - prepend a line
        ///     if (File.Exists(OutputDirectory() + "dedupe_statistics.csv"))
        ///     {
        ///         StreamWriter write_headings;
        ///         FileHeadings stats_headings = new FileHeadings();
        ///     
        ///         // Instantiate the file write
        ///         write_headings = new StreamWriter(new FileStream(OutputDirectory() + "dedupe_statistics.csv", FileMode.Append, FileAccess.Write, FileShare.Read));
        ///     
        ///         // Write headings to file
        ///         write_headings.WriteLine();
        ///         write_headings.WriteLine(stats_headings.GetDedupeStatsHeadings());
        ///     
        ///         // Close file
        ///         write_headings.Dispose();
        ///     }
        ///     else // Do not prepend a line
        ///     {
        ///         StreamWriter write_headings;
        ///         FileHeadings stats_headings = new FileHeadings();
        ///     
        ///         // Instantiate the file write
        ///         write_headings = new StreamWriter(new FileStream(OutputDirectory() + "dedupe_statistics.csv", FileMode.Append, FileAccess.Write, FileShare.Read));
        ///     
        ///         // Write headings to file
        ///         write_headings.WriteLine(stats_headings.GetDedupeStatsHeadings());
        ///     
        ///         // Close file
        ///         write_headings.Dispose();
        ///     }
        ///     
        ///     // Retrieve the inputs node
        ///     XmlNodeList input_files = dedupe.SelectSingleNode("inputs").SelectSingleNode("inputs").ChildNodes;
        ///     int file_count = 1;
        ///     
        ///     // For each file in the dedupe
        ///     foreach (XmlNode file in input_files)
        ///     {
        ///         // Record attributes if mailing file
        ///         if (file.Name == "input")
        ///         {
        ///             string[] dedupe_statistics = new string[12];
        ///     
        ///             if (file_count==1)
        ///             {
        ///                 dedupe_statistics[0] = dedupe.Attributes.GetNamedItem("name").Value; 
        ///             }
        ///     
        ///             // Store incoming file statistics
        ///             dedupe_statistics[1] = file.SelectSingleNode("name").InnerText;
        ///             dedupe_statistics[2] = file.SelectSingleNode("listcode").InnerText;
        ///             dedupe_statistics[3] = file.SelectSingleNode("priority").InnerText;
        ///             dedupe_statistics[4] = file.SelectSingleNode("quantity").InnerText;
        ///     
        ///             // Store drop statistics
        ///             dedupe_statistics[5] = file.SelectSingleNode("drops").SelectSingleNode("suppressed").InnerText;
        ///             dedupe_statistics[6] = file.SelectSingleNode("drops").SelectSingleNode("inter").InnerText;
        ///             dedupe_statistics[7] = file.SelectSingleNode("drops").SelectSingleNode("intra").InnerText;
        ///             dedupe_statistics[8] = file.SelectSingleNode("drops").SelectSingleNode("total").InnerText;
        ///     
        ///             // Store output statistics
        ///             dedupe_statistics[9] = file.SelectSingleNode("output").SelectSingleNode("uniques").InnerText;
        ///             dedupe_statistics[10] = file.SelectSingleNode("output").SelectSingleNode("selected").InnerText;
        ///             dedupe_statistics[11] = file.SelectSingleNode("output").SelectSingleNode("total").InnerText;
        ///     
        ///             // Print statistics for file to csv
        ///             StreamWriter csv_output;
        ///             string concat_information = "";
        ///     
        ///             csv_output = new StreamWriter(new FileStream(OutputDirectory() + "dedupe_statistics.csv", FileMode.Append, FileAccess.Write, FileShare.Read));
        ///     
        ///             // Build line from array
        ///             foreach (var index in dedupe_statistics)
        ///             {
        ///                 concat_information = concat_information + index + ",";
        ///             }
        ///     
        ///             // Write line to the csv file
        ///             csv_output.WriteLine(concat_information);
        ///     
        ///             // Append space if last file
        ///             if (file_count == input_files.Count)
        ///             {
        ///                 csv_output.WriteLine();
        ///             }
        ///     
        ///             // Close csv
        ///             csv_output.Close();
        ///     
        ///             // Increment file count
        ///             file_count++;
        ///         }
        ///     }
        /// </code>
        /// </example>
        private void ExportDedupeStatistics(XmlNode dedupe)
        {
            // Write headings to dedupe statistics 

            // If file exists - prepend a line
            if (File.Exists(OutputDirectory() + "dedupe_statistics.csv"))
            {
                StreamWriter write_headings;
                FileHeadings stats_headings = new FileHeadings();

                // Instantiate the file write
                write_headings = new StreamWriter(new FileStream(OutputDirectory() + "dedupe_statistics.csv", FileMode.Append, FileAccess.Write, FileShare.Read));

                // Write headings to file
                write_headings.WriteLine();
                write_headings.WriteLine(stats_headings.GetDedupeStatsHeadings());

                // Close file
                write_headings.Dispose();

            }
            else // Do not prepend a line
            {
                StreamWriter write_headings;
                FileHeadings stats_headings = new FileHeadings();

                // Instantiate the file write
                write_headings = new StreamWriter(new FileStream(OutputDirectory() + "dedupe_statistics.csv", FileMode.Append, FileAccess.Write, FileShare.Read));

                // Write headings to file
                write_headings.WriteLine(stats_headings.GetDedupeStatsHeadings());

                // Close file
                write_headings.Dispose();
            }
            
            // Retrieve the inputs node
            XmlNodeList input_files = dedupe.SelectSingleNode("inputs").SelectSingleNode("inputs").ChildNodes;
            int file_count = 1;

            // For each file in the dedupe
            foreach (XmlNode file in input_files)
            {
                // Record attributes if mailing file
                if (file.Name == "input")
                {
                    string[] dedupe_statistics = new string[12];

                    if (file_count==1)
                    {
                        dedupe_statistics[0] = dedupe.Attributes.GetNamedItem("name").Value; 
                    }
                    
                    // Store incoming file statistics
                    dedupe_statistics[1] = file.SelectSingleNode("name").InnerText;
                    dedupe_statistics[2] = file.SelectSingleNode("listcode").InnerText;
                    dedupe_statistics[3] = file.SelectSingleNode("priority").InnerText;
                    dedupe_statistics[4] = file.SelectSingleNode("quantity").InnerText;

                    // Store drop statistics
                    dedupe_statistics[5] = file.SelectSingleNode("drops").SelectSingleNode("suppressed").InnerText;
                    dedupe_statistics[6] = file.SelectSingleNode("drops").SelectSingleNode("inter").InnerText;
                    dedupe_statistics[7] = file.SelectSingleNode("drops").SelectSingleNode("intra").InnerText;
                    dedupe_statistics[8] = file.SelectSingleNode("drops").SelectSingleNode("total").InnerText;

                    // Store output statistics
                    dedupe_statistics[9] = file.SelectSingleNode("output").SelectSingleNode("uniques").InnerText;
                    dedupe_statistics[10] = file.SelectSingleNode("output").SelectSingleNode("selected").InnerText;
                    dedupe_statistics[11] = file.SelectSingleNode("output").SelectSingleNode("total").InnerText;

                    // Print statistics for file to csv
                    StreamWriter csv_output;
                    string concat_information = "";

                    csv_output = new StreamWriter(new FileStream(OutputDirectory() + "dedupe_statistics.csv", FileMode.Append, FileAccess.Write, FileShare.Read));

                    // Build line from array
                    foreach (var index in dedupe_statistics)
                    {
                        concat_information = concat_information + index + ",";
                    }

                    // Write line to the csv file
                    csv_output.WriteLine(concat_information);

                    // Append space if last file
                    if (file_count == input_files.Count)
                    {
                        csv_output.WriteLine();
                    }

                    // Close csv
                    csv_output.Close();

                    // Increment file count
                    file_count++;

                }
            }
        }

        /// <summary>
        /// Exports examples of the duplicate groups from the supplied dedupe
        /// </summary>
        /// <param name="dedupe">the XmlNode containing the dedupe information</param>
        /// <example>
        /// <code>
        ///     if (File.Exists(OutputDirectory() + "dedupe_examples.csv") == false)
        ///     {
        ///         // Write headings to dedupe statistics 
        ///         StreamWriter write_headings;
        ///         FileHeadings stats_headings = new FileHeadings();
        ///     
        ///         // Instantiate the file write
        ///         write_headings = new StreamWriter(new FileStream(OutputDirectory() + "dedupe_examples.csv", FileMode.Append, FileAccess.Write, FileShare.Read));
        ///     
        ///         // Write headings to file
        ///         write_headings.WriteLine(stats_headings.GetDedupeExampleHeadings());
        ///     
        ///         // Close file
        ///         write_headings.Dispose();
        ///     }
        ///     
        ///     // Name record switch
        ///     int name_recorded = 1;
        ///     
        ///     // Build array to output
        ///     string[] output = new String[7];
        ///     
        ///     // Select dedupe groups
        ///     XmlNode groups_node = dedupe.SelectSingleNode("groups");
        ///     
        ///     // Loop through group nodes
        ///     foreach (XmlNode group_node in groups_node.ChildNodes)
        ///     {
        ///         // Check the node is dupes => select node
        ///         if (group_node.Name == "dupes")
        ///         {
        ///             // Loop through the group
        ///             foreach (XmlNode group in group_node.ChildNodes)
        ///             {   
        ///                 // Loop through the details
        ///                 foreach (XmlNode attr in group)
        ///                 {
        ///                     if (attr.Name == "number")
        ///                     {
        ///                         output[1] =  attr.InnerText;  
        ///                     }
        ///     
        ///                     if (attr.Name == "member") // Select group member information
        ///                     {
        ///                         foreach (XmlNode mem_attr in attr.ChildNodes)
        ///                         {
        ///                             // Select attributes according to name
        ///                             switch (mem_attr.Name)
        ///                             {
        ///                                 case "listcode":
        ///                                     output[2] = mem_attr.InnerText;
        ///                                     break;
        ///                                 case "priority":
        ///                                     output[3] = mem_attr.InnerText;
        ///                                     break;
        ///                                 case "status": 
        ///                                     output[4] = mem_attr.InnerText;
        ///                                         
        ///                                     // If not the first record in the group then don't print the group number
        ///                                     if (output[4] != "SELECTED")
        ///                                     {
        ///                                         output[1] = "";
        ///                                     }
        ///                                        break;
        ///                                 case "name":
        ///                                     output[5] = "\"" + mem_attr.InnerText + "\"";
        ///                                     break;
        ///                                 case "address":
        ///                                     output[6] = "\"" + mem_attr.InnerText + "\"";
        ///                                     break;
        ///                             }
        ///                         }
        ///     
        ///                         // Print statistics for file to csv
        ///                         StreamWriter csv_output;
        ///                         string concat_information = "";
        ///     
        ///                         csv_output = new StreamWriter(new FileStream(OutputDirectory() + "dedupe_examples.csv", FileMode.Append, FileAccess.Write, FileShare.Read));
        ///     
        ///                         // Store dedupe name
        ///                         if (name_recorded == 1)
        ///                         {
        ///                             output[0] = dedupe.Attributes.GetNamedItem("name").Value;
        ///                             name_recorded = 0;
        ///                         }
        ///                         else
        ///                         {
        ///                             output[0] = "";
        ///                         }
        ///     
        ///                         // Build line from array
        ///                         foreach (var index in output)
        ///                         {
        ///                             concat_information = concat_information + index + ",";
        ///                         }
        ///     
        ///                         // Write line to the csv file
        ///                         csv_output.WriteLine(concat_information);
        ///     
        ///                         // Close csv
        ///                         csv_output.Close();
        ///                     }
        ///                 }
        ///             }
        ///         }
        ///     }
        /// </code>
        /// </example>
        private void ExportDedupeExamples(XmlNode dedupe)
        {
            if (File.Exists(OutputDirectory() + "dedupe_examples.csv") == false)
	        {
		        // Write headings to dedupe statistics 
                StreamWriter write_headings;
                FileHeadings stats_headings = new FileHeadings();

                // Instantiate the file write
                write_headings = new StreamWriter(new FileStream(OutputDirectory() + "dedupe_examples.csv", FileMode.Append, FileAccess.Write, FileShare.Read));

                // Write headings to file
                write_headings.WriteLine(stats_headings.GetDedupeExampleHeadings());

                // Close file
                write_headings.Dispose();
	        }
            
            // Name record switch
            int name_recorded = 1;

            // Build array to output
            string[] output = new String[7];

            // Select dedupe groups
            XmlNode groups_node = dedupe.SelectSingleNode("groups");

            // Loop through group nodes
            foreach (XmlNode group_node in groups_node.ChildNodes)
            {
                // Check the node is dupes => select node
                if (group_node.Name == "dupes")
                {
                    // Loop through the group
                    foreach (XmlNode group in group_node.ChildNodes)
                    {   
                        // Loop through the details
                        foreach (XmlNode attr in group)
                        {
                            if (attr.Name == "number")
                            {
                                output[1] =  attr.InnerText;  
                            }

                            if (attr.Name == "member") // Select group member information
                            {
                                foreach (XmlNode mem_attr in attr.ChildNodes)
                                {

                                    // Select attributes according to name
                                    switch (mem_attr.Name)
                                    {
                                        case "listcode":
                                            output[2] = mem_attr.InnerText;
                                            break;
                                        case "priority":
                                            output[3] = mem_attr.InnerText;
                                            break;
                                        case "status":
                                            
                                            output[4] = mem_attr.InnerText;
                                            
                                            // If not the first record in the group then don't print the group number
                                            if (output[4] != "SELECTED")
                                            {
                                                output[1] = "";
                                            }

                                            break;
                                        case "name":
                                            output[5] = "\"" + mem_attr.InnerText + "\"";
                                            break;
                                        case "address":
                                            output[6] = "\"" + mem_attr.InnerText + "\"";
                                            break;
                                    }
                                }

                                // Print statistics for file to csv
                                StreamWriter csv_output;
                                string concat_information = "";

                                csv_output = new StreamWriter(new FileStream(OutputDirectory() + "dedupe_examples.csv", FileMode.Append, FileAccess.Write, FileShare.Read));

                                // Store dedupe name
                                if (name_recorded == 1)
                                {
                                    output[0] = dedupe.Attributes.GetNamedItem("name").Value;
                                    name_recorded = 0;
                                }
                                else
                                {
                                    output[0] = "";
                                }

                                // Build line from array
                                foreach (var index in output)
                                {
                                    concat_information = concat_information + index + ",";
                                }

                                // Write line to the csv file
                                csv_output.WriteLine(concat_information);

                                // Close csv
                                csv_output.Close();
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Exports the matrix of drops for files and suppression files from the supplied dedupe.
        /// </summary>
        /// <param name="dedupe">The XmlNode of the dedupe</param>
        /// <example>
        /// <code>
        ///     // Create file name
        ///     String file_name = dedupe.Attributes.GetNamedItem("name").Value + "_drops_matrix.csv";
        ///     
        ///     // Delete old matrix
        ///     File.Delete(OutputDirectory() + file_name);
        ///     
        ///     // Select matrix child
        ///     XmlNode matrix_node = dedupe.SelectSingleNode("drops").SelectSingleNode("matrix");
        ///     
        ///     // Variables
        ///     int entries = 0;
        ///     int matrix_rows = 1;
        ///     int matrix_columns = 3;
        ///     string file_header = "";
        ///     int stop_files = 0;
        ///     
        ///     // Count number of entries
        ///     foreach  (XmlNode node in matrix_node.ChildNodes)
        ///     {
        ///         if (node.Name == "entry")
        ///         {
        ///             // Count entry
        ///             entries++;
        ///     
        ///             // Count stop files
        ///             if (node.SelectSingleNode("priority").InnerText == "  A")
        ///             {
        ///                 stop_files++;
        ///             }
        ///         }
        ///     }
        ///     
        ///     // Matrix size 
        ///     matrix_rows += entries;
        ///     matrix_columns += (entries - stop_files);
        ///     
        ///     // Build file header
        ///     for (int i = 1; i &lt;= matrix_columns; i++)
        ///     {
        ///         switch (i)
        ///         {
        ///             case 1: // Col 1
        ///                 file_header = "Description,";
        ///                 break;
        ///             case 2: // Col 2
        ///                 file_header += "Priority,";
        ///                 break;
        ///             case 3: // Col 3
        ///                 file_header += "Index,";
        ///                 break;
        ///             default: // Rest - variable
        ///                 file_header += (i - 3) + ",";
        ///                 break;
        ///         }
        ///     }
        ///     
        ///     // Write header to file
        ///     // Print statistics for file to csv
        ///     StreamWriter header_output;
        ///     
        ///     header_output = new StreamWriter(new FileStream(OutputDirectory() + @"Matrices\" + file_name, FileMode.Append, FileAccess.Write, FileShare.Read));
        ///     
        ///     // Write line to the csv file
        ///     header_output.WriteLine(file_header);
        ///     
        ///     // Close csv
        ///     header_output.Close();
        ///     
        ///     // Build file rows
        ///     foreach (XmlNode node in matrix_node.ChildNodes)
        ///     {
        ///         if (node.Name == "entry")
        ///         {
        ///             // Build line
        ///             string current_entry;
        ///             XmlNode lists = node.SelectSingleNode("lists");
        ///     
        ///             // Store static information
        ///             current_entry = node.SelectSingleNode("name").InnerText + ",";
        ///             current_entry += node.SelectSingleNode("priority").InnerText + ",";
        ///             current_entry += node.SelectSingleNode("index").InnerText + ",";
        ///     
        ///             // for each list
        ///             foreach (XmlNode list in lists.ChildNodes)
        ///             {
        ///                 // If a list - not [totals]
        ///                 if (list.Name == "list")
        ///                 {
        ///                     // Store matches
        ///                     current_entry += list.SelectSingleNode("matches").InnerText + ",";
        ///                 }
        ///             }
        ///     
        ///             // Write the current line
        ///             StreamWriter entry_output;
        ///     
        ///             entry_output = new StreamWriter(new FileStream(OutputDirectory() + @"Matrices\" + file_name, FileMode.Append, FileAccess.Write, FileShare.Read));
        ///     
        ///             // Write line to the csv file
        ///             entry_output.WriteLine(current_entry);
        ///     
        ///             // Close csv
        ///             entry_output.Close(); 
        ///     
        ///         }
        ///     }
        /// </code>
        /// </example>
        private void ExportDropsMatrix(XmlNode dedupe)
        {
            // Create file name
            String file_name = dedupe.Attributes.GetNamedItem("name").Value + "_drops_matrix.csv";
            
            // Delete old matrix
            File.Delete(OutputDirectory() + file_name);
            
            // Select matrix child
            XmlNode matrix_node = dedupe.SelectSingleNode("drops").SelectSingleNode("matrix");

            // Variables
            int entries = 0;
            int matrix_rows = 1;
            int matrix_columns = 3;
            string file_header = "";
            int stop_files = 0;

            // Count number of entries
            foreach  (XmlNode node in matrix_node.ChildNodes)
            {
                if (node.Name == "entry")
                {
                    // Count entry
                    entries++;

                    // Count stop files
                    if (node.SelectSingleNode("priority").InnerText == "  A")
                    {
                        stop_files++;
                    }    
                }
            }

            // Matrix size 
            matrix_rows += entries;
            matrix_columns += (entries - stop_files);

            // Build file header
            for (int i = 1; i <= matrix_columns; i++)
            {
                switch (i)
                {
                    case 1: // Col 1
                        file_header = "Description,";
                        break;
                    case 2: // Col 2
                        file_header += "Priority,";
                        break;
                    case 3: // Col 3
                        file_header += "Index,";
                        break;
                    default: // Rest - variable
                        file_header += (i - 3) + ",";
                        break;
                }
            }

            // Write header to file
            // Print statistics for file to csv
            StreamWriter header_output;

            header_output = new StreamWriter(new FileStream(OutputDirectory() + @"Matrices\" + file_name, FileMode.Append, FileAccess.Write, FileShare.Read));

            // Write line to the csv file
            header_output.WriteLine(file_header);

            // Close csv
            header_output.Close();

            // Build file rows
            foreach (XmlNode node in matrix_node.ChildNodes)
            {
                if (node.Name == "entry")
                {
                    // Build line
                    string current_entry;
                    XmlNode lists = node.SelectSingleNode("lists");
                    
                    // Store static information
                    current_entry = node.SelectSingleNode("name").InnerText + ",";
                    current_entry += node.SelectSingleNode("priority").InnerText + ",";
                    current_entry += node.SelectSingleNode("index").InnerText + ",";

                    // for each list
                    foreach (XmlNode list in lists.ChildNodes)
                    {
                        // If a list - not [totals]
                        if (list.Name == "list")
                        {
                            // Store matches
                            current_entry += list.SelectSingleNode("matches").InnerText + ",";
                        }
                    }

                    // Write the current line
                    StreamWriter entry_output;

                    entry_output = new StreamWriter(new FileStream(OutputDirectory() + @"Matrices\" + file_name, FileMode.Append, FileAccess.Write, FileShare.Read));

                    // Write line to the csv file
                    entry_output.WriteLine(current_entry);

                    // Close csv
                    entry_output.Close(); 

                }
            }
        }

        /// <summary>
        /// Presents the dedupe forename matching level to the user.
        /// </summary>
        /// <param name="dedupe">takes the name of a dedupe module as a string</param>
        /// <returns>an array of size two containing both the forename matching levels for deduping [0] and suppressing [1]</returns>
        /// <example>
        /// <code>
        ///     // Array of size two
        ///     string[] levels = new String[2];
        ///     
        ///     // Set text in form
        ///     Window_ForenameMatching window = new Window_ForenameMatching();
        ///     window.SetWindowText("Please enter the forename matching level for '" + dedupe + "'.");
        ///     
        ///     // Show form to user
        ///     window.ShowDialog();
        ///     
        ///     // Populate array 
        ///     levels[0] = window.GetDedupingLevel();
        ///     levels[1] = window.GetSuppressionLevel();
        ///     
        ///     // Return array
        ///     return levels;
        /// </code>
        /// </example>
        private string[] ForenameMatching(string dedupe)
        {
            // Array of size two
            string[] levels = new String[2];

            // Set text in form
            Window_ForenameMatching window = new Window_ForenameMatching();
            window.SetWindowText("Please enter the forename matching level for '" + dedupe + "'.");

            // Show form to user
            window.ShowDialog();

            // Populate array 
           levels[0] = window.GetDedupingLevel();
           levels[1] = window.GetSuppressionLevel();
  
            // Return array
            return levels;
        }

        // Industry Suppression Functions

        /// <summary>
        /// Determines existence of the Cygnus exported industry suppressions file
        /// </summary>
        /// <returns>boolean true or false</returns>
        /// <example>
        /// <code>
        ///     return File.Exists(RootOnServer() + @"\Industry Suppression Drops.csv");
        /// </code>
        /// </example>
        private bool CheckISExist()
        {
            return File.Exists(RootOnServer() + @"\Industry Suppression Drops.csv");
        }

        /// <summary>
        /// Deletes any previous output files and calls the methods that process both the counts and costs.
        /// </summary>
        /// <example>
        /// <code>
        ///     // Delete existing Industry Suppression information
        ///     File.Delete(OutputDirectory() + "industry_suppressions.csv");
        ///     File.Delete(OutputDirectory() + "industry_costs.csv");
        ///     
        ///     // Export Suppression Counts
        ///     ExportIndustrySuppressions();
        ///     // Calcuate Suppression Totals inc costs
        ///     ExportSuppressionCosts(CalculateIndustryTotals());
        /// </code>
        /// </example>
        private void ProcessIndustrySuppressions()
        {
            // Delete existing Industry Suppression information
            File.Delete(OutputDirectory() + "industry_suppressions.csv");
            File.Delete(OutputDirectory() + "industry_costs.csv");

            // Export Suppression Counts
            ExportIndustrySuppressions();
            // Calcuate Suppression Totals inc costs
            ExportSuppressionCosts(CalculateIndustryTotals());
        }

        /// <summary>
        /// Reads the industry suppression drops csv exported from the Cygnus workflow, builds a unique list of filenames, and then 
        /// reads back through the same csv incrementing the total hits for each list, before exporting in a new csv.
        /// </summary>
        /// <example>
        /// <code>
        ///     // IS file path
        ///     string file_path = RootOnServer() + @"\Industry Suppression Drops.csv";
        ///     ArrayList is_csv = new ArrayList();
        ///     ArrayList file_names = new ArrayList();
        ///     
        ///     // Set columns
        ///     int filenames = 5;
        ///     int supphit = 3;
        ///     int line_count = 1;
        ///     
        ///     // Open CSV
        ///     StreamReader read_csv = new StreamReader(new FileStream(file_path, FileMode.Open, FileAccess.Read)); 
        ///     
        ///     // Read CSV to ArrayList
        ///     do
        ///     {
        ///         // Populate array [one line per index]
        ///         is_csv.Add(read_csv.ReadLine());
        ///     } while (read_csv.EndOfStream == false);
        ///     
        ///     // Close CSV
        ///     read_csv.Close();
        ///     
        ///     // Cycle through array of industry suppressions
        ///     foreach (string line in is_csv)
        ///     {
        ///         // Break line into array
        ///         string[] record = line.Split(new Char[] { ',' });
        ///     
        ///         switch (line_count)
        ///         {
        ///             case 1:
        ///                 // Ignore file heading
        ///                 // Iterate line count
        ///                 line_count++;
        ///                 break;
        ///             case 2:
        ///                 // Store first filename
        ///                 file_names.Add(record[filenames]);
        ///                 // Iterate line count
        ///                 line_count++;
        ///                 break;
        ///             default:
        ///     
        ///                 // Set flag
        ///                 bool exists = false;
        ///     
        ///                 // Check back to added filenames
        ///                 foreach (string filename in file_names)
        ///                 {
        ///                     // if filename already exists then flag
        ///                     if (filename == record[filenames])
        ///                     {
        ///                         exists = true;
        ///                     }
        ///                 }
        ///     
        ///                 // Check flag
        ///                 if (exists == false)
        ///                 {
        ///                     // Add to list of filenames
        ///                     file_names.Add(record[filenames]);
        ///                 }
        ///     
        ///                 // Iterate line count
        ///                 line_count++;
        ///                 break;
        ///         }
        ///     }
        ///     
        ///     // Loop through each file
        ///     foreach (string file in file_names)
        ///     {
        ///         // Create array
        ///         string[] file_suppressions = new string[15];
        ///     
        ///         // Store file name in multi-dimensional array
        ///         file_suppressions[0] = file;
        ///         int ncoa_redir = 0, gas_react = 0, ncoa_dec = 0, tbr = 0, ndr = 0, morta = 0, mps_dec = 0, ncoa = 0, gas = 0, mps = 0;
        ///     
        ///         // Match back file with is array
        ///         foreach (string line in is_csv)
        ///         {
        ///             // Split hit into parts
        ///             string[] parts = line.Split(new Char[] { ',' });
        ///     
        ///             // Check filename matches
        ///             if (file == parts[filenames])
        ///             {
        ///                 // Record input vol
        ///                 file_suppressions[1] = parts[4].ToString();
        ///                 file_suppressions[1] = file_suppressions[1].Replace("\"", "");
        ///     
        ///                 // Switch between hits
        ///                 switch (parts[supphit])
        ///                 {
        ///                     case "\"NCOA REDIR\"":
        ///                         ncoa_redir++;
        ///                         file_suppressions[2] = ncoa_redir.ToString();
        ///                         break;
        ///                     case "\"GAS REACT\"":
        ///                         gas_react++;
        ///                         file_suppressions[3] = gas_react.ToString();
        ///                         break;
        ///                     case "\"NCOA Deceased\"":
        ///                         ncoa_dec++;
        ///                         file_suppressions[4] = ncoa_dec.ToString();
        ///                         break;
        ///                     case "\"TBR\"":
        ///                         tbr++;
        ///                         file_suppressions[5] = tbr.ToString();
        ///                         break;
        ///                     case "\"NDR\"":
        ///                         ndr++;
        ///                         file_suppressions[6] = ndr.ToString();
        ///                         break;
        ///                     case "\"MORTA\"":
        ///                         morta++;
        ///                         file_suppressions[7] = morta.ToString();
        ///                         break;
        ///                     case "\"MPS Deceased\"":
        ///                         mps_dec++;
        ///                         file_suppressions[8] = mps_dec.ToString();
        ///                         break;
        ///                     case "\"NCOA\"":
        ///                         ncoa++;
        ///                         file_suppressions[9] = ncoa.ToString();
        ///                         break;
        ///                     case "\"GAS\"":
        ///                         gas++;
        ///                         file_suppressions[10] = gas.ToString();
        ///                         break;
        ///                     case "\"MPS\"":
        ///                         mps++;
        ///                         file_suppressions[11] = mps.ToString();
        ///                         break;
        ///                     default:
        ///                         break;
        ///                 }
        ///             }
        ///     }
        ///     
        ///     // Totals drops
        ///     int totals = ncoa_dec + tbr + ndr + morta + mps_dec + ncoa + gas + mps;
        ///     file_suppressions[12] = totals.ToString();
        ///     
        ///     // File output
        ///     file_suppressions[13] = Convert.ToString(Convert.ToInt64(file_suppressions[1]) - totals);
        ///     
        ///     file_suppressions[0] = file_suppressions[0].Replace("\"", "");
        ///     
        ///     // Total hits
        ///     file_suppressions[14] = Convert.ToString(Convert.ToInt64(file_suppressions[12]) + Convert.ToInt64(file_suppressions[2]) + Convert.ToInt64(file_suppressions[3]));               
        ///     
        ///     // Export industry suppression counts
        ///     
        ///     // Create line
        ///     StringBuilder current_entry = new StringBuilder();
        ///     
        ///     for (int i = 0; i :lt; file_suppressions.Length; i++)
        ///     {
        ///         current_entry.Append(file_suppressions[i] + ",");
        ///     }
        ///     
        ///     if (File.Exists(OutputDirectory() + "industry_suppressions.csv"))
        ///     {
        ///         // Write the current line
        ///         using (StreamWriter entry_output = new StreamWriter(new FileStream(OutputDirectory() + "industry_suppressions.csv", FileMode.Append, FileAccess.Write, FileShare.Read)))
        ///         {
        ///             // Write line to the csv file
        ///             entry_output.WriteLine(current_entry);
        ///         }
        ///     }
        ///     else
        ///     {
        ///         // Get file heading
        ///         FileHeadings is_headings = new FileHeadings();
        ///               
        ///         // Write the current line
        ///         using (StreamWriter entry_output = new StreamWriter(new FileStream(OutputDirectory() + "industry_suppressions.csv", FileMode.Append, FileAccess.Write, FileShare.Read)))
        ///         {
        ///             // Write line to the csv file
        ///             entry_output.WriteLine(is_headings.GetSuppressionHeadings());
        ///             entry_output.WriteLine(current_entry);
        ///         }
        ///     }
        /// </code>
        /// </example>
        private void ExportIndustrySuppressions()
        { 
            // IS file path
            string file_path = RootOnServer() + @"\Industry Suppression Drops.csv";
            ArrayList is_csv = new ArrayList();
            ArrayList file_names = new ArrayList();

            // Set columns
            int filenames = 5;
            int supphit = 3;
            int line_count = 1;

            // Open CSV
            StreamReader read_csv = new StreamReader(new FileStream(file_path,FileMode.Open,FileAccess.Read)); 

            // Read CSV to ArrayList
            do
            {
                // Populate array [one line per index]
                is_csv.Add(read_csv.ReadLine());
   
            } while (read_csv.EndOfStream == false);

            // Close CSV
            read_csv.Close();

            // Cycle through array of industry suppressions
            foreach (string line in is_csv)
            {
                // Break line into array
                string[] record = line.Split(new Char[]{','});

                switch (line_count)
                {
                    case 1:
                        // Ignore file heading
                        // Iterate line count
                        line_count++;
                        break;
                    case 2:
                        // Store first filename
                        file_names.Add(record[filenames]);
                        // Iterate line count
                        line_count++;
                        break;
                    default:

                        // Set flag
                        bool exists = false;

                        // Check back to added filenames
                        foreach (string filename in file_names)
                        {
                            // if filename already exists then flag
                            if (filename == record[filenames])
                            {
                                exists = true;
                            }
                        }

                        // Check flag
                        if (exists == false)
                        {
                            // Add to list of filenames
                            file_names.Add(record[filenames]);
                        }

                        // Iterate line count
                        line_count++;
                        break;
                }  
            }

            // Loop through each file
            foreach (string file in file_names)
            {
                // Create array
                string[] file_suppressions = new string[15];

                // Store file name in multi-dimensional array
                file_suppressions[0] = file;
                int ncoa_redir = 0, gas_react = 0, ncoa_dec = 0, tbr = 0, ndr = 0, morta = 0, mps_dec = 0, ncoa = 0, gas = 0, mps = 0;

                // Match back file with is array
                foreach (string line in is_csv)
	            {
                    // Split hit into parts
                    string[] parts = line.Split(new Char[] { ',' });

                    // Check filename matches
                    if (file == parts[filenames])
                    {
                        // Record input vol
                        file_suppressions[1] = parts[4].ToString();
                        file_suppressions[1] = file_suppressions[1].Replace("\"", "");

                        // Switch between hits
                        switch (parts[supphit])
                        {
                            case "\"NCOA REDIR\"":
                                ncoa_redir++;
                                file_suppressions[2] = ncoa_redir.ToString();
                                break;
                            case "\"GAS REACT\"":
                                gas_react++;
                                file_suppressions[3] = gas_react.ToString();
                                break;
                            case "\"NCOA Deceased\"":
                                ncoa_dec++;
                                file_suppressions[4] = ncoa_dec.ToString();
                                break;
                            case "\"TBR\"":
                                tbr++;
                                file_suppressions[5] = tbr.ToString();
                                break;
                            case "\"NDR\"":
                                ndr++;
                                file_suppressions[6] = ndr.ToString();
                                break;
                            case "\"MORTA\"":
                                morta++;
                                file_suppressions[7] = morta.ToString();
                                break;
                            case "\"MPS Deceased\"":
                                mps_dec++;
                                file_suppressions[8] = mps_dec.ToString();
                                break;
                            case "\"NCOA\"":
                                ncoa++;
                                file_suppressions[9] = ncoa.ToString();
                                break;
                            case "\"GAS\"":
                                gas++;
                                file_suppressions[10] = gas.ToString();
                                break;
                            case "\"MPS\"":
                                mps++;
                                file_suppressions[11] = mps.ToString();
                                break;
                            default:
                                break;
                        }
                    }
	            }

                // Totals drops
                int totals = ncoa_dec + tbr + ndr + morta + mps_dec + ncoa + gas + mps;
                file_suppressions[12] = totals.ToString();

                // File output
                file_suppressions[13] = Convert.ToString(Convert.ToInt64(file_suppressions[1]) - totals);
 
                file_suppressions[0] = file_suppressions[0].Replace("\"", "");

                // Total hits
                file_suppressions[14] = Convert.ToString(Convert.ToInt64(file_suppressions[12]) + Convert.ToInt64(file_suppressions[2]) + Convert.ToInt64(file_suppressions[3]));               

                // Export industry suppression counts

                // Create line
                StringBuilder current_entry = new StringBuilder();

                for (int i = 0; i < file_suppressions.Length; i++)
                {
                        current_entry.Append(file_suppressions[i] + ",");
                }

                if (File.Exists(OutputDirectory() + "industry_suppressions.csv"))
                {
                    // Write the current line
                    using (StreamWriter entry_output = new StreamWriter(new FileStream(OutputDirectory() + "industry_suppressions.csv", FileMode.Append, FileAccess.Write, FileShare.Read)))
                    {
                        // Write line to the csv file
                        entry_output.WriteLine(current_entry);
                    }
 
                }
                else
                {
                    // Get file heading
                    FileHeadings is_headings = new FileHeadings();
                   
                    // Write the current line
                    using (StreamWriter entry_output = new StreamWriter(new FileStream(OutputDirectory() + "industry_suppressions.csv", FileMode.Append, FileAccess.Write, FileShare.Read)))
                    {
                        // Write line to the csv file
                        entry_output.WriteLine(is_headings.GetSuppressionHeadings());
                        entry_output.WriteLine(current_entry);

                    }
                }
            }
        }

        /// <summary>
        /// Presents the industry suppression forename matching form to the user.
        /// </summary>
        /// <returns>the forename matching level as a string</returns>
        /// <example>
        /// <code>
        ///     // Set text in form
        ///     Window_IndustryMatching window = new Window_IndustryMatching();
        ///     
        ///     // Show form to user
        ///     window.ShowDialog();
        ///     
        ///     // Return array
        ///     return window.GetIndustryLevel();
        /// </code>
        /// </example>
        private string IndustryLevel()
        {
            // Set text in form
            Window_IndustryMatching window = new Window_IndustryMatching();

            // Show form to user
            window.ShowDialog();

            // Return array
            return window.GetIndustryLevel();
        }

        // Total Functions

        /// <summary>
        /// Reads through the previously output industry_suppressions.csv and calculates
        /// the sum of each column containing industry suppression list hits.
        /// </summary>
        /// <returns>a comma delimited string of totals</returns>
        /// <example><code>
        ///     // Declare constants
        ///     string file_path = OutputDirectory() + "industry_suppressions.csv";
        ///     StreamReader is_csv = new StreamReader(file_path);
        ///     
        ///     // Declare variables
        ///     string current_line;
        ///     
        ///     // Declare suppression counter
        ///     int[] suppression_counts = new int[15];
        ///     
        ///     // Open file and count suppressions
        ///     do
        ///     {
        ///         // Record line
        ///         current_line = is_csv.ReadLine();
        ///     
        ///         // Split line into array of parts
        ///         string[] line = current_line.Split(new char[] { ',' });
        ///     
        ///         // Ignore heading line and blanks
        ///         if ((line[0] != "File") :amp;:amp; (line[0] != ""))
        ///         {
        ///             // Add values to the totals array
        ///             for (int i = 1; i :lt; (line.Length - 1); i++)
        ///             {
        ///                 // Default value if blank
        ///                 if (line[i] == "")
        ///                 {
        ///                     line[i] = "0";
        ///                 }
        ///     
        ///                 // Add values - minus index one as we're not holding the file name in the array
        ///                 suppression_counts[i - 1] = suppression_counts[i - 1] + Convert.ToInt32(line[i]);
        ///             }
        ///         }
        ///     } while (is_csv.EndOfStream == false);
        ///     
        ///     // Reset reader
        ///     is_csv.Close();
        ///     is_csv.Dispose();
        ///     
        ///     // Build string to print
        ///     StringBuilder totals_output = new StringBuilder();
        ///     
        ///     // Add first field
        ///     totals_output.Append("Totals,");
        ///     
        ///     // Loop through array and build string
        ///     foreach (int total in suppression_counts)
        ///     {
        ///         totals_output.Append(total + ",");
        ///     }
        ///     
        ///     // Write totals to Industry Suppressions csv - true defines appending opposed to overwrite
        ///     StreamWriter is_output = new StreamWriter(file_path, true);
        ///     
        ///     // Append blank line
        ///     is_output.WriteLine();
        ///     
        ///     // Append totals line
        ///     is_output.WriteLine(totals_output);
        ///     
        ///     // Close writer
        ///     is_output.Close();
        ///     is_output.Dispose();
        ///     
        ///     // Return totals
        ///     return totals_output.ToString();</code></example>
        private string CalculateIndustryTotals()
        {
            // Declare constants
            string file_path = OutputDirectory() + "industry_suppressions.csv";
            StreamReader is_csv = new StreamReader(file_path);
            
            // Declare variables
            string current_line;

            // Declare suppression counter
            int[] suppression_counts = new int[15];

            // Open file and count suppressions
            do
            {
                // Record line
                current_line = is_csv.ReadLine();

                // Split line into array of parts
                string[] line = current_line.Split(new char[] { ',' });

                // Ignore heading line and blanks
                if ((line[0] != "File") && (line[0] != ""))
                {
                    // Add values to the totals array
                    for (int i = 1; i < (line.Length - 1); i++)
                    {
                        // Default value if blank
                        if (line[i] == "")
                        {
                            line[i] = "0";
                        }

                        // Add values - minus index one as we're not holding the file name in the array
                        suppression_counts[i - 1] = suppression_counts[i - 1] + Convert.ToInt32(line[i]);
                    }
                }

            } while (is_csv.EndOfStream == false);

            // Reset reader
            is_csv.Close();
            is_csv.Dispose();
            
            // Build string to print
            StringBuilder totals_output = new StringBuilder();

            // Add first field
            totals_output.Append("Totals,");

            // Loop through array and build string
            foreach (int total in suppression_counts)
            {
                totals_output.Append(total + ",");
            }

            // Write totals to Industry Suppressions csv - true defines appending opposed to overwrite
            StreamWriter is_output = new StreamWriter(file_path,true);

            // Append blank line
            is_output.WriteLine();

            // Append totals line
            is_output.WriteLine(totals_output);

            // Close writer
            is_output.Close();
            is_output.Dispose();

            // Return totals
            return totals_output.ToString();

        }

        /// <summary>
        /// Outputs the costs for industry flags and deletes from the totals received as a parameter.
        /// </summary>
        /// <param name="calculated_totals">totals as a string delimited by comma.</param>
        /// <example>
        /// <code>
        ///     // Declare variables
        ///     string file_path = OutputDirectory() + "industry_costs.csv";
        ///     StringBuilder flag_costs = new StringBuilder();
        ///     StringBuilder delete_costs = new StringBuilder();
        ///     
        ///     // Retrieve file headings
        ///     FileHeadings file_headings = new FileHeadings();
        ///     
        ///     // Break totals into array
        ///     string[] totals = calculated_totals.Split(new char[] { ',' });
        ///     
        ///     // Calculate flag costs
        ///     for (int i = 0; i :lt; totals.Length; i++)
        ///     {
        ///             // Default if blank
        ///             if (totals[i] == "")
        ///             {
        ///                 totals[i] = "0.00";
        ///             }
        ///             
        ///             // Ignore filename
        ///             switch (i)
        ///             {
        ///                 case 0:
        ///                     flag_costs.Append("Flag,,");
        ///                     break;
        ///                case 1:
        ///                     // Ignore input totals
        ///                      break;
        ///                 case 14:
        ///                     // Ignore input totals
        ///                     break;
        ///                 case 12:
        ///                     // Append totals
        ///                     double total = Convert.ToDouble(totals[14]) * 0.6;
        ///                     flag_costs.Append(total.ToString() + ",");
        ///                     break;
        ///                 case 13:
        ///                     // Ignore output totals
        ///                     break;
        ///                 case 8:
        ///                     flag_costs.Append("FOC,");
        ///                     break;
        ///                 case 11:
        ///                     flag_costs.Append("FOC,");
        ///                     break;
        ///                 case 15:
        ///                     // Ignore output totals
        ///                     break;
        ///                 case 16:
        ///                     // Ignore
        ///                     break;
        ///                 default:
        ///                     // Calculate flag cost and append to the totals string
        ///                     double calc = Convert.ToDouble(totals[i]) * 0.6;
        ///                     flag_costs.Append(calc.ToString() + ",");
        ///                     break;
        ///             }
        ///      }
        ///     
        ///     
        ///     // Calculate delete costs
        ///     for (int i = 0; i :lt; totals.Length; i++)
        ///     {
        ///             // Default if blank
        ///             if (totals[i] == "")
        ///             {
        ///                 totals[i] = "0.00";
        ///             }
        ///     
        ///             // Ignore filename
        ///             switch (i)
        ///             {
        ///                 case 0:
        ///                     delete_costs.Append("Delete,,");
        ///                     break;
        ///                 case 1:
        ///                     // Ignore input totals
        ///                     break;
        ///                 case 2:
        ///                     // NCOA Redirects
        ///                     delete_costs.Append("REDIRECT,");
        ///                     break;
        ///                 case 3:
        ///                     // GAS Reactive
        ///                     delete_costs.Append("ONLY,");
        ///                     break;
        ///                 case 12:
        ///                     // Append totals
        ///                     double total = Convert.ToDouble(totals[14]) * 0.2;
        ///                     delete_costs.Append(total.ToString() + ",");
        ///                     break;
        ///                 case 13:
        ///                     // Ignore input totals
        ///                     break;
        ///                 case 14:
        ///                     // Ignore input totals
        ///                     break;
        ///                 case 15:
        ///                     // Ignore output totals
        ///                     break;
        ///                 case 8:
        ///                     delete_costs.Append("FOC,");
        ///                     break;
        ///                 case 11:
        ///                     delete_costs.Append("FOC,");
        ///                     break;
        ///                case 16:
        ///                     // Ignore
        ///                     break;
        ///                default:
        ///                     // Calculate flag cost and append to the totals string
        ///                     double calc = Convert.ToDouble(totals[i]) * 0.2;
        ///                     delete_costs.Append(calc.ToString() + ",");
        ///                     break;
        ///             }
        ///     }
        ///     
        ///     // Write output totals to file
        ///     StreamWriter cost_output = new StreamWriter(file_path, true);
        ///     
        ///     // Write heading
        ///     cost_output.WriteLine(file_headings.GetCostHeadings());
        ///     // Write flags
        ///     cost_output.WriteLine(flag_costs.ToString());
        ///     // Write deletes
        ///     cost_output.WriteLine(delete_costs.ToString());
        ///     
        ///     // Close writer
        ///     cost_output.Close();
        ///     cost_output.Dispose();
        /// </code>
        /// </example>
        private void ExportSuppressionCosts(string calculated_totals)
        {
            // Declare variables
            string file_path = OutputDirectory() + "industry_costs.csv";
            StringBuilder flag_costs = new StringBuilder();
            StringBuilder delete_costs = new StringBuilder();

            // Retrieve file headings
            FileHeadings file_headings = new FileHeadings();

            // Break totals into array
            string[] totals = calculated_totals.Split(new char[] { ',' });

            // Calculate flag costs
            for (int i = 0; i < totals.Length; i++)
            {
                // Default if blank
                if (totals[i] == "")
                {
                    totals[i] = "0.00";
                }
                
                // Ignore filename
                switch (i)
                {
                    case 0:
                        flag_costs.Append("Flag,,");
                        break;
                    case 1:
                        // Ignore input totals
                        break;
                    case 14:
                        // Ignore input totals
                        break;
                    case 12:
                        // Append totals
                        double total = Convert.ToDouble(totals[14]) * 0.6;
                        flag_costs.Append(total.ToString() + ",");
                        break;
                    case 13:
                        // Ignore output totals
                        break;
                    case 8:
                        flag_costs.Append("FOC,");
                        break;
                    case 11:
                        flag_costs.Append("FOC,");
                        break;
                    case 15:
                        // Ignore output totals
                        break;
                    case 16:
                        // Ignore
                        break;
                    default:
                        // Calculate flag cost and append to the totals string
                        double calc = Convert.ToDouble(totals[i]) * 0.6;
                        flag_costs.Append(calc.ToString() + ",");
                        break;
                }
            }


            // Calculate delete costs
            for (int i = 0; i < totals.Length; i++)
            {
                // Default if blank
                if (totals[i] == "")
                {
                    totals[i] = "0.00";
                }

                // Ignore filename
                switch (i)
                {
                    case 0:
                        delete_costs.Append("Delete,,");
                        break;
                    case 1:
                        // Ignore input totals
                        break;
                    case 2:
                        // Qin movers
                        delete_costs.Append("REDIRECT,");
                        break;
                    case 3:
                        // NCOA Redirects
                        delete_costs.Append("ONLY,");
                        break;
                    case 12:
                        // Append totals
                        double total = Convert.ToDouble(totals[14]) * 0.2;
                        delete_costs.Append(total.ToString() + ",");
                        break;
                    case 13:
                        // Ignore input totals
                        break;
                    case 14:
                        // Ignore input totals
                        break;
                    case 15:
                        // Ignore output totals
                        break;
                    case 8:
                        delete_costs.Append("FOC,");
                        break;
                    case 11:
                        delete_costs.Append("FOC,");
                        break;
                    case 16:
                        // Ignore
                        break;
                    default:
                        // Calculate flag cost and append to the totals string
                        double calc = Convert.ToDouble(totals[i]) * 0.2;
                        delete_costs.Append(calc.ToString() + ",");
                        break;
                }
            }

            // Write output totals to file
            StreamWriter cost_output = new StreamWriter(file_path, true);

            // Write heading
            cost_output.WriteLine(file_headings.GetCostHeadings());
            // Write flags
            cost_output.WriteLine(flag_costs.ToString());
            // Write deletes
            cost_output.WriteLine(delete_costs.ToString());

            // Close writer
            cost_output.Close();
            cost_output.Dispose();

        }

        /// <summary>
        /// Reads through the previously created download stats csv and produces the totals for each column.
        /// </summary>
        /// <example>
        /// <code>
        ///     // Store download stats path
        ///     string download_path = OutputDirectory() + "download_stats.csv";
        ///     
        ///     // Open download stats file
        ///     StreamReader reader = new StreamReader(download_path);
        ///     
        ///     // Skip first line
        ///     reader.ReadLine();
        ///     
        ///     // Create totals array
        ///     int[] totals = new int[44];
        ///     
        ///     // Count number of mailing files in stats file
        ///     do
        ///     {
        ///         // Split line into array
        ///         string[] line = reader.ReadLine().Split(new char[] { ',' });
        ///     
        ///         // Read line into array
        ///         for (int i = 1; i :lt; line.Length; i++)
        ///         {
        ///             try
        ///             {
        ///                 if (String.IsNullOrWhiteSpace(line[i]))
        ///                 {
        ///                     totals[i] = 0;
        ///                 }
        ///                 else
        ///                 {
        ///                     // Calculate totals
        ///                     totals[i] = Convert.ToInt32(totals[i] + Convert.ToInt32(line[i]));
        ///                 }
        ///             }
        ///             catch (Exception ex)
        ///             {
        ///                 MessageBox.Show(ex.Message + "; there is a problem calculating download totals.");
        ///             }
        ///         }
        ///     } while (reader.EndOfStream == false);
        ///     
        ///     // Close the reader
        ///     reader.Dispose();
        ///     
        ///     // Save download stats file
        ///     StringBuilder totals_line = new StringBuilder();
        ///     
        ///     // Loop through the totals array
        ///     for (int i = 0; i :lt; (totals.Length-1); i++)
        ///     {
        ///         totals_line.Append((i == 0) ? "Total," : totals[i] + ",");
        ///     }
        ///     
        ///     // Write to download stats
        ///     StreamWriter writer = new StreamWriter(download_path, true);
        ///     
        ///     // Write totals
        ///     writer.WriteLine(totals_line);
        ///     
        ///     // Close file
        ///     writer.Dispose();
        /// </code>
        /// </example>
        private void CalculateDownloadTotals()
        {
            // Store download stats path
            string download_path = OutputDirectory() + "download_stats.csv";

            // Open download stats file
            StreamReader reader = new StreamReader(download_path);

            // Skip first line
            reader.ReadLine();

            // Create totals array
            int[] totals = new int[44];

            // Count number of mailing files in stats file
            do
            {
                // Split line into array
                string[] line = reader.ReadLine().Split(new char[] {','});

                // Read line into array
                for (int i = 1; i < line.Length; i++)
                {
                    try
                    {
                        if (String.IsNullOrWhiteSpace(line[i]))
                        {
                            totals[i] = 0;
                        }
                        else
                        {
                            totals[i] = Convert.ToInt32(totals[i] + Convert.ToInt32(line[i]));
                        } 
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message + "; there is a problem calculating download totals.");
                    }// Calculate totals
                   
                }

            } while (reader.EndOfStream == false);

            // Close the reader
            reader.Dispose();

            // Save download stats file
            StringBuilder totals_line = new StringBuilder();

            // Loop through the totals array
            for (int i = 0; i < (totals.Length-1); i++)
            {
                totals_line.Append((i == 0) ? "Total," : totals[i] + ",");
            }

            // Write to download stats
            StreamWriter writer = new StreamWriter(download_path,true);

            // Write totals
            writer.WriteLine(totals_line);

            // Close file
            writer.Dispose();

        }

        /// <summary>
        /// Reads through the previously created dedupe stats csv and produces the totals for each dedupe
        /// </summary>
        /// <example><code>
        ///     // Open dedupe statistics csv
        ///     StreamReader reader = new StreamReader(OutputDirectory() + "dedupe_statistics.csv");
        ///     StringBuilder filename_string = new StringBuilder();
        ///     
        ///     // Build string of file names - loop through file
        ///     do
        ///     {               
        ///         // Split current line into array
        ///         string[] line = reader.ReadLine().Split(new char[] { ',' });
        ///     
        ///         // Add filename on line to filename_string
        ///         try
        ///         {
        ///             // If line is populated then
        ///             if (line.Length > 1)
        ///             {
        ///                 // Trim whitespaces
        ///                 line[1] = line[1].Trim();
        ///                    
        ///                 // Store filename in array 
        ///                 if ((String.IsNullOrWhiteSpace(line[1])==false) &amp;&amp; (line[1] != "") &amp;&amp; (line[1] != "FILE_NAME"))
        ///                 {
        ///                     // Remove any non-original characters using FormatName
        ///                     string formatted_filename = FormatName(line[1]).Trim();
        ///     
        ///                     // Add to string [pipe delimited]
        ///                      filename_string.Append(formatted_filename + " | ");
        ///                 }
        ///             }
        ///          }
        ///         catch (Exception ex)
        ///         {
        ///             MessageBox.Show(ex.Message);
        ///         }
        ///     
        ///     } while (reader.EndOfStream == false);
        ///     
        ///     // Dispose reader
        ///     reader.Dispose();
        ///     
        ///     // Break string into array
        ///     string[] filename_array = filename_string.ToString().Split(new char[] { '|' });
        ///     
        ///     // Remove whitespace
        ///     for (int i = 0; i :lt; filename_array.Length; i++)
        ///     {
        ///         filename_array[i] = filename_array[i].Trim();
        ///     }
        ///     
        ///      // Remove duplicate filenames from array
        ///     filename_array = filename_array.Distinct().ToArray();
        ///     
        ///     // Loop through array of unique filenames
        ///     foreach (string filename in filename_array)
        ///     {
        ///           // Proceed if not blank
        ///           if (String.IsNullOrWhiteSpace(filename) == false)
        ///           {
        ///                 // Array of details for filename
        ///                 string[] file_stats = new string[12];
        ///     
        ///                 // Reopen reader
        ///                 StreamReader new_reader = new StreamReader(OutputDirectory() + "dedupe_statistics.csv");
        ///                 
        ///                 // Loop through CSV
        ///                 do
        ///                 {
        ///                     try
        ///                     {
        ///                         // Split current line into array
        ///                         string[] new_line = new_reader.ReadLine().Split(new char[] { ',' });
        ///                            
        ///                         // If line is populated then
        ///                         if (new_line.Length > 1)
        ///                        {
        ///                             // Trim whitespaces
        ///                             new_line[1] = new_line[1].Trim();
        ///     
        ///                             // Store filename in array 
        ///                             if ((String.IsNullOrWhiteSpace(new_line[1]) == false) :amp;:amp; (new_line[1] != "") :amp;:amp; (new_line[1] != "FILE_NAME"))
        ///                             {
        ///                                 // Format current file in CSV
        ///                                 string cur_filename = FormatName(new_line[1]);
        ///                                
        ///                                 // Check if filename matches
        ///                                 if (filename.Equals(cur_filename))
        ///                                 {
        ///                                     // Store filename
        ///                                     file_stats[0] = cur_filename;
        ///     
        ///                                     // Store List Code
        ///                                     file_stats[1] = new_line[2]; 
        ///     
        ///                                     // Store Priority
        ///                                     file_stats[2] = new_line[3]; 
        ///     
        ///                                     // Store input records - HIGHEST
        ///                                     if (Convert.ToInt32(new_line[4]) > Convert.ToInt32(file_stats[3]))
        ///                                     {
        ///                                         file_stats[3] = new_line[4]; 
        ///                                     }
        ///     
        ///                                     // Store suppressed records - SUM OF
        ///                                     file_stats[4] += new_line[5];
        ///     
        ///                                     // Store inter dupes - SUM OF
        ///                                     file_stats[5] += new_line[6];
        ///     
        ///                                     // Store intra dupes - SUM OF
        ///                                     file_stats[6] += new_line[7];
        ///     
        ///                                     // Store total drops - SUM OF
        ///                                     file_stats[7] += new_line[8];
        ///     
        ///                                     // Store uniques - SUM OF
        ///                                     file_stats[8] += new_line[9];
        ///     
        ///                                     // Store selected - SUM OF
        ///                                     file_stats[9] += new_line[10]; 
        ///     
        ///                                     // Store output records - LOWEST
        ///                                     if (Convert.ToInt32(new_line[11]) > Convert.ToInt32(file_stats[10]))
        ///                                     {
        ///                                         file_stats[10] = new_line[11];
        ///                                     }
        ///     
        ///                                     // Line to print
        ///                                     StringBuilder current_line = new StringBuilder();
        ///     
        ///                                     // Build line
        ///                                     foreach (string field in file_stats)
        ///                                     {
        ///                                         current_line.Append(field + ",");
        ///                                     }
        ///     
        ///                                     // Write heading if file doesn't exist
        ///                                     if (File.Exists(OutputDirectory() + "dedupe_totals.csv")==false)
        ///                                     {
        ///                                         // Open StreamWriter
        ///                                         StreamWriter output = new StreamWriter(OutputDirectory() + "dedupe_totals.csv");
        ///     
        ///                                         // Open headings
        ///                                         FileHeadings headings = new FileHeadings();
        ///     
        ///                                         // Write headings to file
        ///                                         output.WriteLine(headings.GetDedupeTotalHeadings());
        ///     
        ///                                         // Write line to a file
        ///                                         output.WriteLine(current_line.ToString());
        ///     
        ///                                         // Dispose output writer
        ///                                         output.Dispose();
        ///     
        ///                                     }
        ///                                     else
        ///                                     {
        ///                                         // Open StreamWriter
        ///                                         StreamWriter output = new StreamWriter(OutputDirectory() + "dedupe_totals.csv", true);
        ///     
        ///                                         // Write line to a file
        ///                                         output.WriteLine(current_line.ToString());
        ///     
        ///                                         // Dispose output writer
        ///                                         output.Dispose();
        ///                                     }
        ///                                 }
        ///                             }
        ///                         }
        ///                     }
        ///                         catch (Exception ex)
        ///                     {
        ///                         MessageBox.Show(ex.Message);
        ///                     }
        ///     
        ///             } while (new_reader.EndOfStream == false);
        ///         }
        ///     }
        /// </code></example>
        private void CalculateDedupeFileTotals()
        {
            // Open dedupe statistics csv
            StreamReader reader = new StreamReader(OutputDirectory() + "dedupe_statistics.csv");
            StringBuilder filename_string = new StringBuilder();

            // Build string of file names - loop through file
            do
            {               
                // Split current line into array
                string[] line = reader.ReadLine().Split(new char[] { ',' });

                // Add filename on line to filename_string
                try
                {
                    // If line is populated then
                    if (line.Length > 1)
                    {
                        // Trim whitespaces
                        line[1] = line[1].Trim();
                        
                        // Store filename in array 
                        if ((String.IsNullOrWhiteSpace(line[1])==false) && (line[1] != "") && (line[1] != "FILE_NAME"))
                        {
                            // Remove any non-original characters using FormatName
                            string formatted_filename = FormatName(line[1]).Trim();

                            // Add to string [pipe delimited]
                            filename_string.Append(formatted_filename + " | ");
                        }
                    }      
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

            } while (reader.EndOfStream == false);

            // Dispose reader
            reader.Dispose();

            // Break string into array
            string[] filename_array = filename_string.ToString().Split(new char[] { '|' });

            // Remove whitespace
            for (int i = 0; i < filename_array.Length; i++)
            {
                filename_array[i] = filename_array[i].Trim();
            }

            // Remove duplicate filenames from array
            filename_array = filename_array.Distinct().ToArray();

            // Loop through array of unique filenames
            foreach (string filename in filename_array)
            {
               // Proceed if not blank
                if (String.IsNullOrWhiteSpace(filename) == false)
                {
                    // Array of details for filename
                    string[] file_stats = new string[12];
                    
                    // Reopen reader
                    StreamReader new_reader = new StreamReader(OutputDirectory() + "dedupe_statistics.csv");
                    
                    // Loop through CSV
                    do
                    {
                        try
                        {
                            // Split current line into array
                            string[] new_line = new_reader.ReadLine().Split(new char[] { ',' });
                               
                            // If line is populated then
                            if (new_line.Length > 1)
                            {
                                // Trim whitespaces
                                new_line[1] = new_line[1].Trim();

                                // Store filename in array 
                                if ((String.IsNullOrWhiteSpace(new_line[1]) == false) && (new_line[1] != "") && (new_line[1] != "FILE_NAME"))
                                {
                                    // Format current file in CSV
                                    string cur_filename = FormatName(new_line[1]);
                                    
                                    // Check if filename matches
                                    if (filename.Equals(cur_filename))
                                    {
                                        // Store filename
                                        file_stats[0] = cur_filename;
 
                                        // Store List Code
                                        file_stats[1] = new_line[2]; 

                                        // Store Priority
                                        file_stats[2] = new_line[3]; 

                                        // Store input records - HIGHEST
                                        if (Convert.ToInt32(new_line[4]) > Convert.ToInt32(file_stats[3]))
                                        {
                                            file_stats[3] = new_line[4]; 
                                        } 

                                        // Store suppressed records - SUM OF
                                        file_stats[4] += new_line[5];

                                        // Store inter dupes - SUM OF
                                        file_stats[5] += new_line[6];

                                        // Store intra dupes - SUM OF
                                        file_stats[6] += new_line[7];

                                        // Store total drops - SUM OF
                                        file_stats[7] += new_line[8];

                                        // Store uniques - SUM OF
                                        file_stats[8] += new_line[9];

                                        // Store selected - SUM OF
                                        file_stats[9] += new_line[10]; 

                                        // Store output records - LOWEST
                                        if (Convert.ToInt32(new_line[11]) > Convert.ToInt32(file_stats[10]))
                                        {
                                            file_stats[10] = new_line[11];
                                        } 

                                        // Line to print
                                        StringBuilder current_line = new StringBuilder();

                                        // Build line
                                        foreach (string field in file_stats)
                                        {
                                            current_line.Append(field + ",");
                                        }

                                        // Write heading if file doesn't exist
                                        if (File.Exists(OutputDirectory() + "dedupe_totals.csv")==false)
                                        {
                                            // Open StreamWriter
                                            StreamWriter output = new StreamWriter(OutputDirectory() + "dedupe_totals.csv");

                                            // Open headings
                                            FileHeadings headings = new FileHeadings();

                                            // Write headings to file
                                            output.WriteLine(headings.GetDedupeTotalHeadings());

                                            // Write line to a file
                                            output.WriteLine(current_line.ToString());

                                            // Dispose output writer
                                            output.Dispose();

                                        }
                                        else
                                        {
                                            // Open StreamWriter
                                            StreamWriter output = new StreamWriter(OutputDirectory() + "dedupe_totals.csv", true);

                                            // Write line to a file
                                            output.WriteLine(current_line.ToString());

                                             // Dispose output writer
                                            output.Dispose();
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }

                    } while (new_reader.EndOfStream == false);
                }
            }
        }

        /// <summary>
        /// Reads through the previously created dedupe totals csv and produces the totals for each file
        /// </summary>
        /// <example>
        /// <code>
        ///     // Store dedupe totals path
        ///     string dedupe_path = OutputDirectory() + "dedupe_totals.csv";
        ///     
        ///     // Open dedupe totals file
        ///     StreamReader reader = new StreamReader(dedupe_path);
        ///     
        ///     // Skip first line
        ///     reader.ReadLine();
        ///     
        ///     // Create totals array
        ///     int[] totals = new int[12];
        ///     
        ///     // Count number of mailing files in dedupe totals file
        ///     do
        ///     {
        ///         // Split line into array
        ///         string[] line = reader.ReadLine().Split(new char[] { ',' });
        ///     
        ///         // Read line info into array
        ///         for (int i = 0; i :lt; line.Length - 2; i++)
        ///         {
        ///            try
        ///            {
        ///                if (i > 2)
        ///                {
        ///                     if (String.IsNullOrWhiteSpace(line[i]))
        ///                     {
        ///                         totals[i] = 0;
        ///                     }
        ///                     else
        ///                     {
        ///                         // Calculate totals                        
        ///                         totals[i] = Convert.ToInt32(totals[i] + Convert.ToInt32(line[i]));
        ///                     }
        ///                 }
        ///             }
        ///             catch (Exception ex)
        ///             {
        ///                 MessageBox.Show(ex.Message + " there is a problem calculating dedupe totals.");
        ///             } 
        ///         }
        ///     
        ///     } while (reader.EndOfStream == false);
        ///     
        ///     // Close the reader
        ///     reader.Dispose();
        ///     
        ///     // Save dedupe totals
        ///     StringBuilder totals_line = new StringBuilder();
        ///     
        ///     // Loop through the totals array
        ///     for (int i = 0; i :lt; (totals.Length - 1); i++)
        ///     {
        ///         totals_line.Append((i == 0) ? "Total," : totals[i] + ",");
        ///     }
        ///     
        ///     // Write to dedupe totals
        ///     StreamWriter writer = new StreamWriter(dedupe_path, true);
        ///     
        ///     // Write totals
        ///     writer.WriteLine();
        ///     writer.WriteLine(totals_line);
        ///     
        ///     // Close file
        ///     writer.Dispose();
        /// </code>
        /// </example>
        private void CalculateFinalDedupeTotals()
        {
            // Store dedupe totals path
            string dedupe_path = OutputDirectory() + "dedupe_totals.csv";

            // Open dedupe totals file
            StreamReader reader = new StreamReader(dedupe_path);

            // Skip first line
            reader.ReadLine();

            // Create totals array
            int[] totals = new int[12];

            // Count number of mailing files in dedupe totals file
            do
            {
                // Split line into array
                string[] line = reader.ReadLine().Split(new char[] { ',' });

                // Read line info into array
                for (int i = 0; i < line.Length - 2; i++)
                {
                    try
                    {
                        if (i > 2)
                        {
                            if (String.IsNullOrWhiteSpace(line[i]))
                            {
                                totals[i] = 0;
                            }
                            else
                            {
                                totals[i] = Convert.ToInt32(totals[i] + Convert.ToInt32(line[i]));
                            }
                        }   
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message + " there is a problem calculating dedupe totals.");
                    } // Calculate totals

                }

            } while (reader.EndOfStream == false);

            // Close the reader
            reader.Dispose();

            // Save dedupe totals
            StringBuilder totals_line = new StringBuilder();

            // Loop through the totals array
            for (int i = 0; i < (totals.Length - 1); i++)
            {
                totals_line.Append((i == 0) ? "Total," : totals[i] + ",");
            }

            // Write to dedupe totals
            StreamWriter writer = new StreamWriter(dedupe_path, true);

            // Write totals
            writer.WriteLine();
            writer.WriteLine(totals_line);

            // Close file
            writer.Dispose();
        }

        /// <summary>
        /// Exports the basic information of the selected job for use in the report
        /// </summary>
        /// <example>
        /// <code>
        ///     // Job details
        ///     string job_number = JobDescription.Content.ToString().Substring(0, 5);
        ///     string job_description = JobDescription.Content.ToString().Substring(6, JobDescription.Content.ToString().Length - 6);
        ///     string date = DateTime.Now.ToShortDateString();
        ///     
        ///     // Check file currently exists
        ///     if (File.Exists(OutputDirectory() + "job_information.csv"))
        ///     {
        ///         // Delete old job information
        ///         File.Delete(OutputDirectory() + "job_information.csv");
        ///     }
        ///     
        ///     // Re-create file and write job information
        ///     StreamWriter writer = new StreamWriter(OutputDirectory() + "job_information.csv");
        ///     
        ///     // Write headings
        ///     writer.WriteLine("JOB_NUMBER,JOB_DESCRIPTION,DATE");
        ///     
        ///     // Write info line
        ///     writer.WriteLine(job_number + "," + job_description + "," + date);
        ///     
        ///     // Close writer
        ///     writer.Dispose();
        /// </code>
        /// </example>
        private void JobInformation()
        {
            // Job details
            string job_number = JobDescription.Content.ToString().Substring(0,5);
            string job_description = JobDescription.Content.ToString().Substring(6, JobDescription.Content.ToString().Length - 6);
            string date = DateTime.Now.ToShortDateString();

            // Check file currently exists
            if (File.Exists(OutputDirectory() + "job_information.csv"))
            {
                // Delete old job information
                File.Delete(OutputDirectory() + "job_information.csv");
            }

            // Re-create file and write job information
            StreamWriter writer = new StreamWriter(OutputDirectory() + "job_information.csv");

            // Write headings
            writer.WriteLine("JOB_NUMBER,JOB_DESCRIPTION,DATE");

            // Write info line
            writer.WriteLine(job_number + "," + job_description + "," + date);

            // Close writer
            writer.Dispose();

        }

        // Report Functions

        private void CreateHealthcheckReport()
        {
            // Create paths
            string jobNumber = JobDescription.Content.ToString().Substring(0, 5);
            string zipName = OutputDirectory() + jobNumber + " Healthcheck Report.zip";
            string batch = OutputDirectory() + "CreateReport.bat";
            string wfd = "";
            CreateReport.IsEnabled = false;

            // Determine whether dedupe, industry suppressions exist
            bool dedupesExist = (File.Exists(OutputDirectory() + "dedupe_statistics.csv")) ? true : false;
            bool suppsExist = (File.Exists(OutputDirectory() + "industry_suppressions.csv")) ? true : false;

            // Determine WFD to use
            if (dedupesExist && suppsExist)
            {
                wfd = "WFD_DEFAULT.wfd";
            }
            else if (dedupesExist && !suppsExist)
            {
                wfd = "WFD_NO_SUPPRESSIONS.wfd";
            }
            else if (!dedupesExist && suppsExist)
            {
                wfd = "WFD_NO_DEDUPE.wfd";
            }
            else if (!dedupesExist && !suppsExist)
            {
                wfd = "WFD_DOWNLOAD_ONLY.wfd";
            }

            // If single file, then append _single to switch batch type
            wfd = (fileCount == 1) ? wfd.Replace(".wfd", "_1.wfd") : wfd;
            TextSummary.AppendText(Environment.NewLine + " -   Healthcheck type determined." + Environment.NewLine);
            UsingForms.Application.DoEvents();

            // Create batch file from template to output folder, and start process
            using (StreamWriter batchWriter = new StreamWriter(batch))
            {
                batchWriter.WriteLine("REM ----- Check if SOURCE folder exists");
                batchWriter.WriteLine("");
                batchWriter.WriteLine("if not exist \"%~dp0Source\\\" mkdir \"%~dp0Source\\\" ");
                batchWriter.WriteLine("");
                batchWriter.WriteLine("REM ----- Copy WFD from resource location to this directory");
                batchWriter.WriteLine("");
                batchWriter.WriteLine("xcopy /s/y \"" + ReadInDirectories()[2] + "" + wfd + "\" \"%~dp0\"");
                batchWriter.WriteLine("");
                batchWriter.WriteLine("REM ----- Run the WFD output to create the PDF report");
                batchWriter.WriteLine("");
                batchWriter.WriteLine("\"C:\\Program Files (x86)\\GMC\\Inspire Designer\\PNetTC.exe\" \"%~dp0" + wfd + "\" -e PDF -f \"%~dp0Healthcheck Report.pdf\" -o Output1");
                batchWriter.WriteLine("");
                batchWriter.WriteLine("REM Move used files into SOURCE folder");
                batchWriter.WriteLine("");
                batchWriter.WriteLine("move \"%~dp0*.csv\" \"%~dp0Source\\\"");
                batchWriter.WriteLine("move \"%~dp0*.txt\" \"%~dp0Source\\\"");
                batchWriter.WriteLine("");
                batchWriter.WriteLine("REM Move dedupe drops matrix pdfs to current folder");
                batchWriter.WriteLine("");
                batchWriter.WriteLine("move \"%~dp0Matrices\\*.pdf\" \"%~dp0\"");
                batchWriter.WriteLine("");
                batchWriter.WriteLine("REM Delete WFD");
                batchWriter.WriteLine("");
                batchWriter.WriteLine("del \"%~dp0" + wfd + "\"");
            }

            TextSummary.AppendText(" -   Report creation started." + Environment.NewLine);
            UsingForms.Application.DoEvents();
            Process proc = Process.Start(batch);
            proc.WaitForExit();

            TextSummary.AppendText(" -   Report creation  completed." + Environment.NewLine);
            UsingForms.Application.DoEvents();

            // Remove batch script
            File.Delete(batch);

            if (!File.Exists(OutputDirectory() + "Healthcheck Report.pdf"))
            {
                // Wait ten seconds
                System.Threading.Thread.Sleep(10000);

                // If file still doesn't exist then alert user
                if (!File.Exists(OutputDirectory() + "Healthcheck Report.pdf"))
                {
                    TextSummary.AppendText(" -   Report creation failed. Investigate, or get James to." + Environment.NewLine);
                }
                else
                {
                    TextSummary.AppendText(" -   Report created." + Environment.NewLine);
                    File.Copy(OutputDirectory() + "Healthcheck Report.pdf", OutputDirectory() + jobNumber + " Healthcheck Report.pdf", true);
                    File.Delete(OutputDirectory() + "Healthcheck Report.pdf");
                }
            }
            else
            {
                TextSummary.AppendText(" -   Report created." + Environment.NewLine);
                File.Copy(OutputDirectory() + "Healthcheck Report.pdf", OutputDirectory() + jobNumber + " Healthcheck Report.pdf", true);
                File.Delete(OutputDirectory() + "Healthcheck Report.pdf");
            }
        }

        /// <summary>
        /// Deletes all previous zip files in the output directory, before proceeding to ask the user for a valid password.
        /// Some basic validation is performed on the password, if it qualifies, the user is no longer asked, and the method 
        /// then creates a zip file applying the password, and compressing the healthcheck report pdf and the Matrices folder
        /// containing each dedupe drops matrix.
        /// </summary>
        /// <example>
        /// <code>
        ///     // Create paths
        ///     string jobNumber = JobDescription.Content.ToString().Substring(0, 5);
        ///     string zipName = OutputDirectory() + jobNumber + " Healthcheck Report.zip";
        ///     string pdfName = zipName.Replace(".zip", ".pdf");
        ///     string zipPassword = "";
        ///     bool qualifies = false;
        ///     int numeric;
        /// 
        ///     // Delete previous .zip files
        ///     foreach (string file in Directory.GetFiles(OutputDirectory()))
        ///     {
        ///         if (System.IO.Path.GetExtension(file) == ".zip")
        ///         {
        ///             File.Delete(file);
        ///         }
        ///     }
        /// 
        ///     // Retrieve password for the .zip file
        ///     do
        ///     {
        ///         zipPassword = VBInteraction.InputBox("Please enter a password for the .ZIP - '99999PMDM'", "Enter .ZIP Password");
        /// 
        ///         if (zipPassword.Length >= 9)
        ///         {
        ///             if (int.TryParse(zipPassword.Substring(0, 5), out numeric))
        ///             {
        ///                 qualifies = true;
        ///             }
        ///             else
        ///             {
        ///                  MessageBox.Show("Correct length, wrong format. Use '99999PMDM'.");
        ///             }
        ///         }
        ///         else
        ///         {
        ///             MessageBox.Show("Password should be at least 9 characters long. Use '99999PMDM'");
        ///         }
        /// 
        ///      } while (!qualifies);
        /// 
        ///     // Create zip file
        ///     using (ZipFile zip = new ZipFile())
        ///     {
        ///         // Add password to the zip
        ///         zip.Password = zipPassword;
        ///         // Add the healthcheck report pdf to the archive
        ///         zip.AddFile(pdfName,"");
        ///         // Add the matrices folder to the archive;
        ///         zip.AddDirectoryByName("Matrices");
        ///         zip.AddDirectory(OutputDirectory() + "Matrices","Matrices");
        ///         // Save the zip
        ///         zip.Save(zipName);    
        ///     }
        /// </code>
        /// </example>
        private void CompressHealthcheckReport()
        {
            // Create paths
            string jobNumber = JobDescription.Content.ToString().Substring(0, 5);
            string zipName = OutputDirectory() + jobNumber + " Healthcheck Report.zip";
            string pdfName = zipName.Replace(".zip", ".pdf");
            string zipPassword = "";
            bool qualifies = false;
            int numeric;

            // Delete previous .zip files
            foreach (string file in Directory.GetFiles(OutputDirectory()))
            {
                if (System.IO.Path.GetExtension(file) == ".zip")
                {
                    File.Delete(file);
                }
            }

            // Retrieve password for the .zip file
            do
            {
                zipPassword = VBInteraction.InputBox("Please enter a password for the .ZIP - '99999PMDM'", "Enter .ZIP Password");

                if (zipPassword.Length >= 9)
                {
                    if (int.TryParse(zipPassword.Substring(0, 5), out numeric))
                    {
                        qualifies = true;
                    }
                    else
                    {
                        MessageBox.Show("Correct length, wrong format. Use '99999PMDM'.");
                    }
                }
                else
                {
                    MessageBox.Show("Password should be at least 9 characters long. Use '99999PMDM'");
                }

            } while (!qualifies);

            // Create zip file
            using (ZipFile zip = new ZipFile())
            {
                // Add password to the zip
                zip.Password = zipPassword;
                // Add the healthcheck report pdf to the archive
                zip.AddFile(pdfName,"");
                // Add the matrices folder to the archive;
                zip.AddDirectoryByName("Matrices");
                zip.AddDirectory(OutputDirectory() + "Matrices","Matrices");
                // Save the zip
                zip.Save(zipName);    
            }
        }

        /// <summary>
        /// Uses the CDO namespace to send an email with the report zip attached, to the PM via the google relay, this relay only works at the Cheltenham Office.
        /// </summary>
        /// <example>
        /// <code>
        ///       // Declare and initialise email variables
        ///       string jobNumber = JobDescription.Content.ToString().Substring(0, 5);
        ///       string zipName = OutputDirectory() + jobNumber + " Healthcheck Report.zip";
        ///       string sender = SelectUser(Environment.UserName);
        ///       string senderEmail = Environment.UserName + "@brightsource.co.uk";
        ///       string pmEmailAddress = "";
        ///
        ///       // Display email address input box until entry is received
        ///       do
        ///       {
        ///           pmEmailAddress = VBInteraction.InputBox("Enter the PM's email address");
        ///       } while (String.IsNullOrEmpty(pmEmailAddress));
        ///
        ///        // Create email if the zip file exists
        ///       if (File.Exists(zipName))
        ///       {
        ///             // Create Email Message
        ///             CDO.Message email = new CDO.Message();
        ///             // Message Properties
        ///             email.From = senderEmail;
        ///             email.To = pmEmailAddress + ";" + senderEmail + ";";
        ///             email.Subject = "Healthcheck Report - " + JobDescription.Content;
        ///             email.TextBody = "Hi," + Environment.NewLine + Environment.NewLine + "I have attached the healthcheck report for this job, standard password." + Environment.NewLine + Environment.NewLine + "Location: " + OutputDirectory() + Environment.NewLine + Environment.NewLine + "Can you please confirm receipt of this report." + Environment.NewLine + Environment.NewLine + "Thanks," + Environment.NewLine + Environment.NewLine + sender;
        ///             email.AddAttachment(zipName);
        ///
        ///             try
        ///             {
        ///                 // Define config
        ///                 Configuration gmailConfig = new Configuration();
        ///                 // Configure email relay and send email
        ///                 gmailConfig.Fields["http://schemas.microsoft.com/cdo/configuration/sendusing"].Value = 2;
        ///                 gmailConfig.Fields["http://schemas.microsoft.com/cdo/configuration/smtpserver"].Value = "smtp-relay.gmail.com";
        ///                 gmailConfig.Fields["http://schemas.microsoft.com/cdo/configuration/smtpserverport"].Value = 25;
        ///                 gmailConfig.Fields.Update();
        ///                 email.Configuration = gmailConfig;
        ///
        ///                 // Send the email to the project manager
        ///                 email.Send();
        ///
        ///                 MessageBox.Show("Email sent to '" + pmEmailAddress + "'.");
        ///
        ///                 // Close this application
        ///                 App.Current.Shutdown();
        ///             }
        ///             catch (Exception ex)
        ///             {
        ///                 MessageBox.Show(ex.Message);
        ///             }
        ///         }
        ///         else
        ///         {
        ///             MessageBox.Show("'" + zipName + "' does not exist.");
        ///         }
        /// </code>
        /// </example>
        private void SendReportEmail()
        {
            // Declare and initialise email variables
            string jobNumber = JobDescription.Content.ToString().Substring(0, 5);
            string zipName = OutputDirectory() + jobNumber + " Healthcheck Report.zip";
            string sender = SelectUser(Environment.UserName);
            string senderEmail = Environment.UserName + "@brightsource.co.uk";
            string pmEmailAddress = "";

            // Display email address input box until entry is received
            do
            {
                pmEmailAddress = VBInteraction.InputBox("Enter the PM's email address");
            } while (String.IsNullOrEmpty(pmEmailAddress));

            // Create email if the zip file exists
            if (File.Exists(zipName))
            {
                // Create Email Message
                CDO.Message email = new CDO.Message();
                // Message Properties
                email.From = senderEmail;
                email.To = pmEmailAddress + ";" + senderEmail + ";";
                email.Subject = "Healthcheck Report - " + JobDescription.Content;
                email.TextBody = "Hi," + Environment.NewLine + Environment.NewLine + "I have attached the healthcheck report for this job, standard password." + Environment.NewLine + Environment.NewLine + "Location: " + OutputDirectory() + Environment.NewLine + Environment.NewLine + "Can you please confirm receipt of this report." + Environment.NewLine + Environment.NewLine + "Thanks," + Environment.NewLine + Environment.NewLine + sender;
                email.AddAttachment(zipName);

                try
                {
                    // Define config
                    Configuration gmailConfig = new Configuration();
                    // Configure email relay and send email
                    gmailConfig.Fields["http://schemas.microsoft.com/cdo/configuration/sendusing"].Value = 2;
                    gmailConfig.Fields["http://schemas.microsoft.com/cdo/configuration/smtpserver"].Value = "smtp-relay.gmail.com";
                    gmailConfig.Fields["http://schemas.microsoft.com/cdo/configuration/smtpserverport"].Value = 25;
                    gmailConfig.Fields.Update();
                    email.Configuration = gmailConfig;

                    // Send the email to the project manager
                    email.Send();

                    MessageBox.Show("Email sent to '" + pmEmailAddress + "'.");

                    // Close this application
                    App.Current.Shutdown();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                MessageBox.Show("'" + zipName + "' does not exist.");
            }

        }

        /// <summary>
        /// Converts the current logged in user to a readable first name.
        /// </summary>
        /// <param name="user">string in the format jcooke, tambler etc. Produced by <c>Environment.UserName</c>.</param>
        /// <returns>first name</returns>
        /// <example>
        /// <code>
        ///     // Store user determined by UserName
        ///     user = (user == "jcooke") ? "James" : user;
        ///     user = (user == "dash") ? "Danny" : user;
        ///     user = (user == "tambler") ? "Tim" : user;
        ///     user = (user == "mcurtis") ? "Matt" : user;
        ///     user = (user == "pcorcoran") ? "Phil" : user;
        ///     user = (user == "mvincent") ? "Mike" : user;
        ///     user = (user == "emarlow") ? "Elliott" : user;
        ///     user = (user == "jwyke") ? "Jay" : user;
        ///     user = (user == "sking") ? "Sam" : user;
        ///
        ///     return user;
        /// </code>
        /// </example>
        private string SelectUser(string user)
        {
            // Store user determined by UserName
            user = (user == "jcooke") ? "James" : user;
            user = (user == "dash") ? "Danny" : user;
            user = (user == "tambler") ? "Tim" : user;
            user = (user == "mcurtis") ? "Matt" : user;
            user = (user == "pcorcoran") ? "Phil" : user;
            user = (user == "mvincent") ? "Mike" : user;
            user = (user == "emarlow") ? "Elliott" : user;
            user = (user == "jwyke") ? "Jay" : user;
            user = (user == "sking") ? "Sam" : user;

            return user; 
        }

        // Directory Functions

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public string[] ReadInDirectories()
        {
            string dirFileLocation = @"P:\Data Services\HealthCheckPlus\Development\HealthCheck2\directories.txt";
            string[] directories = new string[3];

            // Add directories to array
            using (StreamReader reader = new StreamReader(dirFileLocation))
            {
                directories[0] = reader.ReadLine().Split(new char[] { '>' })[1]; // Read in client data directory
                directories[1] = reader.ReadLine().Split(new char[] { '>' })[1]; // Read in cygnus jobs directory
                directories[2] = reader.ReadLine().Split(new char[] { '>' })[1]; // Read in wfds directory

            }

            // Remove quotes and extra characters
            for (int d = 0; d < directories.Length; d++)
            {
                directories[d] = directories[d].Trim();
                directories[d] = directories[d].Replace("\"", "");

                // Check \ exists as last character, if not add it
                if (directories[d].Substring(directories[d].Length - 1,1) != "\\")
                {
                    directories[d] += "\\"; 
                }
            }

            return directories;
        }
 
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private string OutputDirectory()
        {
            //return RootOnServer() + @"\Project Management\Data Health Check\Report\Source\"; FOR SERVER DEPLOYMENT

            return @"C:\Users\jcooke\Documents\Visual Studio 2013\Projects\Healthcheck Plus v2.0\Test Output\"; // FOR TESTING

        }

        /// <summary>
        /// Returns the Cygnus job directory by amending the merged XML path.
        /// </summary>
        /// <returns>Cygnus job directory</returns>
        private string CygnusJobDirectory()
        {
            return MergedXMLPath().Replace("_Merged.xml","");
        }

        /// <summary>
        /// Checks the installed programs on the computer for GMC Inspire Designer.
        /// </summary>
        /// <returns>returns true or false if GMC is installed</returns>
        /// <example>
        /// <code>
        ///     string uninstallKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
        ///     bool gmcExists = false;
        ///     string installedProgram = "";
        ///
        ///     using (RegistryKey rk = Registry.LocalMachine.OpenSubKey(uninstallKey))
        ///     {
        ///            // Loop through each reg sub key of installed applications
        ///           foreach (string skName in rk.GetSubKeyNames())
        ///           {
        ///                using (RegistryKey sk = rk.OpenSubKey(skName))
        ///                {
        ///                     try
        ///                     {
        ///                         // Store uppercase name of installed application
        ///                         installedProgram = sk.GetValue("DisplayName").ToString().ToUpper();
        ///                     }
        ///                     catch (Exception ex)
        ///                     { }
        ///
        ///                     // Check gmc exists
        ///                     gmcExists = (installedProgram.Contains("GMC INSPIRE DESIGNER")) ? true : gmcExists;
        ///
        ///                 }
        ///            } 
        ///     }
        ///
        ///    // Return true if installation exists
        ///    return gmcExists;
        /// </code>
        /// </example>
        private bool CheckInstallations()
        {
            string uninstallKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
            bool gmcExists = false;
            string installedProgram = "";

            using (RegistryKey rk = Registry.LocalMachine.OpenSubKey(uninstallKey))
            {
                // Loop through each reg sub key of installed applications
                foreach (string skName in rk.GetSubKeyNames())
                {
                    using (RegistryKey sk = rk.OpenSubKey(skName))
                    {
                        try
                        {
                            // Store uppercase name of installed application
                            installedProgram = sk.GetValue("DisplayName").ToString().ToUpper();
                        }
                        catch (Exception ex)
                        { }

                        // Check gmc exists
                        gmcExists = (installedProgram.Contains("GMC INSPIRE DESIGNER")) ? true : gmcExists;

                    }
                } 
            }

            // Return true if installation exists
            return gmcExists;
        }

        /// <summary>
        /// Stores the path to the _Merged.xml file for the Cygnus job.
        /// </summary>
        /// <returns>the path to the file as a string</returns>
        /// <example><code>
        ///     // Declare variables
        ///     string jobs_directory;
        ///     string job_dir;
        ///     string job_title;
        ///     string[] job_folders;
        ///     string xml_path;
        ///          
        ///     // Initialise variables
        ///     jobs_directory = "Q:\\Cygnus DP\\Jobs\\";
        ///     job_title = JobSelect.SelectedItem.ToString();
        ///     job_dir = jobs_directory + job_title;
        ///     xml_path = "";
        ///     
        ///     // Create xml_path
        ///     job_folders = Directory.GetDirectories(job_dir);
        ///     
        ///     // Find the folder with the _Merged.xml (cygnus folder)
        ///     foreach (var folder in job_folders)
        ///     {
        ///         if (folder.Contains("(" + job_title.Substring(0, 5) + ")"))
        ///         {
        ///             // Store XML path
        ///             xml_path = folder + "\\_Merged.xml";
        ///         }
        ///     }
        ///     
        ///     return xml_path;
        ///</code>
        ///</example>
        private string MergedXMLPath()
        {
            // Declare variables
            string jobs_directory;
            string job_dir;
            string job_title;
            string[] job_folders;
            string xml_path;

            // Initialise variables
            jobs_directory = ReadInDirectories()[1];
            job_title = JobSelect.SelectedItem.ToString();
            job_dir = jobs_directory + job_title;
            xml_path = "";

            // Create xml_path
            job_folders = Directory.GetDirectories(job_dir);

            // Find the folder with the _Merged.xml (cygnus folder)
            foreach (var folder in job_folders)
            {
                if (folder.Contains("(" + job_title.Substring(0, 5) + ")"))
                {
                    // Store XML path
                    xml_path = folder + "\\_Merged.xml";
                }
            }

            return xml_path;
        }

        /// <summary>
        /// Returns the root of the selected job output folder on the server.
        /// </summary>
        /// <returns>returns the path of the folder as string</returns>
        /// <example>
        /// <code>
        ///    // Function to return the root folder on server
        ///    string jobs_directory;
        ///    string job_dir;
        ///    string job_title;
        ///
        ///    // Initialise variables
        ///    jobs_directory = "W:\\Client data\\";
        ///    job_title = JobSelect.SelectedItem.ToString();
        ///    job_dir = jobs_directory + job_title;
        ///
        ///    return job_dir;
        /// </code>
        /// </example>
        private string RootOnServer()
        {
            // Function to return the root folder on server
            string jobs_directory;
            string job_dir;
            string job_title;

            // Initialise variables
            jobs_directory = ReadInDirectories()[0];
            job_title = JobSelect.SelectedItem.ToString();
            job_dir = jobs_directory + job_title;

            return job_dir;

        }

        // Animation Functions

        /// <summary>
        /// Performs initial animations upon running on the window.
        /// </summary>
        private void RunAnimations()
        {
            // Create fade animation
            DoubleAnimation fade = new DoubleAnimation();
            fade.From = 1.0;
            fade.To = 0.0;
            fade.Duration = new Duration(TimeSpan.FromSeconds(0.5));
            // Apply storyboard
            fading = new Storyboard();
            fading.Children.Add(fade);

            // Set storyboard target - FADE DROP DOWN MENU
            Storyboard.SetTargetName(fade, JobSelect.Name);
            // Target opacity property
            Storyboard.SetTargetProperty(fade, new PropertyPath(ComboBox.OpacityProperty));
            // Begin animation
            fading.Begin(this);

            // Set storyboard target - FADE JOB CONFIRM
            Storyboard.SetTargetName(fade, JobConfirm.Name);
            // Target opacity property
            Storyboard.SetTargetProperty(fade, new PropertyPath(Button.OpacityProperty));
            // Begin animation
            fading.Begin(this);

            // Set storyboard target - FADE RECENT JOB LABEL
            Storyboard.SetTargetName(fade, RecentlyReportedLabel.Name);
            // Target opacity property
            Storyboard.SetTargetProperty(fade, new PropertyPath(Label.OpacityProperty));
            // Begin animation
            fading.Begin(this);

            // Set storyboard target - FADE RECENT JOB LIST
            Storyboard.SetTargetName(fade, RecentJobList.Name);
            // Target opacity property
            Storyboard.SetTargetProperty(fade, new PropertyPath(ListBox.OpacityProperty));
            // Begin animation
            fading.Begin(this);

            // Make inactive
            JobConfirm.IsEnabled = false;
            JobSelect.IsEnabled = false;
            RecentJobList.IsEnabled = false;

            // Make invisible
            JobConfirm.Visibility = Visibility.Hidden;
            JobSelect.Visibility = Visibility.Hidden;
            RecentlyReportedLabel.Visibility = Visibility.Hidden;
            RecentJobList.Visibility = Visibility.Hidden;

            // Populate JobDescription Label
            JobDescription.Content = JobSelect.SelectedItem.ToString();

            // Move JobDescription Label
            JobDescription.Visibility = Visibility.Visible;

            // Create movement animation
            ThicknessAnimation move = new ThicknessAnimation();
            move.From = new Thickness(10, 35, 0, 0);
            move.To = new Thickness(5, 0, 0, 0);
            move.Duration = new Duration(TimeSpan.FromSeconds(0.5));
            // Apply storyboard
            moving = new Storyboard();
            moving.Children.Add(move);

            // Set storyboard target - MOVE JOB DESC
            Storyboard.SetTargetName(move, JobDescription.Name);
            // Target opacity property
            Storyboard.SetTargetProperty(move, new PropertyPath(Label.MarginProperty));
            // Begin animation
            moving.Begin(this);

            // Create appear animation
            DoubleAnimation appear = new DoubleAnimation();
            appear.From = 0.0;
            appear.To = 1.0;
            appear.Duration = new Duration(TimeSpan.FromSeconds(0.5));
            // Apply storyboard
            appearing = new Storyboard();
            appearing.Children.Add(appear);

            // Set storyboard target - APPEAR FILE LIST BOX
            Storyboard.SetTargetName(appear, MailingFileList.Name);
            // Target opacity property
            Storyboard.SetTargetProperty(appear, new PropertyPath(ListBox.OpacityProperty));
            // Begin animation
            appearing.Begin(this);

            // Set storyboard target - APPEAR FILES CONFIRM
            Storyboard.SetTargetName(appear, FilesConfirm.Name);
            // Target opacity property
            Storyboard.SetTargetProperty(appear, new PropertyPath(Button.OpacityProperty));
            // Begin animation
            appearing.Begin(this);
        }

        /// <summary>
        /// Performs animations after the user clicks confirm on the window.
        /// </summary>
        private void ConfirmAnimations()
        {
            // Create fade animation
            DoubleAnimation fade = new DoubleAnimation();
            fade.From = 1.0;
            fade.To = 0.0;
            fade.Duration = new Duration(TimeSpan.FromSeconds(0.5));
            // Apply storyboard
            fading = new Storyboard();
            fading.Children.Add(fade);

            // Set storyboard target - FADE CONFIRM FILES
            Storyboard.SetTargetName(fade, FilesConfirm.Name);
            // Target opacity property
            Storyboard.SetTargetProperty(fade, new PropertyPath(Button.OpacityProperty));
            // Begin animation
            fading.Begin(this);

            // Set storyboard target - FADE FILE LIST
            Storyboard.SetTargetName(fade, MailingFileList.Name);
            // Target opacity property
            Storyboard.SetTargetProperty(fade, new PropertyPath(ListBox.OpacityProperty));
            // Begin animation
            fading.Begin(this);

        }

        /// <summary>
        /// Performs beginning dedupe animations on the window.
        /// </summary>
        private void DedupeAnimations()
        {
            // Disable previous button
            FilesConfirm.IsEnabled = false;
            FilesConfirm.Visibility = Visibility.Hidden;
            
            // Dedupe Button visibility
            RunDedupe.Visibility = Visibility.Visible;
            RunDedupe.Opacity = 0.0;
            
            // Create appear animation
            DoubleAnimation appear = new DoubleAnimation();
            appear.From = 0.0;
            appear.To = 1.0;
            appear.Duration = new Duration(TimeSpan.FromSeconds(0.5));
            // Apply storyboard
            appearing = new Storyboard();
            appearing.Children.Add(appear);

            // Set storyboard target - APPEAR FILE LIST BOX
            Storyboard.SetTargetName(appear, MailingFileList.Name);
            // Target opacity property
            Storyboard.SetTargetProperty(appear, new PropertyPath(ListBox.OpacityProperty));
            // Begin animation
            appearing.Begin(this);

            // Set storyboard target - APPEAR DEDUPE
            Storyboard.SetTargetName(appear, RunDedupe.Name);
            // Target opacity property
            Storyboard.SetTargetProperty(appear, new PropertyPath(Button.OpacityProperty));
            // Begin animation
            appearing.Begin(this);
        }

        /// <summary>
        /// Performs ending dedupe animatios on the window.
        /// </summary>
        private void EndDedupeAnimations()
        {
            // Disable previous button
            RunDedupe.IsEnabled = false;
            RunDedupe.Visibility = Visibility.Hidden;

            // Disable previous button
            FilesConfirm.IsEnabled = false;
            FilesConfirm.Visibility = Visibility.Hidden;

            // Create fade animation
            DoubleAnimation fade = new DoubleAnimation();
            fade.From = 1.0;
            fade.To = 0.0;
            fade.Duration = new Duration(TimeSpan.FromSeconds(0.5));
            // Apply storyboard
            fading = new Storyboard();
            fading.Children.Add(fade);

            // Set storyboard target - FADE CONFIRM FILES
            Storyboard.SetTargetName(fade, RunDedupe.Name);
            // Target opacity property
            Storyboard.SetTargetProperty(fade, new PropertyPath(Button.OpacityProperty));
            // Begin animation
            fading.Begin(this);

            // Clear the mailing file list
            MailingFileList.Items.Clear();

            // Set storyboard target - FADE FILE LIST
            Storyboard.SetTargetName(fade, MailingFileList.Name);
            // Target opacity property
            Storyboard.SetTargetProperty(fade, new PropertyPath(ListBox.OpacityProperty));
            // Begin animation
            fading.Begin(this);
        }

        /// <summary>
        /// Performs final animations on the window.
        /// </summary>
        private void SummaryAnimations()
        {
            // Disable previous button
            FilesConfirm.IsEnabled = false;
            FilesConfirm.Visibility = Visibility.Hidden;

            // Dedupe Button visibility
            RunDedupe.IsEnabled = false;
            RunDedupe.Visibility = Visibility.Hidden;

            // TextSummary visibility
            TextSummary.Visibility = Visibility.Visible; 

            // Create appear animation
            DoubleAnimation appear = new DoubleAnimation();
            appear.From = 0.0;
            appear.To = 1.0;
            appear.Duration = new Duration(TimeSpan.FromSeconds(0.5));
            // Apply storyboard
            appearing = new Storyboard();
            appearing.Children.Add(appear);

            // Set storyboard target - APPEAR SUMMARY
            Storyboard.SetTargetName(appear, TextSummary.Name);
            // Target opacity property
            Storyboard.SetTargetProperty(appear, new PropertyPath(RichTextBox.OpacityProperty));
            // Begin animation
            appearing.Begin(this);

            // Set storyboard target - APPEAR CREATE REPORT
            Storyboard.SetTargetName(appear, CreateReport.Name);
            // Target opacity property
            Storyboard.SetTargetProperty(appear, new PropertyPath(Button.OpacityProperty));
            // Begin animation
            appearing.Begin(this);

            
        }
 
    }

    /// <summary>
    /// When instantiated provides headings for the majority of the output files.
    /// </summary>
    /// <example>
    /// <code>
    ///  // Retrieve headings
    ///  FileHeadings headings = new FileHeadings();
    ///
    ///  // Write line to the csv file
    ///  csv_output.WriteLine(headings.GetDownloadStatsHeadings()); // Writing field names
    /// </code>
    /// </example>
    public class FileHeadings
    {
        /// <summary>
        /// field heading string
        /// </summary>
        string download_stats_headings;
        /// <summary>
        /// field heading string
        /// </summary>
        string file_structure_headings;
        /// <summary>
        /// field heading string
        /// </summary>
        string dedupe_information_headings;
        /// <summary>
        /// field heading string
        /// </summary>
        string dedupe_stats_headings;
        /// <summary>
        /// field heading string
        /// </summary>
        string dedupe_total_headings;
        /// <summary>
        /// field heading string
        /// </summary>
        string dedupe_example_headings;
        /// <summary>
        /// field heading string
        /// </summary>
        string industry_suppression_headings;
        /// <summary>
        /// field heading string
        /// </summary>
        string suppression_cost_headings;

        /// <summary>
        /// The constructor for the class that populates the different file heading properties
        /// with their comma separated values, returned as a string to be written out as the first line in a CSV.
        /// </summary>
        /// <example>
        /// <code>
        /// industry_suppression_headings = "File,Input,NCOA_REDIR,GAS_REACT,NCOA_DECEASED,TBR,NDR,MORTA,MPS_DECEASED,NCOA,GAS,MPS_COLD,TOTAL DROPS, OUTPUT, TOTAL HITS,MATCH_LEVEL";
        /// </code>
        /// </example>
        public FileHeadings()
        {
            download_stats_headings = "FILE_NAME" + "," + "INPUT_RECORDS" + "," + "SPLIT_RECORDS" + "," + "OUTPUT_RECORDS" + "," + "POSSIBLE_OBSCENITIES" + "," + "HAS_DIACRITICS" + "," + "PREPARSED" + "," + "JUMBLED" + "," + "SUSPECT_DATA" + "," + "NOTABLES" + "," + "NAMES_MALE" + "," + "NAMES_FEMALE" + "," + "NAMES_NEUTRAL" + "," + "NAMES_JOINT" + "," + "NAMES_BLANK" + "," + "NAMES_DEFAULT" + "," + "NAMES_SINGLE" + "," + "NAMES_OTHER" + "," + "SAL_MALE" + "," + "SAL_FEMALE" + "," + "SAL_NEUTRAL" + "," + "SAL_JOINT" + "," + "SAL_BLANK" + "," + "SAL_DEFAULT" + "," + "SAL_UNRECOG" + "," + "POSSIBLE_DECEASED" + "," + "HAS_AFFIX" + "," + "ADD_UKIDENTIFIED" + "," + "ADD_UKASSUMED" + "," + "ADD_OVERSEAS" + "," + "ADD_BLANK" + "," + "ADD_BUSINESS" + "," + "ADD_RESIDENTIAL" + "," + "ADD_SUSPECT" + "," + "ADD_BFPO" + "," + "ZIP_FULL" + "," + "ZIP_OUTBOUND" + "," + "ZIP_NONE" + "," + "ZIP_INVALID" + "," + "DIRECT" + "," + "RESIDUE" + "," + "DPS_PRESENT" + "," + "DPS_DEFAULT" + ",";
            file_structure_headings = "FILE_NAME" + "," + "NAME" + "," + "START" + "," + "LENGTH" + "," + "DATA" + "," + "CONSTANT" + "," + "BLANKS" + "," + "AVERAGE" + "," + "MIN" + "," + "MAX" + "," + "FIELD_NUMBER";
            dedupe_information_headings = "DINFO_NAME" + "," + "DINFO_CATEGORY" + "," + "DINFO_BASIS" + "," + "DINFO_CRITERION" + "," + "DINFO_ACTION" + "," + "DINFO_LVLACTION" + "," + "DINFO_TARGET" + "," + "DINFO_MISMATCH" + "," + "FORENAME_DEDUPE" + "," + "FORENAME_SUPPRESS";
            dedupe_stats_headings = "DEDUPE_NAME" + "," + "FILE_NAME" + "," + "LIST_CODE" + "," + "PRIORITY" + "," + "INPUT_RECORDS" + "," + "SUPPRESSED" + "," + "INTER_DUPES" + "," + "INTRA_DUPES" + "," + "TOTAL_DROPS" + "," + "UNIQUES" + "," + "SELECTED" + "," + "OUTPUT_RECORDS";
            dedupe_total_headings = "FILE_NAME" + "," + "LIST_CODE" + "," + "PRIORITY" + "," + "INPUT_RECORDS" + "," + "SUPPRESSED" + "," + "INTER_DUPES" + "," + "INTRA_DUPES" + "," + "TOTAL_DROPS" + "," + "UNIQUES" + "," + "SELECTED" + "," + "OUTPUT_RECORDS";
            dedupe_example_headings = "DEDUPE_NAME" + "," + "DUPE_GROUP" + "," + "FILENAME" + "," + "PRIORITY" + "," + "ACTION" + "," + "NAME" + "," + "ADDRESS";
            industry_suppression_headings = "File,Input,NCOA_REDIR,GAS_REACT,NCOA_DECEASED,TBR,NDR,MORTA,MPS_DECEASED,NCOA,GAS,MPS_COLD,TOTAL DROPS, OUTPUT, TOTAL HITS,MATCH_LEVEL";
            suppression_cost_headings = "Cost Type,,NCOA_REDIR,GAS_REACT,NCOA_DECEASED,TBR,NDR,MORTA,MPS_DECEASED,NCOA,GAS,MPS_COLD,Total_Cost";
        }

        /// <summary>
        /// Returns the headings required for the download statistics CSV.
        /// </summary>
        /// <returns>download stats field headings</returns>
        public string GetDownloadStatsHeadings()
        {
            return download_stats_headings;
        }

        /// <summary>
        /// Returns the headings required for the file structures CSV.
        /// </summary>
        /// <returns>file structures field headings</returns>
        public string GetStructureHeadings()
        {
            return file_structure_headings;
        }

        /// <summary>
        /// Returns the headings required for the dedupe info CSV.
        /// </summary>
        /// <returns>dedupe info field headings</returns>
        public string GetDedupeInformationHeadings()
        {
            return dedupe_information_headings;
        }

        /// <summary>
        /// Returns the headings required for the dedupe examples CSV.
        /// </summary>
        /// <returns>dedupe examples field headings</returns>
        public string GetDedupeExampleHeadings()
        {
            return dedupe_example_headings;
        }

        /// <summary>
        /// Returns the headings required for the dedupe stats CSV.
        /// </summary>
        /// <returns>dedupe stats field headings</returns>
        public string GetDedupeStatsHeadings()
        {
            return dedupe_stats_headings;
        }

        /// <summary>
        /// Returns the headings required for the dedupe totals CSV.
        /// </summary>
        /// <returns>dedupe totals field headings</returns>
        public string GetDedupeTotalHeadings()
        {
            return dedupe_total_headings;
        }

        /// <summary>
        /// Returns the headings required for the industry suppressions CSV.
        /// </summary>
        /// <returns>industry suppressions field headings</returns>
        public string GetSuppressionHeadings()
        {
            return industry_suppression_headings;
        }

        /// <summary>
        /// Returns the headings required for the industry suppression costs CSV.
        /// </summary>
        /// <returns>industry suppression costs field headings</returns>
        public string GetCostHeadings()
        {
            return suppression_cost_headings;
        }
    } 
}
