using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Configuration;
using System.Data.SQLite;
using System.IO;
using System.Data;
using PicImp;
using System.Security;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Windows.Media;

namespace PhotoSorter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        SQLiteConnection m_dbConnection;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Load(object sender, RoutedEventArgs e)
        {
            try
            {
                //Hide-show Shutterfly
                bool useShutterfly = false;
                Boolean.TryParse(ConfigurationManager.AppSettings["useShutterfly"].ToString(), out useShutterfly);
                ApplicationState.SetValue("passwdChanged", false);
                ApplicationState.SetValue("sharedSecretChanged", false);
                txtInstallDir.Text = System.AppDomain.CurrentDomain.BaseDirectory.TrimEnd('\\');

                if (useShutterfly)
                {
                    ((TabItem)tcMain.Items[3]).Visibility = Visibility.Visible;
                }
                else
                {
                    ((TabItem)tcMain.Items[3]).Visibility = Visibility.Hidden;
                }

                if (!File.Exists("PhotoSorter.db"))
                {
                    SQLiteConnection.CreateFile("PhotoSorter.db");
                    OpenDB();

                    //Initialize config table
                    string defaultImportDir = @"%HOMEDRIVE%%HOMEPATH%\Pictures\Import";
                    string defaultPictureDir = @"%HOMEDRIVE%%HOMEPATH%\Pictures";
                    string defaultVideoDir = @"%HOMEDRIVE%%HOMEPATH%\Videos";
                    string newsql = "CREATE TABLE tbl_configs (user_pk int, importDirectory varchar(5000), pictureDirectory varchar(5000), videoDirectory varchar(5000), prefix varchar(20), suffix varchar(20), sortTime nullable datetime, includeCameraMake bit, sfUserID varchar(100), sfPasswd varchar(100), sfAppID varchar(100), sfSS varchar(100), sfAuthID varchar(100))";
                    RunSQLCMD(newsql);
                    newsql = String.Format("INSERT INTO tbl_configs (user_pk, importDirectory, pictureDirectory, videoDirectory) values (1, '{0}','{1}','{2}')", defaultImportDir, defaultPictureDir, defaultVideoDir);
                    RunSQLCMD(newsql);

                    //Initialize video table
                    newsql = "CREATE TABLE tbl_videos (extension varchar(20))";
                    RunSQLCMD(newsql);

                    var defaultVideoExts = new string[] { ".mp4", ".m4v", ".avi", ".mov", ".mpg", ".wmv", ".webm" };
                    foreach (string videoExt in defaultVideoExts)
                    {
                        newsql = String.Format("INSERT INTO tbl_videos (extension) VALUES ('{0}')", videoExt);
                        RunSQLCMD(newsql);
                    }
                }
                else
                {
                    OpenDB();
                }

                List<string> exts = GetVideoExtensions();
                lstExts.ItemsSource = exts;

                string sql = "SELECT * FROM tbl_configs";
                SQLiteCommand command = new SQLiteCommand(sql, m_dbConnection);
                SQLiteDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    if (reader["includeCameraMake"] != DBNull.Value)
                    {
                        chkCameraName.IsChecked = Convert.ToBoolean(reader["includeCameraMake"]);
                    }
                    txtBrowseDumpDir.Text = Environment.ExpandEnvironmentVariables(reader["importDirectory"].ToString());
                    txtBrowseDestDir.Text = Environment.ExpandEnvironmentVariables(reader["pictureDirectory"].ToString());
                    txtBrowseVidDestDir.Text = Environment.ExpandEnvironmentVariables(reader["videoDirectory"].ToString());
                    txtPrefix.Text = reader["prefix"].ToString();
                    txtSuffix.Text = reader["suffix"].ToString();
                    if (reader["includeCameraMake"] != DBNull.Value)
                    {
                        chkCameraName.IsChecked = Convert.ToBoolean(reader["includeCameraMake"]);
                    }
                    if (reader["sortTime"] != DBNull.Value && reader["sortTime"] != "")
                    {
                        tpSortTime.Value = Convert.ToDateTime(reader["sortTime"]);
                    }
                    //Shutterfly
                    txtSFUserID.Text = reader["sfUserID"].ToString();
                    if (reader["sfPasswd"] != DBNull.Value)
                    {
                        txtSFPasswd.Password = "placeholder";
                    }
                    txtSFAppID.Text = reader["sfAppID"].ToString();
                    if (reader["sfSS"] != DBNull.Value)
                    {
                        txtSFSharedSecret.Password = "placeholder";
                    }
                }
                reader.Close();
            }
            catch (Exception ex){ }
            finally
            {
                m_dbConnection.Close();
            }
        }

        

        private int RunSQLCMD(string sql)
        {
            SQLiteCommand command = new SQLiteCommand(sql, m_dbConnection);
            int recUpdated = command.ExecuteNonQuery();
            return recUpdated;
        }

        private List<string> GetVideoExtensions()
        {
            List<string> exts = new List<string>();
            string query = "SELECT * FROM tbl_videos ORDER BY extension";
            SQLiteCommand cmd = new SQLiteCommand(query, m_dbConnection);
            SQLiteDataReader rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                exts.Add(rdr["extension"].ToString());
            }
            rdr.Close();
            return exts;
        }

        private void OpenDB()
        {
            m_dbConnection = new SQLiteConnection("Data Source=PhotoSorter.db;Version=3;");
            m_dbConnection.Open();
        }

        private void txtPrefix_TextChanged(object sender, TextChangedEventArgs e)
        {
            BuildExampleText();
        }

        private void btnBrowseDumpDir_Click(object sender, RoutedEventArgs e)
        {
            string orgDumpDir = txtBrowseDumpDir.Text;
            var folderDialog = new System.Windows.Forms.FolderBrowserDialog();
            var result = folderDialog.ShowDialog();
            switch (result)
            {
                case System.Windows.Forms.DialogResult.OK:
                    var path = folderDialog.SelectedPath;
                    txtBrowseDumpDir.Text = path;
                    txtBrowseDumpDir.ToolTip = path;
                    break;
                case System.Windows.Forms.DialogResult.Cancel:
                default:
                    txtBrowseDumpDir.Text = orgDumpDir;
                    txtBrowseDumpDir.ToolTip = orgDumpDir;
                    break;
            }
        }

        private void btnBrowseDestDir_Click(object sender, RoutedEventArgs e)
        {
            string orgDestDir = txtBrowseDestDir.Text;
            var folderDialog = new System.Windows.Forms.FolderBrowserDialog();
            var result = folderDialog.ShowDialog();
            switch (result)
            {
                case System.Windows.Forms.DialogResult.OK:
                    var path = folderDialog.SelectedPath;
                    txtBrowseDestDir.Text = path;
                    txtBrowseDestDir.ToolTip = path;
                    break;
                case System.Windows.Forms.DialogResult.Cancel:
                default:
                    txtBrowseDestDir.Text = orgDestDir;
                    txtBrowseDestDir.ToolTip = orgDestDir;
                    break;
            }
        }

        private void btnBrowseVidDestDir_Click(object sender, RoutedEventArgs e)
        {
            string orgDestDir = txtBrowseVidDestDir.Text;
            var folderDialog = new System.Windows.Forms.FolderBrowserDialog();
            var result = folderDialog.ShowDialog();
            switch (result)
            {
                case System.Windows.Forms.DialogResult.OK:
                    var path = folderDialog.SelectedPath;
                    txtBrowseVidDestDir.Text = path;
                    txtBrowseVidDestDir.ToolTip = path;
                    break;
                case System.Windows.Forms.DialogResult.Cancel:
                default:
                    txtBrowseVidDestDir.Text = orgDestDir;
                    txtBrowseVidDestDir.ToolTip = orgDestDir;
                    break;
            }
        }

        private void txtBrowseDestDir_TextChanged(object sender, TextChangedEventArgs e)
        {
            BuildExampleText();
        }

        private void txtBrowseVidDestDir_TextChanged(object sender, TextChangedEventArgs e)
        {
            BuildExampleText();
        }

        private void BuildExampleText()
        {
            string cameraName = "";
            if (chkCameraName.IsChecked.Value)
            {
                cameraName = "_NikonL340";
            }
            txtPhotoExample.Text = GetPicDateDirStructurePath(DateTime.Now) + @"\" + txtPrefix.Text + DateTime.Now.ToString("yyyyMMdd_HHmmss") + cameraName + txtSuffix.Text + ".jpg";
            txtVideoExample.Text = txtBrowseVidDestDir.Text + @"\" + DateTime.Now.Year.ToString() + @"\" + txtPrefix.Text + DateTime.Now.ToString("yyyyMMdd_HHmmss") + cameraName + txtSuffix.Text + ".mp4";
        }


        private void txtSuffix_TextChanged(object sender, TextChangedEventArgs e)
        {
            BuildExampleText();
        }

        private string GetPicDateDirStructurePath(DateTime dt)
        {
            string year = dt.Year.ToString();
            string monthNum = dt.Month.ToString().PadLeft(2, '0');
            string monAbbrv = dt.ToString("MMM");

            return txtBrowseDestDir.Text + @"\" + year + @"\" + monthNum + "_" + monAbbrv;
        }

        private void chkCameraName_Click(object sender, RoutedEventArgs e)
        {
            BuildExampleText();
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            string msg = SaveMainInfo();
            System.Windows.MessageBox.Show(msg, "Alert", MessageBoxButton.OK, MessageBoxImage.Exclamation);
        }

        private string SaveMainInfo()
        {
            string msg = "Successfully updated.";
            try
            {
                OpenDB();
                string sql = String.Format("UPDATE tbl_configs SET importDirectory = '{0}', pictureDirectory = '{1}', videoDirectory='{2}', prefix = '{3}', suffix = '{4}', sortTime = '{5}', includeCameraMake = {6} WHERE user_pk = 1", txtBrowseDumpDir.Text, txtBrowseDestDir.Text, txtBrowseVidDestDir.Text, txtPrefix.Text, txtSuffix.Text, tpSortTime.Value, chkCameraName.IsChecked.Value ? 1 : 0);
                int recUpdated = RunSQLCMD(sql);
                
                if (recUpdated != 1)
                {
                    msg = "Error saving settings.";
                }
            }
            catch { }
            finally
            {
                m_dbConnection.Close();
            }

            return msg;
        }

        private void RemoveItem(object sender, RoutedEventArgs e)
        {
            try
            {
                System.Windows.Controls.MenuItem mi = ((System.Windows.Controls.MenuItem)sender);
                System.Windows.Controls.ContextMenu cm = mi.CommandParameter as System.Windows.Controls.ContextMenu;
                System.Windows.Controls.ListBox lb = cm.PlacementTarget as System.Windows.Controls.ListBox;
                int selIndex = lb.SelectedIndex;
                string name = lb.Items[selIndex].ToString();

                OpenDB();
                string sql = String.Format("DELETE FROM tbl_videos WHERE extension = '{0}'", name);
                RunSQLCMD(sql);

                List<string> exts = GetVideoExtensions();
                lstExts.ItemsSource = exts;
            }
            catch { }
            finally
            {
                m_dbConnection.Close();
            }

        }

        private void btnAddExt_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string name = '.' + txtNewExt.Text.TrimStart('.');

                if (lstExts.Items.Cast<string>().Where(l => l.ToUpper() == name.ToUpper()).Any())
                {
                    System.Windows.MessageBox.Show("This extension is already added.", "Alert", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    return;
                }

                OpenDB();
                string sql = String.Format("INSERT INTO tbl_videos(extension) values('{0}')", name);
                RunSQLCMD(sql);

                List<string> exts = GetVideoExtensions();
                lstExts.ItemsSource = exts;
            }
            catch (Exception ex) { }
            finally
            {
                m_dbConnection.Close();
            }
        }

        private void btnSFSave_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                bool passwdChanged = ApplicationState.GetValue<bool>("passwdChanged");
                bool sharedSecretChanged = ApplicationState.GetValue<bool>("sharedSecretChanged");

                string sql = String.Format("UPDATE tbl_configs SET sfUserID = '{0}', sfAppID = '{1}' WHERE user_pk = 1", txtSFUserID.Text, txtSFAppID.Text);
                if (passwdChanged && !sharedSecretChanged)
                {
                    SecureString securedString = txtSFPasswd.SecurePassword;
                    string encryptedString = Encrypting.EncryptString(securedString);
                    sql = String.Format("UPDATE tbl_configs SET sfUserID = '{0}', sfAppID = '{1}', sfPasswd = '{2}' WHERE user_pk = 1", txtSFUserID.Text, txtSFAppID.Text, encryptedString);
                }
                if (sharedSecretChanged && !passwdChanged)
                {
                    SecureString securedString = txtSFSharedSecret.SecurePassword;
                    string encryptedString = Encrypting.EncryptString(securedString);
                    sql = String.Format("UPDATE tbl_configs SET sfUserID = '{0}', sfAppID = '{1}', sfSS = '{2}' WHERE user_pk = 1", txtSFUserID.Text, txtSFAppID.Text, encryptedString);
                }
                if (sharedSecretChanged && passwdChanged)
                {
                    SecureString securedPsswd = txtSFPasswd.SecurePassword;
                    string encryptedPsswd = Encrypting.EncryptString(securedPsswd);
                    SecureString securedSS = txtSFSharedSecret.SecurePassword;
                    string encryptedSS = Encrypting.EncryptString(securedSS);

                    sql = String.Format("UPDATE tbl_configs SET sfUserID = '{0}', sfAppID = '{1}', sfPasswd = '{2}', sfSS = '{3}' WHERE user_pk = 1", txtSFUserID.Text, txtSFAppID.Text, encryptedPsswd, encryptedSS);
                }

                OpenDB();
                RunSQLCMD(sql);

                System.Windows.MessageBox.Show("Information saved.", "Alert", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
            catch { }
            finally
            {
                ApplicationState.SetValue("passwdChanged", false);
                ApplicationState.SetValue("sharedSecretChanged", false);
                m_dbConnection.Close();
            }

        }

        private void txtSFPasswd_PasswordChanged(object sender, RoutedEventArgs e)
        {
            ApplicationState.SetValue("passwdChanged", true);
        }

        private void txtSFSharedSecret_PasswordChanged(object sender, RoutedEventArgs e)
        {
            ApplicationState.SetValue("sharedSecretChanged", true);
        }

        public static class ApplicationState
        {
            private static Dictionary<string, object> _values =
                       new Dictionary<string, object>();

            public static void SetValue(string key, object value)
            {
                if (_values.ContainsKey(key))
                {
                    _values[key] = value;
                }
                else
                {
                    _values.Add(key, value);
                }
            }

            public static T GetValue<T>(string key)
            {
                return (T)_values[key];
            }
        }

        private void btnRun_Click(object sender, RoutedEventArgs e)
        {
            biMain.IsBusy = true;

            //do a quick save in case something changed
            SaveMainInfo();

            Task.Factory.StartNew(() =>
            {
                Importer imp = new Importer();
                imp.RunImport();
                System.Windows.MessageBox.Show("Import complete", "Alert", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
            ).ContinueWith((task) =>
            {
                biMain.IsBusy = false;
            }, TaskScheduler.FromCurrentSynchronizationContext()
            );
        }

        public void OpenHelpDoc_Click(object sender, RoutedEventArgs e)
        {
            Process wordProcess = new Process();
            wordProcess.StartInfo.FileName = @"docs\Photo_Sorter_Help_Document.rtf";
            wordProcess.StartInfo.UseShellExecute = true;
            wordProcess.Start();  
        }

        public void OpenInstallDir_Click(object sender, RoutedEventArgs e)
        {
            Process.Start(txtInstallDir.Text);
        }

        public void OnCopyHyperLink_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.MenuItem mi = ((System.Windows.Controls.MenuItem)sender);
            System.Windows.Controls.ContextMenu cm = mi.CommandParameter as System.Windows.Controls.ContextMenu;
            System.Windows.Controls.TextBlock txt = cm.PlacementTarget as System.Windows.Controls.TextBlock;
            System.Windows.Controls.TextBlock lnk = txt.FindName("txtInstallDir") as System.Windows.Controls.TextBlock;
            Clipboard.SetText(lnk.Text);
        }
    }
}
