using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.ServiceProcess;
using System.Text;
using Microsoft.Win32;
using System.Threading;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Net.Mail;
using System.Configuration;
using System.Data.SqlClient;
using System.Management;
using System.Net.Sockets;

namespace viraBackup
{
    public partial class viraBackup : ServiceBase
    {
        public bool OnFirstTime;
        public Timer serviceTimer;
        public bool OnWorking;
        public Int64 workingID = 1;


        // Distributor: VIRAWARE™ VIRAW.IR 2025
        // Developer & Contributor: Babak Arjomandi @ www.babakarjomandi.com


        // Install        
        // C:\Windows\Microsoft.NET\Framework64\v4.0.30319\installutil.exe c:\viraBackup\viraBackup.exe

        // Uninstall
        // C:\Windows\Microsoft.NET\Framework64\v4.0.30319\installutil.exe /u c:\viraBackup\viraBackup.exe


        public viraBackup()
        {
            InitializeComponent();
        }


        public static string GetPageAsString(Uri address)
        {
            try
            {
                System.Net.WebClient wc = new System.Net.WebClient();            
                wc.Encoding = Encoding.UTF8;
                return wc.DownloadString(address.AbsoluteUri);
            }
            catch
            {
                return null;
            }
        }


        public static void SendSMS(string toCells, string message, string template)
        {
            if (!string.IsNullOrEmpty(ConfigurationManager.AppSettings["smsProviderUrl"]))
            {
                string _smsP = ConfigurationManager.AppSettings["smsProviderUrl"].ToString().Replace("|", "&"); // <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

                string smsRequest = _smsP.Replace("(cell)", toCells).Replace("(code)", message).Replace("(template)", template);
                string sms = GetPageAsString(new Uri(smsRequest));
            }
        }


        protected override void OnStart(string[] args)
        {
            TimerCallback timerDelegate = new TimerCallback(startForFirstTime);
            Int32 intLoop = 1000;
            serviceTimer = new Timer(timerDelegate, null, 0, intLoop);
            if (!string.IsNullOrEmpty(ConfigurationManager.AppSettings["smsProviderUrl"]))
            {
                SendSMS("+98" + ConfigurationManager.AppSettings["cell"], ConfigurationManager.AppSettings["serverName"] + "-" + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString(), "BackupStart");
            }
            else
            {
                SendEmail("[" + ConfigurationManager.AppSettings["serverName"] + "] Vira Backup service started!", "Vira Backup service started! process ID:" + workingID.ToString() + "<br>" +
                           "<hr/>" + "viraBackup " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString() + "<br>" +
                           "Powered by: https://www.viraw.ir");
            }
            OnWorking = false;
            OnFirstTime = true;
        }


        protected override void OnStop()
        {
            if (!string.IsNullOrEmpty(ConfigurationManager.AppSettings["smsProviderUrl"]))
            {
                SendSMS("+98" + ConfigurationManager.AppSettings["cell"], ConfigurationManager.AppSettings["serverName"], "BackupStop");
            }
            else
            {
                SendEmail("[" + ConfigurationManager.AppSettings["serverName"] + "] Vira Backup service stoped!", "Vira Backup service stoped! Last process ID:" + workingID.ToString());
            }
        }


        public void startForFirstTime(object state)
        {
            if ((DateTime.Now.Hour == Convert.ToInt32(ConfigurationManager.AppSettings["firstTimeSatrtBackupTime"])) && (DateTime.Now.Minute == 0) && (OnFirstTime == true))
            {
                OnFirstTime = false;
                try
                {
                    TimerCallback timerDelegate = new TimerCallback(DoWork);
                    Int32 intLoop = Convert.ToInt32(ConfigurationManager.AppSettings["intervalBackup"]) * 3600000;
                    serviceTimer = new Timer(timerDelegate, null, 0, intLoop);
                    if (!string.IsNullOrEmpty(ConfigurationManager.AppSettings["smsProviderUrl"]))
                    {
                        SendSMS("+98" + ConfigurationManager.AppSettings["cell"], ConfigurationManager.AppSettings["serverName"], "BackupFirst");
                    }
                    else
                    {
                        SendEmail("[" + ConfigurationManager.AppSettings["serverName"] + "] Vira Backup process started!", "Vira Backup process started at " + DateTime.Now.ToString());
                    }
                }
                catch (Exception e)
                {
                    SubmitLog("Error on start main timer :: " + e.Message);
                }
            }
        }


        public void DoWork(object state)
        {
            if (!OnWorking)
            {
                int errorCount = 0;
                OnWorking = true;
                workingID += 1;
                DateTime dtStart = DateTime.Now;

                // Create Backup From files and SQL ------------------------------------------
                string folderDate = DateTime.Now.ToString("yyyy-MM-dd-HH");

                if (!Directory.Exists(ConfigurationManager.AppSettings["rootPathBackup"]))
                {
                    Directory.CreateDirectory(ConfigurationManager.AppSettings["rootPathBackup"]);
                }
                if (!Directory.Exists(ConfigurationManager.AppSettings["rootPathBackup"] + folderDate))
                {
                    Directory.CreateDirectory(ConfigurationManager.AppSettings["rootPathBackup"] + folderDate);
                }
                if (!Directory.Exists(ConfigurationManager.AppSettings["rootPathBackup"] + folderDate + "\\sql"))
                {
                    Directory.CreateDirectory(ConfigurationManager.AppSettings["rootPathBackup"] + folderDate + "\\sql");
                }
                if (!Directory.Exists(ConfigurationManager.AppSettings["rootPathBackup"] + folderDate + "\\wwwroot"))
                {
                    Directory.CreateDirectory(ConfigurationManager.AppSettings["rootPathBackup"] + folderDate + "\\wwwroot");
                }


                string dbConnection = ConfigurationManager.AppSettings["sqlconnection"];

                DataSet dt = new DataSet();

                DateTime dtStartZip = DateTime.Now;
                string zipError = null;

                if (File.Exists(ConfigurationManager.AppSettings["xmlDataFilePath"]))
                {
                    dt.ReadXml(ConfigurationManager.AppSettings["xmlDataFilePath"]);

                    Int64 totalZipSize = 0;
                    string zipFileSize = "";
                    string zipFilePath = "";

                    for (int i = 0; i < dt.Tables["website"].Rows.Count; i++)
                    {

                        string domainName = dt.Tables["website"].Rows[i][0].ToString();
                        string path = dt.Tables["website"].Rows[i][1].ToString();
                        string excludePath = dt.Tables["website"].Rows[i][2].ToString();
                        string dbName = dt.Tables["website"].Rows[i][3].ToString();

                        // Create Database Backup ----------------------------
                        if (!string.IsNullOrEmpty(dbName))
                        {
                            try
                            {
                                SqlConnection MySqlConnection = new SqlConnection();
                                MySqlConnection.ConnectionString = dbConnection;
                                MySqlConnection.Open();

                                string[] dbNames = dbName.Split(',');
                                for (int ii = 0; ii < dbNames.Length; ii++)
                                {
                                    SqlCommand c = new SqlCommand("[viraBackup] @databaseName = '" + dbNames[ii] + "' , @path = '" + ConfigurationManager.AppSettings["rootPathBackup"] + folderDate + "\\sql\\';", MySqlConnection);
                                    c.ExecuteNonQuery();
                                }

                                MySqlConnection.Close();
                            }
                            catch (Exception e)
                            {
                                SubmitLog("Error: " + e.Message);
                                errorCount += 1;
                            }
                        }

                        if (Convert.ToBoolean(ConfigurationManager.AppSettings["oneZipFile"]))
                        {
                            // Copy Files -----------------------------------------
                            if (!string.IsNullOrEmpty(path))
                            {
                                if (!Directory.Exists(ConfigurationManager.AppSettings["rootPathBackup"] + folderDate + "\\" + domainName))
                                {
                                    Directory.CreateDirectory(ConfigurationManager.AppSettings["rootPathBackup"] + folderDate + "\\wwwroot\\" + domainName);
                                }

                                try
                                {
                                    CopyDirectory(new DirectoryInfo(path), new DirectoryInfo(ConfigurationManager.AppSettings["rootPathBackup"] + folderDate + "\\wwwroot\\" + domainName));
                                }
                                catch (Exception e)
                                {
                                    SubmitLog("Error: " + e.Message);
                                    errorCount += 1;
                                }
                            }
                        }
                        else
                        {
                            // Direct Zip Folders -----------------------------------------
                            if (!string.IsNullOrEmpty(path))
                            {
                                if (!zip(path, ConfigurationManager.AppSettings["rootPathBackup"], folderDate + "\\wwwroot\\" + domainName + ".7z", ConfigurationManager.AppSettings["zipPassword"], excludePath, ref zipError))
                                {
                                    SubmitLog(zipError);
                                    errorCount += 1;
                                }
                                if (File.Exists(ConfigurationManager.AppSettings["rootPathBackup"] + folderDate + "\\wwwroot\\" + domainName + ".7z"))
                                {
                                    FileInfo ff = new FileInfo(ConfigurationManager.AppSettings["rootPathBackup"] + folderDate + "\\wwwroot\\" + domainName + ".7z");
                                    totalZipSize += ff.Length;
                                }
                            }
                        }
                    }



                    // One Zip File Process ---------------------------------------------------------------
                    if (Convert.ToBoolean(ConfigurationManager.AppSettings["oneZipFile"]))
                    {
                        dtStartZip = DateTime.Now;
                        zipError = null;
                        if (!zip(ConfigurationManager.AppSettings["rootPathBackup"] + folderDate, ConfigurationManager.AppSettings["rootPathBackup"], folderDate + ".7z", ConfigurationManager.AppSettings["zipPassword"], "", ref zipError))
                        {
                            SubmitLog(zipError);
                            errorCount += 1;
                        }
                        if (File.Exists(ConfigurationManager.AppSettings["rootPathBackup"] + folderDate + ".7z"))
                        {
                            FileInfo ff = new FileInfo(ConfigurationManager.AppSettings["rootPathBackup"] + folderDate + ".7z");
                            totalZipSize = ff.Length;
                        }
                    }
                    else
                    {
                        if (!zip(ConfigurationManager.AppSettings["rootPathBackup"] + folderDate + "\\sql", ConfigurationManager.AppSettings["rootPathBackup"], folderDate + "\\sql.7z", ConfigurationManager.AppSettings["zipPassword"], "", ref zipError))
                        {
                            SubmitLog(zipError);
                            errorCount += 1;
                        }
                        if (File.Exists(ConfigurationManager.AppSettings["rootPathBackup"] + folderDate + "\\sql.7z"))
                        {
                            FileInfo ff = new FileInfo(ConfigurationManager.AppSettings["rootPathBackup"] + folderDate + "\\sql.7z");
                            totalZipSize += ff.Length;
                        }
                    }

                    DateTime dtEndZip = DateTime.Now;
                    TimeSpan diffResultZip = dtEndZip.Subtract(dtStartZip);

                    zipFileSize = byteConvert(totalZipSize);

                    // FTP Process ---------------------------------------------------------------
                    DateTime dtStartFTP = DateTime.Now;
                    if (Convert.ToBoolean(ConfigurationManager.AppSettings["backupToFTPServer"]))
                    {
                        if (File.Exists(ConfigurationManager.AppSettings["rootPathBackup"] + folderDate + ".7z"))
                        {
                            FileInfo f = new FileInfo(ConfigurationManager.AppSettings["rootPathBackup"] + folderDate + ".7z");
                            zipFileSize = byteConvert(f.Length);
                            if (!upload(ConfigurationManager.AppSettings["ftpServerIP"],
                                        ConfigurationManager.AppSettings["ftpUserID"],
                                        ConfigurationManager.AppSettings["ftpPassword"],
                                        ConfigurationManager.AppSettings["rootPathBackup"] + folderDate + ".7z"))
                            {
                                SubmitLog("FTP Error");
                                errorCount += 1;
                                zipFilePath = "[NA]";
                            }
                            else
                            {
                                zipFilePath = ConfigurationManager.AppSettings["downloadZipPath"] + folderDate + ".7z";
                            }
                        }
                    }
                    DateTime dtEndFTP = DateTime.Now;
                    TimeSpan diffResultFTP = dtEndFTP.Subtract(dtStartFTP);


                    // Report ----------------------------------------------------------------------
                    DateTime dtEnd = DateTime.Now;
                    TimeSpan diffResult = dtEnd.Subtract(dtStart);

                    SubmitLog("[Process Duration]: " + diffResult.ToString() + "  - For Backup Hosting");
                    if (!string.IsNullOrEmpty(ConfigurationManager.AppSettings["smsProviderUrl"]))
                    {
                        SendSMS("+98" + ConfigurationManager.AppSettings["cell"], ConfigurationManager.AppSettings["serverName"] + "-" + errorCount.ToString(), "BackupDone");
                    }
                    else
                    {
                        SendEmail("[" + ConfigurationManager.AppSettings["serverName"] + "] Vira Backup Service Successfully Completed! " + errorCount.ToString() + " Error(s)",
                                  "Vira Backup service successfully completed!<br>" +
                                  "<b>" + ConfigurationManager.AppSettings["serverName"] + "</b> [" + System.Environment.MachineName + "]<br>" +
                                  "<hr/>" +
                                  "Process ID:" + workingID.ToString() + "<br>" +
                                  "Process Start Time: " + dtStart.ToString() + "<br>" +
                                  "Process Finish Time: " + DateTime.Now.ToString() + "<br>" +
                                  "Zip File size: " + zipFileSize + "<br>" +
                                  "Zip Process Duration: " + diffResultZip.ToString() + "<br>" +
                                  "FTP Process Duration: " + diffResultFTP.ToString() + "<br>" +
                                  "Total Process Duration: " + diffResult.ToString() + "<br>" +
                                  "Process error(s): " + errorCount.ToString() + "<br>" +
                                  "Backup to FTP Server: " + ConfigurationManager.AppSettings["backupToFTPServer"] + "<br>" +
                                  "Backup File download Path: " + zipFilePath + "<br>" +
                                  "<hr/>" + "viraBackup " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString() + "<br>" +
                                  "Powered by: https://www.viraw.ir");
                    }


                    if (Convert.ToBoolean(ConfigurationManager.AppSettings["deleteTempFiles"]))
                    {
                        if (Convert.ToBoolean(ConfigurationManager.AppSettings["oneZipFile"]))
                        {
                            if (Directory.Exists(ConfigurationManager.AppSettings["rootPathBackup"] + folderDate))
                            {
                                try
                                {
                                    Directory.Delete(ConfigurationManager.AppSettings["rootPathBackup"] + folderDate, true);
                                    SubmitLog(ConfigurationManager.AppSettings["rootPathBackup"] + folderDate + " Folder Deleted!");
                                }
                                catch (Exception ex)
                                {
                                    SubmitLog(ConfigurationManager.AppSettings["rootPathBackup"] + folderDate + " Folder Can Not Deleted!" + Environment.NewLine + ex.Message);
                                }
                            }
                        }
                        else
                        {
                            if (Directory.Exists(ConfigurationManager.AppSettings["rootPathBackup"] + folderDate + "\\sql"))
                            {
                                try
                                {
                                    Directory.Delete(ConfigurationManager.AppSettings["rootPathBackup"] + folderDate + "\\sql", true);
                                    SubmitLog(ConfigurationManager.AppSettings["rootPathBackup"] + folderDate + "\\sql Folder Deleted!");
                                }
                                catch (Exception ex)
                                {
                                    SubmitLog(ConfigurationManager.AppSettings["rootPathBackup"] + folderDate + "\\sql Folder Can Not Deleted!" + Environment.NewLine + ex.Message);
                                }
                            }
                        }
                    }
                }
                OnWorking = false;
            }
        }


        #region Custom Functions


        public bool SubmitLog(string LogString)
        {
            bool result = false;
            try
            {
                if (!Directory.Exists(@"c:\viraBackup_logs"))
                {
                    Directory.CreateDirectory(@"c:\viraBackup_logs");
                }

                string fname = @"c:\viraBackup_logs\" + DateTime.Now.Year + DateTime.Now.Month + DateTime.Now.Day + ".log";

                FileInfo MyFile = new FileInfo(fname);

                if (File.Exists(fname))
                {
                    StreamWriter sw = MyFile.AppendText();
                    sw.WriteLine(DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + ":" + DateTime.Now.Millisecond + " , " + LogString);
                    sw.Close();
                }
                else
                {
                    StreamWriter sw = MyFile.CreateText();
                    sw.WriteLine("viraBackup Logs//////////////////////////////////////////////////////////////////////////////////////");
                    sw.WriteLine(DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + ":" + DateTime.Now.Millisecond + " , " + LogString);
                    sw.Close();
                }

                result = true;
            }
            catch
            {
                result = false;
            }

            return result;
        }


        public static void SendEmail(string mSubject, string mMessage)
        {

            try
            {
                MailMessage email = new MailMessage();
                email.Subject = mSubject;
                //email.To.Add(new MailAddress("barjomandi@viraware.com"));
                string[] mailingList = System.Configuration.ConfigurationManager.AppSettings["emailsContact"].Split(';');
                for (int i = 0; i < mailingList.Length; i++)
                {
                    email.To.Add(new MailAddress(mailingList[i]));
                }
                email.Body = mMessage;
                email.Priority = MailPriority.High;
                email.IsBodyHtml = true;
                email.BodyEncoding = Encoding.UTF8;
                SmtpClient smtp = new SmtpClient();
                smtp.Send(email);
            }
            catch
            {
            }
        }


        public bool CheckPort(string IP, int Port, int Timeout)
        {
            try
            {
                TcpClient aSocket = new TcpClientWithTimeout(IP, Port, Timeout).Connect();

                if (aSocket.Connected)
                {
                    aSocket.Close();
                    aSocket = null;
                    return true;
                }
                else
                {
                    aSocket.Close();
                    aSocket = null;
                    return false;
                }
            }
            catch
            {
                return false;
            }
        }


        public bool ScanPort(IPAddress IP, int Port)
        {
            TcpClient aSocket;

            aSocket = new TcpClient();
            try
            {
                aSocket.Connect(IP, Port);
            }
            catch
            {
                // Something went wrong
                return false;
            }
            if (aSocket.Connected)
            {
                // Got connected to Address+Port
                aSocket.Close();
                aSocket = null;
                return true;
            }
            else
            {
                // Not connected
                aSocket.Close();
                aSocket = null;
                return false;
            }
        }


        public static void CopyDirectory(DirectoryInfo source, DirectoryInfo destination)
        {
            if (!destination.Exists)
            {
                destination.Create();
            }

            // Copy all files. 
            FileInfo[] files = source.GetFiles();
            foreach (FileInfo file in files)
            {
                file.CopyTo(Path.Combine(destination.FullName, file.Name),true);
            }

            // Process subdirectories. 
            DirectoryInfo[] dirs = source.GetDirectories();
            foreach (DirectoryInfo dir in dirs)
            {
                // Get destination directory. 
                string destinationDir = Path.Combine(destination.FullName, dir.Name);

                // Call CopyDirectory() recursively. 
                CopyDirectory(dir, new DirectoryInfo(destinationDir));
            }
        }


        public static bool upload(string ftpServerIP, string ftpUserID, string ftpPassword, string fileNamePath)
        {
            bool result = false;
            FileInfo fileInf = new FileInfo(fileNamePath);

            string uri = "ftp://" + ftpServerIP + "/" + fileInf.Name;

            FtpWebRequest reqFTP;

            reqFTP = (FtpWebRequest)FtpWebRequest.Create(new Uri("ftp://" + ftpServerIP + "/" + fileInf.Name));

            reqFTP.Credentials = new NetworkCredential(ftpUserID, ftpPassword);
            reqFTP.KeepAlive = false;
            reqFTP.Method = WebRequestMethods.Ftp.UploadFile;
            //reqFTP.EnableSsl = true;
            reqFTP.UseBinary = true;
            reqFTP.ContentLength = fileInf.Length;
            int buffLength = 2048;
            byte[] buff = new byte[buffLength];
            int contentLen;

            FileStream fs = fileInf.OpenRead();
            try
            {
                Stream strm = reqFTP.GetRequestStream();
                contentLen = fs.Read(buff, 0, buffLength);
                while (contentLen != 0)
                {
                    strm.Write(buff, 0, contentLen);
                    contentLen = fs.Read(buff, 0, buffLength);
                }
                strm.Close();
                fs.Close();
                result = true;
            }
            catch (Exception ex)
            {

            }

            return result;
        }



        public static bool zip(string SourcePath, string DestinationPath, string zipFileName, string zipPassword, string excludeFolder, ref string resultOutput)
        {
            bool result = false;
            try
            {
                if (DestinationPath.Substring(DestinationPath.Length - 1, 1) != "\\")
                {
                    DestinationPath = DestinationPath + "\\";
                }
                System.Diagnostics.Process p = new System.Diagnostics.Process();
                // Redirect the output stream of the child process.
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.RedirectStandardOutput = true;
                p.StartInfo.FileName = ConfigurationManager.AppSettings["7zipPath"];

                if (!string.IsNullOrEmpty(excludeFolder))
                {
                    excludeFolder = " -xr!" + excludeFolder + " ";
                }

                if (string.IsNullOrEmpty(zipPassword))
                {
                    p.StartInfo.Arguments = " a " + DestinationPath + zipFileName + " " + SourcePath + excludeFolder;
                }
                else
                {
                    p.StartInfo.Arguments = " a " + DestinationPath + zipFileName + " " + SourcePath + excludeFolder + " -p" + zipPassword;
                }
                p.Start();
                // Do not wait for the child process to exit before
                // reading to the end of its redirected stream.
                // p.WaitForExit();
                // Read the output stream first and then wait.
                resultOutput = p.StandardOutput.ReadToEnd();
                p.WaitForExit();
                if (File.Exists(DestinationPath + zipFileName))
                {
                    result = true;
                }
                else
                {
                    resultOutput = DestinationPath + zipFileName + "not found!";
                }
            }
            catch (Exception ex)
            {
                resultOutput = ex.Message;
            }
            return result;
        }


        public static string byteConvert(double byteVar)
        {
            double Kb = 1024;
            double Mb = Kb * 1024;
            double Gb = Mb * 1024;
            double Tb = Gb * 1024;
            double result = byteVar;
            string unit = "Byte";

            if (byteVar > Kb && byteVar <= Mb) { result = byteVar / Kb; unit = "Kb"; }
            if (byteVar > Mb && byteVar <= Gb) { result = byteVar / Mb; unit = "Mb"; }
            if (byteVar > Gb && byteVar <= Tb) { result = byteVar / Gb; unit = "Gb"; }
            if (byteVar > Tb) { result = byteVar / Tb; unit = "Tb"; }
            return result.ToString("0.00") + " " + unit;
        }


        #endregion


    }
}
