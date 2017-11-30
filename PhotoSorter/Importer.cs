using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using Goheer.EXIF;
using PicImp;

namespace PhotoSorter
{
    class Importer
    {
        private string importDir { get; set; }
        private string picDir { get; set; }
        private string videoDir { get; set; }

        internal void RunImport()
        {
            SQLiteConnection m_dbConnection = new SQLiteConnection("Data Source=PhotoSorter.db;Version=3;");
            try
            {
                //Get Data
                m_dbConnection.Open();
                string sql = "SELECT * FROM tbl_configs";
                SQLiteCommand command = new SQLiteCommand(sql, m_dbConnection);
                SQLiteDataReader reader = command.ExecuteReader();
                string prefix = "";
                string suffix = "";
                bool useCameraMake = false;
                string sfUserID = "";
                string sfPasswd = "";
                string sfAppID = "";
                string sfSS = "";
                string sfAuthID = "";
                while (reader.Read())
                {
                    this.importDir = Environment.ExpandEnvironmentVariables(reader["importDirectory"].ToString());
                    this.picDir = Environment.ExpandEnvironmentVariables(reader["pictureDirectory"].ToString());
                    this.videoDir = Environment.ExpandEnvironmentVariables(reader["videoDirectory"].ToString());
                    prefix = reader["prefix"].ToString();
                    suffix = reader["suffix"].ToString();
                    if (reader["includeCameraMake"] != DBNull.Value)
                    {
                        useCameraMake = Convert.ToBoolean(reader["includeCameraMake"]);
                    }
                    sfUserID = reader["sfUserID"].ToString();
                    if (reader["sfPasswd"] != DBNull.Value)
                    {
                        sfPasswd = Encrypting.ToInsecureString(Encrypting.DecryptString(reader["sfPasswd"].ToString()));
                    }
                    sfAppID = reader["sfAppID"].ToString();
                    if (reader["sfSS"] != DBNull.Value)
                    {
                        sfSS = Encrypting.ToInsecureString(Encrypting.DecryptString(reader["sfSS"].ToString()));
                    }
                    sfAuthID = reader["sfAuthID"].ToString();
                }
                reader.Close();

                //Initialize Shutterfly
                if (sfUserID != "" && sfPasswd != "" && sfAppID != "" && sfSS != "")
                {
                    string authenticationID = Shutterfly.getAuthenticationID(sfUserID, sfPasswd, sfAppID, sfSS, sfAuthID);
                    if (!authenticationID.StartsWith("Failed:") && authenticationID != sfAuthID)
                    {
                        string sfsql = String.Format("UPDATE tbl_configs SET sfAuthID = '{0}' WHERE user_pk = 1", authenticationID);
                        SQLiteCommand sqlcmd = new SQLiteCommand(sql, m_dbConnection);
                        sqlcmd.ExecuteNonQuery();
                        sfAuthID = authenticationID;
                    }
                }

                //Get video extensions
                List<string> videoEXT = new List<string>();
                string query = "SELECT * FROM tbl_videos ORDER BY extension";
                SQLiteCommand cmd = new SQLiteCommand(query, m_dbConnection);
                SQLiteDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    videoEXT.Add(rdr["extension"].ToString());
                }
                rdr.Close();

                if (Directory.Exists(importDir))
                {
                    if (!Directory.Exists(picDir))
                    {
                        Directory.CreateDirectory(picDir);
                    }

                    if (!Directory.Exists(videoDir))
                    {
                        Directory.CreateDirectory(videoDir);
                    }

                    FileInfo[] Pictures = (from fi in new DirectoryInfo(importDir).GetFiles("*.*", SearchOption.AllDirectories)
                                           where !videoEXT.Contains(fi.Extension.ToLower())
                                           select fi)
                                                .ToArray();

                    FileInfo[] Videos = (from fi in new DirectoryInfo(importDir).GetFiles("*.*", SearchOption.AllDirectories)
                                         where videoEXT.Contains(fi.Extension.ToLower())
                                         select fi)
                                                .ToArray();

                    List<string> moveErrors = new List<string>();
                    List<string> uploadErrors = new List<string>();

                    foreach (FileInfo pictureFile in Pictures)
                    {
                        processPicture(pictureFile, sfAuthID, sfAppID, prefix, suffix, useCameraMake, ref moveErrors, ref uploadErrors);
                    }

                    foreach (FileInfo videoFile in Videos)
                    {
                        processVideo(videoFile, prefix, suffix, useCameraMake, ref moveErrors);
                    }

                    if (moveErrors.Count > 0 || uploadErrors.Count > 0)
                    {
                        //sendEmail(moveErrors, uploadErrors);
                    }

                    cleanUp();
                }
                else
                {
                    string error = "Import directory does not exist. Check your settings.";
                }
            }
            catch (Exception ex) { }
            finally
            {
                m_dbConnection.Close();
            }
        }

        private void cleanUp()
        {
            DirectoryInfo[] subDirs = (from di in new DirectoryInfo(importDir).GetDirectories("*.*", SearchOption.AllDirectories)
                                       where (Directory.EnumerateFileSystemEntries(di.FullName).Any())
                                       select di)
                                             .ToArray();

            foreach (DirectoryInfo dir in subDirs)
            {
                if (File.Exists(dir.FullName + "\\Thumbs.db"))
                {
                    File.Delete(dir.FullName + "\\Thumbs.db");
                }
                Directory.Delete(dir.FullName);
            }
        }

        //private static string sendEmail(List<string> moveErrors, List<string> uploadErrors)
        //{
        //    try
        //    {
        //        string smtpServer = ConfigurationManager.AppSettings["smtpServer"].ToString();
        //        string fromEmail = ConfigurationManager.AppSettings["fromEmail"].ToString();
        //        string eFromPsswd = ConfigurationManager.AppSettings["fromPsswd"].ToString();
        //        SecureString securedPsswd = Encrypting.DecryptString(eFromPsswd);
        //        string toEmail = ConfigurationManager.AppSettings["adminEmail"].ToString();

        //        MailMessage msg = new MailMessage();
        //        msg.IsBodyHtml = true;
        //        msg.From = new MailAddress(fromEmail);
        //        msg.To.Add(new MailAddress(toEmail));

        //        msg.Subject = "PicImp error moving or uploading";
        //        StringBuilder sb = new StringBuilder();

        //        if (moveErrors.Count > 0)
        //        {
        //            sb.AppendLine("Error moving the following files:");
        //            foreach (string mError in moveErrors)
        //            {
        //                sb.AppendLine("\t" + mError);
        //            }
        //        }

        //        if (uploadErrors.Count > 0)
        //        {
        //            sb.AppendLine("Error uploading the following files:");
        //            foreach (string uError in uploadErrors)
        //            {
        //                sb.AppendLine("\t" + uError);
        //            }
        //        }

        //        string body = System.Net.WebUtility.HtmlEncode(sb.ToString());
        //        //HTML encode does not modify whitespace
        //        body = body.Replace("\n", "<br>");
        //        body = body.Replace("\t", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;");

        //        msg.Body = body;
        //        SmtpClient smtp = new SmtpClient
        //        {
        //            Host = smtpServer,
        //            Port = 587,
        //            EnableSsl = true,
        //            UseDefaultCredentials = false,
        //            Credentials = new System.Net.NetworkCredential(fromEmail, Encrypting.ToInsecureString(securedPsswd))
        //        };
        //        smtp.Send(msg);

        //        return "true";
        //    }
        //    catch (Exception ex)
        //    {
        //        return ex.ToString();
        //    }
        //}

        private void processVideo(FileInfo videoFile, string prefix, string suffix, bool useCameraMake, ref List<string> moveErrors)
        {
            DateTime myDateTaken = videoFile.CreationTime;
            string ext = videoFile.Extension.ToLower();
            StringBuilder newFileName = new StringBuilder();
            newFileName.Append(prefix);
            newFileName.Append(myDateTaken.ToString("yyyyMMdd_HHmmss"));
            newFileName.Append(suffix);

            Bitmap bmp;
            try
            {
                bmp = new Bitmap(videoFile.FullName);
            }
            catch
            {
                moveErrors.Add(videoFile.Name);
                //moveToManualFolder(videoFile.FullName, importDir, newFileName.ToString(), ext);
                return;
            }
            EXIFextractor exif = new EXIFextractor(ref bmp, "n"); // get source from http://www.codeproject.com/KB/graphics/exifextractor.aspx?fid=207371
            if (useCameraMake)
            {
                string cameraModel = GetCameraModel(exif);
                if (cameraModel != "")
                {
                    newFileName.Append("_");
                    newFileName.Append(cameraModel);
                }
            }

            string moveToPath = CreateVidDirStructure(ref myDateTaken);
            

            //Attempt to move
            if (moveToPath != "Error")
            {
                try
                {
                    string result = moveVideo(videoFile.FullName, moveToPath, newFileName.ToString(), ext);
                    if (result == "Error")
                    {
                        moveErrors.Add(videoFile.Name);
                        moveToManualFolder(videoFile.FullName, importDir, newFileName.ToString(), ext);
                    }
                }
                catch
                {
                    moveErrors.Add(videoFile.Name);
                    moveToManualFolder(videoFile.FullName, importDir, newFileName.ToString(), ext);
                }
            }
            else
            {
                moveErrors.Add(videoFile.Name);
                moveToManualFolder(videoFile.FullName, importDir, newFileName.ToString(), ext);
            }
        }

        private void processPicture(FileInfo pictureFile, string authenticationID, string sfAppID, string prefix, string suffix, bool useCameraMake, ref List<string> moveErrors, ref List<string> uploadErrors)
        {
            if (pictureFile.Name == "Thumbs.db" || pictureFile.Extension == ".ini")
            {
                return;   //skip
            }

            //Set Defaults
            DateTime myDateTaken = DateTime.Now;
            StringBuilder newFileName = new StringBuilder();
            newFileName.Append(prefix);
            newFileName.Append(myDateTaken.ToString("yyyyMMdd_HHmmss"));
            string ext = pictureFile.Extension.ToLower();
            string moveToPath = CreatePicDirStructure(ref myDateTaken);

            Bitmap bmp;
            try
            {
                bmp = new Bitmap(pictureFile.FullName);
            }
            catch
            {
                moveErrors.Add(pictureFile.Name);
                moveToManualFolder(pictureFile.FullName, importDir, newFileName.ToString(), ext);
                return;
            }

            EXIFextractor exif = new EXIFextractor(ref bmp, "n"); // get source from http://www.codeproject.com/KB/graphics/exifextractor.aspx?fid=207371

            //Get Date Taken
            if (exif["Date Time"] != null)
            {
                myDateTaken = DateTime.ParseExact(exif["Date Time"].ToString().TrimEnd('\0'), "yyyy:MM:dd HH:mm:ss", null);
                newFileName.Clear();
                newFileName.Append(prefix);
                newFileName.Append(myDateTaken.ToString("yyyyMMdd_HHmmss"));
                moveToPath = CreatePicDirStructure(ref myDateTaken);
            }

            //Auto Rotate
            autoRotate(ref exif, ref bmp, pictureFile);

            bmp.Dispose();

            if (useCameraMake)
            {
                string cameraModel = GetCameraModel(exif);
                if (cameraModel != "")
                {
                    newFileName.Append("_");
                    newFileName.Append(cameraModel);
                }
            }

            newFileName.Append(suffix);

            string result = "";
            //Attempt to move
            if (moveToPath != "Error")
            {
                try
                {
                    result = movePicture(pictureFile.FullName, moveToPath, newFileName.ToString(), ext);
                }
                catch
                {
                    moveErrors.Add(pictureFile.Name);
                    moveToManualFolder(pictureFile.FullName, importDir, newFileName.ToString(), ext);
                }

                if (result == "Error")
                {
                    moveErrors.Add(pictureFile.Name);
                    moveToManualFolder(pictureFile.FullName, importDir, newFileName.ToString(), ext);
                }
                else
                {
                    if (!authenticationID.StartsWith("Failed") && authenticationID != "")
                    {
                        string uploadResult = uploadPicture(authenticationID, sfAppID, myDateTaken, moveToPath, newFileName.ToString(), ext);
                        if (uploadResult.StartsWith("Failed"))
                        {
                            uploadErrors.Add(pictureFile.Name);
                        }
                    }
                    else
                    {
                        uploadErrors.Add(pictureFile.Name);
                    }
                }

            }
            else
            {
                moveErrors.Add(pictureFile.Name);
                moveToManualFolder(pictureFile.FullName, importDir, newFileName.ToString(), ext);
            }

            result = "Successfully moved file " + pictureFile.Name + " to " + moveToPath + " as " + newFileName.ToString() + ext;
        }

        private string GetCameraModel(EXIFextractor exif)
        {
            string cameraModel = "";
            if (exif["Equip Model"] != null)
            {
                if (exif["Equip Model"] != null)
                {
                    cameraModel = exif["Equip Model"].ToString().TrimEnd('\0').Trim();
                }
            }
            else
            {
                if (exif["Equip Make"] != null)
                {
                    cameraModel = exif["Equip Make"].ToString().TrimEnd('\0').Trim();
                }
            }
            return cameraModel;
        }

        private void autoRotate(ref EXIFextractor exif, ref Bitmap bmp, FileInfo file)
        {
            //Auto-Rotate file
            if (exif["Orientation"] != null)
            {
                string orient = exif["Orientation"].ToString();
                RotateFlipType flip = OrientationToFlipType(exif["Orientation"].ToString().TrimEnd('\0'));

                if (flip != RotateFlipType.RotateNoneFlipNone) // don't flip of orientation is correct
                {
                    bmp.RotateFlip(flip);

                    // Optional: reset orientation tag
                    try
                    {
                        //exif.setTag(0x112, "1");  doesn't work
                        int propertyId = 0x112;
                        int propertyLen = 2;
                        short propertyType = 3;
                        byte[] propertyValue = { 1, 0 };
                        exif.setTag(propertyId, propertyLen, propertyType, propertyValue);
                    }
                    catch { }

                    switch (file.Extension.ToLower())
                    {
                        case ".png":
                            bmp.Save(file.FullName, ImageFormat.Png);
                            break;
                        case ".bmp":
                            bmp.Save(file.FullName, ImageFormat.Bmp);
                            break;
                        case ".gif":
                            bmp.Save(file.FullName, ImageFormat.Gif);
                            break;
                        default:
                            bmp.Save(file.FullName, ImageFormat.Jpeg);
                            break;
                    }
                }
            }
        }     

    private void moveToManualFolder(string fromFilePath, string moveToPath, string fileName, string ext)
        {
            moveToPath = moveToPath + "\\ManualMove";
            if (!Directory.Exists(moveToPath))
            {
                Directory.CreateDirectory(moveToPath);
            }

            if (!File.Exists(moveToPath + "\\" + fileName + ext))
            {
                File.Copy(fromFilePath, moveToPath + "\\" + fileName + ext);
            }
            else
            {
                int i;
                for (i = 1; i < 10; i++)
                {
                    if (!File.Exists(moveToPath + "\\" + fileName + "_" + i + ext))
                    {
                        File.Copy(fromFilePath, moveToPath + "\\" + fileName + "_" + i + ext);
                        break;
                    }
                }
                if (i >= 10)
                {
                    File.Copy(fromFilePath, moveToPath + "\\" + fileName + "_" + DateTime.Now.ToString("MMddyyyyThhmmssffftt") + ext);
                }
            }
        }

        private string movePicture(string fromFilePath, string toPath, string fileName, string ext)
        {
            if (!File.Exists(toPath + "\\" + fileName + ext))
            {
                File.Move(fromFilePath, toPath + "\\" + fileName + ext);
            }
            else
            {
                int i;
                for (i = 1; i < 10; i++)
                {
                    if (!File.Exists(toPath + "\\" + fileName + "_" + i + ext))
                    {
                        File.Move(fromFilePath, toPath + "\\" + fileName + "_" + i + ext);
                        break;
                    }
                }
                if (i >= 10)
                {
                    return "Error";
                }
            }

            return "Success";
        }

        private string uploadPicture(string authenticationID, string sfAppID, DateTime dateTaken, string newPath, string fileName, string ext)
        {
            System.Globalization.DateTimeFormatInfo mfi = new System.Globalization.DateTimeFormatInfo();
            string monthName = mfi.GetMonthName(dateTaken.Month).ToString();
            return Shutterfly.uploadToShutterfly(authenticationID, sfAppID, monthName, dateTaken.Year.ToString(), newPath, fileName + ext);
        }

        private string moveVideo(string fromFilePath, string toPath, string fileName, string ext)
        {
            if (!File.Exists(toPath + "\\" + fileName + ext))
            {
                File.Move(fromFilePath, toPath + "\\" + fileName + ext);
            }
            else
            {
                int i;
                for (i = 1; i < 10; i++)
                {
                    if (!File.Exists(toPath + "\\" + fileName + "_" + i + ext))
                    {
                        File.Move(fromFilePath, toPath + "\\" + fileName + "_" + i + ext);
                        break;
                    }
                }
                if (i >= 10)
                {
                    return "Error";
                }
            }
            return "Success";
        }

        private string CreatePicDirStructure(ref DateTime dt)
        {
            string year = dt.Year.ToString();
            string monthNum = dt.Month.ToString().PadLeft(2, '0');
            string monAbbrv = dt.ToString("MMM");

            try
            {
                if (!Directory.Exists(picDir + "\\" + year))
                {
                    Directory.CreateDirectory(picDir + "\\" + year);
                }
                if (!Directory.Exists(picDir + "\\" + year + "\\" + monthNum + "_" + monAbbrv))
                {
                    Directory.CreateDirectory(picDir + "\\" + year + "\\" + monthNum + "_" + monAbbrv);
                }
            }
            catch
            {
                return "Error";
            }

            return picDir + "\\" + year + "\\" + monthNum + "_" + monAbbrv;
        }

        private string CreateVidDirStructure(ref DateTime dt)
        {
            string year = dt.Year.ToString();
            try
            {
                if (!Directory.Exists(videoDir + "\\" + year))
                {
                    Directory.CreateDirectory(videoDir + "\\" + year);
                }
            }
            catch
            {
                return "Error";
            }

            return videoDir + "\\" + year;
        }

        private RotateFlipType OrientationToFlipType(string orientation)
        {
            switch (orientation)
            {
                case "1":
                case "1H":
                    return RotateFlipType.RotateNoneFlipNone;
                case "2":
                    return RotateFlipType.RotateNoneFlipX;
                case "3":
                    return RotateFlipType.Rotate180FlipNone;
                case "4":
                    return RotateFlipType.Rotate180FlipX;
                case "5":
                    return RotateFlipType.Rotate90FlipX;
                case "6":
                    return RotateFlipType.Rotate90FlipNone;
                case "7":
                    return RotateFlipType.Rotate270FlipX;
                case "8":
                    return RotateFlipType.Rotate270FlipNone;
                default:
                    return RotateFlipType.RotateNoneFlipNone;
            }
        }
    }
}
