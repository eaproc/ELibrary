using System;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Runtime.Serialization.Formatters.Soap;
using ELibrary.Standard.Modules;
using Microsoft.VisualBasic;

namespace ELibrary.Standard
{

    /// <summary>
/// Static Class for I/O functions
/// </summary>
/// <remarks></remarks>
    public class EIO
    {
        public const string WINDOWS_TEMP_FOLDER = @"C:\Windows\Temp";






        /// <summary>
    /// Try to empty a directory withought deleting the directory itself
    /// </summary>
    /// <param name="direc"></param>
    /// <remarks></remarks>
        public static void DeleteAllFromDirectory(string direc)
        {
            var SubDirs = Directory.GetDirectories(direc);
            var SubFiles = Microsoft.VisualBasic.FileIO.FileSystem.GetFiles(direc).ToArray();
            if (SubDirs is object)
            {
                foreach (string subDir in SubDirs)
                {
                    try
                    {
                        Directory.Delete(subDir, true);
                    }
                    catch (Exception ex)
                    {
                    }
                }
            }

            if (SubFiles is object)
            {
                foreach (string subFile in SubFiles)
                {
                    try
                    {
                        File.Delete(subFile);
                    }
                    catch (Exception ex)
                    {
                    }
                }
            }
        }


        /// <summary>
    /// Deletes a directory and all its contents if exist. It ignores error on delete
    /// </summary>
    /// <param name="DirectoryPath"></param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static bool DeleteLongPathDirectory(string DirectoryPath)
        {
            try
            {
                {
                    var withBlock = new MicrosoftOS.CommandPrompt(false, false);
                    return withBlock.Execute(string.Format("rmdir /S /Q \"{0}\" ", DirectoryPath));
                }
            }
            catch (Exception ex)
            {
                return false;
            }

            return true;
        }

        /// <summary>
    /// Deletes a directory and all its contents if exist. It ignores error on delete
    /// </summary>
    /// <param name="DirectoryPath"></param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static bool DeleteDirectory(string DirectoryPath)
        {
            try
            {
                if (Microsoft.VisualBasic.FileIO.FileSystem.DirectoryExists(DirectoryPath))
                {
                    Microsoft.VisualBasic.FileIO.FileSystem.DeleteDirectory(DirectoryPath, Microsoft.VisualBasic.FileIO.DeleteDirectoryOption.DeleteAllContents);
                }
            }
            catch (Exception ex)
            {
                return false;
            }

            return true;
        }

        /// <summary>
    /// Deletes a directory and all its contents if exist and recreate it as empty directory. Creates new if it doesnt exist
    /// </summary>
    /// <param name="DirectoryPath"></param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static bool DeleteAndRecreateDirectory(string DirectoryPath)
        {
            try
            {
                if (Microsoft.VisualBasic.FileIO.FileSystem.DirectoryExists(DirectoryPath))
                {
                    Microsoft.VisualBasic.FileIO.FileSystem.DeleteDirectory(DirectoryPath, Microsoft.VisualBasic.FileIO.DeleteDirectoryOption.DeleteAllContents);
                }

                Microsoft.VisualBasic.FileIO.FileSystem.CreateDirectory(DirectoryPath);
            }
            catch (Exception ex)
            {
                return false;
            }

            return true;
        }

        /// <summary>
    /// Deletes a file if it exists
    /// </summary>
    /// <param name="_FileName"></param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static bool DeleteFileIfExists(string _FileName)
        {
            try
            {
                if (Microsoft.VisualBasic.FileIO.FileSystem.FileExists(_FileName))
                    Microsoft.VisualBasic.FileIO.FileSystem.DeleteFile(_FileName);
            }
            catch (Exception ex)
            {
                return false;
            }

            return true;
        }






        #region Path Naming




        /// <summary>
    /// Gets the extension of a file. It returns (.) with the extension name
    /// </summary>
    /// <param name="__FileName"></param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static string GetFileExtension(string __FileName)
        {
            try
            {
                return Path.GetExtension(__FileName);
            }
            catch (Exception ex)
            {
            }

            return Constants.vbNullString;
        }



        /// <summary>
    /// Fetches the directory name.. using backward slash. Only the current directoy like [MyFolder] is returned not full directory
    /// </summary>
    /// <param name="DirectoryPath"></param>
    /// <param name="IncludeSlash"></param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static string GetDirectoryName(string DirectoryPath, bool IncludeSlash = false)
        {
            var Dir_Path = Strings.Split(DirectoryPath, @"\");
            if (IncludeSlash)
            {
                return @"\" + Dir_Path[Dir_Path.Length - 1];
            }
            else
            {
                return Dir_Path[Dir_Path.Length - 1];
            }
        }

        /// <summary>
    /// Get's a directory full path from a file full path
    /// </summary>
    /// <param name="FilePath"></param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static string getDirectoryFullPath(string FilePath)
        {
            try
            {
                return Microsoft.VisualBasic.FileIO.FileSystem.GetParentPath(FilePath);
            }
            catch (Exception ex)
            {
                return Constants.vbNullString;
            }
        }

        /// <summary>
    /// Get a file name without extension
    /// </summary>
    /// <param name="__FileName">a File Name without the full path</param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static string getFileNameWithoutExtension(string __FileName)
        {
            string fName = Path.GetFileName(__FileName);
            return Strings.Left(fName, Strings.Len(fName) - Strings.Len(Path.GetExtension(__FileName)));
        }

        /// <summary>
    /// Get the file name in the provided uri
    /// </summary>
    /// <param name="uri"></param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static string getFileName(Uri uri)
        {
            try
            {
                return Microsoft.VisualBasic.FileIO.FileSystem.GetName(uri.AbsolutePath);
            }
            catch (Exception ex)
            {
                Program.Logger.Print(uri.ToString(), ex);
            }

            return Constants.vbNullString;
        }


        /// <summary>
    /// Get the file name in the provided uri
    /// </summary>
    /// <param name="URL"></param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static string getFileName(string URL)
        {
            try
            {
                return Microsoft.VisualBasic.FileIO.FileSystem.GetName(URL);
            }
            catch (Exception ex)
            {
                Program.Logger.Print(URL, ex);
            }

            return Constants.vbNullString;
        }






        // Check the Modeling file for more explanation


        /// <summary>
    /// get a suggested file full path name. like c:\...File.txt, c:\...File___1.txt
    /// </summary>
    /// <param name="pIntendedFileFullPath">the intended file name</param>
    /// <param name="pIncrementSeparator">if you want a separator between filename and increment. NB: Increment must be file naming compatible.</param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static string suggestFileUniqueFullFilePathName(string pIntendedFileFullPath, string pIncrementSeparator = "___")
        {
            return suggestFileUniqueFullFilePathName(getDirectoryFullPath(pIntendedFileFullPath), GetFileExtension(pIntendedFileFullPath), pIncrementSeparator);
        }


        /// <summary>
    /// get a suggested file full path name. like c:\...File.txt, c:\...File___1.txt
    /// </summary>
    /// <param name="pIntendedFileFolderWithBackSlash">The directory you wish to create the file in with back slash C:\</param>
    /// <param name="intendedFileName">the intended file name</param>
    /// <param name="IncrementSeparator">if you want a separator between filename and increment. NB: Increment must be file naming compatible.</param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static string suggestFileUniqueFullFilePathName(string pIntendedFileFolderWithBackSlash, string intendedFileName, string IncrementSeparator)
        {
            try
            {

                // REM Path.getExtension returns dot as well like .csv

                if (!Directory.Exists(pIntendedFileFolderWithBackSlash))
                    return string.Empty;
                if (!pIntendedFileFolderWithBackSlash.EndsWith(@"\"))
                    pIntendedFileFolderWithBackSlash += @"\";
                if (!File.Exists(pIntendedFileFolderWithBackSlash + intendedFileName))
                    return pIntendedFileFolderWithBackSlash + intendedFileName;
                var Increments = Directory.GetFiles(pIntendedFileFolderWithBackSlash, string.Format("{0}{1}*{2}", Path.GetFileNameWithoutExtension(intendedFileName), IncrementSeparator, Path.GetExtension(intendedFileName)));
                if (Increments is null || Increments.Length == 0)
                    return string.Format("{0}{1}{2}1{3}", pIntendedFileFolderWithBackSlash, Path.GetFileNameWithoutExtension(intendedFileName), IncrementSeparator, Path.GetExtension(intendedFileName));


                // REM It is always sorted according to name ascending order by default [NOT GAURANTEED]
                // Array.Sort(Increments)  REM Gaurantee sorting

                Increments = Increments.ToList().OrderByDescending(x => x).ToArray();
                string pTopmostIncrementFileFullPath = Increments.First();


                // REM Just the filename without folder path and extension
                string pIntendedFileNameWithoutExtension = Path.GetFileNameWithoutExtension(intendedFileName);



                // Fetching the incrementValue
                int pIncrement = 0;

                // 
                if (pIntendedFileNameWithoutExtension.equalsIgnoreCase(Path.GetFileNameWithoutExtension(pTopmostIncrementFileFullPath)))
                {
                    pIncrement = 1;
                }
                else
                {
                    string pTopmostIncrementFileName = Path.GetFileNameWithoutExtension(pTopmostIncrementFileFullPath);
                    string pDifference = pTopmostIncrementFileName.Substring(pIntendedFileNameWithoutExtension.Length, pTopmostIncrementFileName.Length - pIntendedFileNameWithoutExtension.Length);
                    pIncrement = InputsParsing.TextParsing.parseOutIntegers(pDifference).toInt32() + 1;
                }

                return string.Format("{0}{1}{2}{3}{4}", pIntendedFileFolderWithBackSlash, Path.GetFileNameWithoutExtension(intendedFileName), IncrementSeparator, pIncrement, Path.GetExtension(intendedFileName));
            }
            catch (Exception ex)
            {
                basMain.MyLogFile.Print("suggestFileUniqueFullFilePathName", ex);
            }

            return string.Empty;
        }


        /// <summary>
    /// Scan a directory for available file name you can use and returns full path of the available file name
    /// </summary>
    /// <param name="FolderPath"></param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static string getAvailableFileName(string FolderPath, string suggestedFileName = "File.dat")
        {
            long iCount = 0L;
            if (!Directory.Exists(FolderPath))
                return "";
            string __ext = new FileInfo(suggestedFileName).Extension;
            string sName = getFileNameWithoutExtension(suggestedFileName);
            string chkFile = sName;


            // REM Last slash if it exist
            if (Strings.Right(FolderPath, 1) == @"\")
                FolderPath = Strings.Left(FolderPath, Strings.Len(FolderPath) - 1);
            while (true)
            {
                if (File.Exists(string.Format(@"{0}\{1}{2}", FolderPath, chkFile, __ext)))
                {
                    iCount += 1L;
                    chkFile = string.Format("{0}_{1}", sName, iCount);
                }
                else
                {
                    return string.Format(@"{0}\{1}{2}", FolderPath, chkFile, __ext);
                }
            }

            return Constants.vbNullString;
        }




        #endregion






        /// <summary>
    /// Return the File path with the file name for the type that was sent
    /// </summary>
    /// <param name="objType"></param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static string get_Code_File_StartUp_File_FullPathName(Type objType)
        {
            return objType.Assembly.Location;
        }

        /// <summary>
    /// Return the Only File path WITHOUT the file name for the type that was sent
    /// </summary>
    /// <param name="objType"></param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static string get_Code_File_StartUp_Path(Type objType)
        {
            return Directory.GetParent(objType.Assembly.Location).ToString();
        }

        /// <summary>
    /// Loads a File (.exe,.dll) and reads it Assembly for GUID attributes
    /// </summary>
    /// <param name="FilePath"></param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static string get_GUID_From_File(string FilePath)
        {
            try
            {
                var fLoadedByte = File.ReadAllBytes(FilePath);
                var asy = System.Reflection.Assembly.Load(fLoadedByte);
                return ((System.Runtime.InteropServices.GuidAttribute)asy.GetCustomAttributes(typeof(System.Runtime.InteropServices.GuidAttribute), true)[0]).Value;
            }
            catch (Exception ex)
            {
                return string.Empty;
            }
        }



        /// <summary>
    /// Confirms if a directory is writable by this current application. [If it has the access level]
    /// </summary>
    /// <param name="FolderPath"></param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static bool isFolderWritableAndReadable(string FolderPath)
        {
            try
            {
                File.WriteAllText(string.Format(@"{0}\{1}", FolderPath, "Test.dll"), "Test");
                if (File.ReadAllText(string.Format(@"{0}\{1}", FolderPath, "Test.dll")) == "Test")
                {
                    DeleteFileIfExists(string.Format(@"{0}\{1}", FolderPath, "Test.dll"));
                    return true;
                }
            }
            catch (Exception ex)
            {
            }

            return false;
        }


        /// <summary>
    /// Checks if this a file by checking File.GetAttributes(__FileName) NOT Equals FileAttributes.Directory
    /// </summary>
    /// <param name="__FileName">a File Name without the full path</param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static bool isThis_a_file(string __FileName)
        {
            try
            {
                var p = File.GetAttributes(__FileName);
                return p != FileAttributes.Directory && (int)p != (int)(FileAttribute.Directory | FileAttribute.Archive);
            }
            catch (Exception ex)
            {
                basMain.MyLogFile.Print(ex);
                return false;
            }
        }






        #region Serialization


        #region SOAP

        // SOAP Serialization is great for deep serialization.
        // It requires implementation of ISerializable and/or Serializable attribute
        // 
        // It is same as Binary Serializable with the addition of portability
        // It is portable because it only depends on the class.


        /// <summary>
    /// Writes Serializable Object to File. Throws all exceptions
    /// </summary>
    /// <param name="obj"></param>
    /// <param name="FileFullPath"></param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static bool WriteToFileSOAP(object obj, string FileFullPath)
        {
            // REM Create path first
            if (!Directory.Exists(Microsoft.VisualBasic.FileIO.FileSystem.GetParentPath(FileFullPath)))
                Directory.CreateDirectory(Microsoft.VisualBasic.FileIO.FileSystem.GetParentPath(FileFullPath));
            var fs = new FileStream(FileFullPath, FileMode.Create);
            var objSerializer = new SoapFormatter();
            objSerializer.Serialize(fs, obj);
            fs.Close();
            return true;
        }

        /// <summary>
    /// Reads a serialized object from file. Throws all exceptions
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="FileFullPath"></param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static T ReadFromFileSOAP<T>(string FileFullPath)
        {
            var fs = new FileStream(FileFullPath, FileMode.Open, FileAccess.Read);
            var objSerializer = new SoapFormatter();
            T rst = (T)objSerializer.Deserialize(fs);
            fs.Close();
            return rst;
        }

        #endregion




        #region Binary

        // Binary Serialization is great for deep serialization
        // It requires implementation of ISerializable and/or Serializable attribute
        // 
        // It is NOT portable because it is deep and depends fully on the Namespace.Class that wrote it.


        /// <summary>
    /// Writes Serializable Object to File. Throws all exceptions
    /// </summary>
    /// <param name="obj"></param>
    /// <param name="FileFullPath"></param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static bool WriteToFile(object obj, string FileFullPath)
        {
            // REM Create path first
            if (!Directory.Exists(Microsoft.VisualBasic.FileIO.FileSystem.GetParentPath(FileFullPath)))
                Directory.CreateDirectory(Microsoft.VisualBasic.FileIO.FileSystem.GetParentPath(FileFullPath));
            var fs = new FileStream(FileFullPath, FileMode.Create);
            var objSerializer = new BinaryFormatter();
            objSerializer.Serialize(fs, obj);
            fs.Close();
            return true;
        }

        /// <summary>
    /// Reads a serialized object from file. Throws all exceptions
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="FileFullPath"></param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static T ReadFromFile<T>(string FileFullPath)
        {
            var fs = new FileStream(FileFullPath, FileMode.Open, FileAccess.Read);
            var objSerializer = new BinaryFormatter();
            T rst = (T)objSerializer.Deserialize(fs);
            fs.Close();
            return rst;
        }

        #endregion


        #region XML


        // XML Serialization is good for shallow serialization.
        // It does NOT require Serializable attribute or implementation
        // It ONLY requires a default constructor (Parameterless)
        // 
        // It serialize only property declared public read and write
        // 
        // It is portable because it only depends on the class that calls it



        /// <summary>
    /// Writes Serializable Object to File. Throws all exceptions
    /// </summary>
    /// <param name="obj"></param>
    /// <param name="FileFullPath"></param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static bool WriteToFileXML<T>(T obj, string FileFullPath)
        {
            // REM Create path first
            if (!Directory.Exists(Microsoft.VisualBasic.FileIO.FileSystem.GetParentPath(FileFullPath)))
                Directory.CreateDirectory(Microsoft.VisualBasic.FileIO.FileSystem.GetParentPath(FileFullPath));
            var fs = new FileStream(FileFullPath, FileMode.Create);
            var objSerializer = new System.Xml.Serialization.XmlSerializer(obj.GetType());
            objSerializer.Serialize(fs, obj);
            fs.Close();
            return true;
        }

        /// <summary>
    /// Reads a serialized object from file. Throws all exceptions
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="FileFullPath"></param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static T ReadFromFileXML<T>(string FileFullPath)
        {
            var fs = new FileStream(FileFullPath, FileMode.Open, FileAccess.Read);
            var objSerializer = new System.Xml.Serialization.XmlSerializer(typeof(T));
            var reader = System.Xml.XmlReader.Create(fs);
            T rst = (T)objSerializer.Deserialize(reader);
            fs.Close();
            return rst;
        }




        #endregion




        #endregion





        /// <summary>
    /// Safely copies a file from one destination to another by making sure destination folder exists. Throws other exception
    /// </summary>
    /// <param name="src"></param>
    /// <param name="dst"></param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static bool CopyFile(string src, string dst, bool overwrite = true)
        {
            if (!Directory.Exists(getDirectoryFullPath(dst)))
                Directory.CreateDirectory(getDirectoryFullPath(dst));
            File.Copy(src, dst, overwrite);
            return true;
        }

        /// <summary>
    /// Safely copies a file from one destination to another by making sure destination folder exists. Throws other exception
    /// </summary>
    /// <param name="src"></param>
    /// <param name="dstDir"></param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static bool CopyFileToDir(string src, string dstDir, bool overwrite = true)
        {
            if (!Directory.Exists(dstDir))
                Directory.CreateDirectory(dstDir);
            return CopyFile(src, string.Format(@"{0}\{1}", dstDir, getFileName(src)), overwrite);
        }
    }
}