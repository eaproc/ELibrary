using System;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Xml.Serialization;

namespace ELibrary.Standard
{
    /// <summary>
    /// Static Class for I/O functions
    /// </summary>
    public static class EIO
    {
        public const string WINDOWS_TEMP_FOLDER = @"C:\Windows\Temp";

        /// <summary>
        /// Try to empty a directory without deleting the directory itself.
        /// </summary>
        public static void DeleteAllFromDirectory(string directory)
        {
            if (!Directory.Exists(directory))
                return;

            // Delete subdirectories
            foreach (var subDir in Directory.GetDirectories(directory))
            {
                try
                {
                    Directory.Delete(subDir, true);
                }
                catch { /* Log error if necessary */ }
            }

            // Delete files
            foreach (var subFile in Directory.GetFiles(directory))
            {
                try
                {
                    File.Delete(subFile);
                }
                catch { /* Log error if necessary */ }
            }
        }

        /// <summary>
        /// Deletes a directory and all its contents if exists. Ignores errors on delete.
        /// </summary>
        public static bool DeleteDirectory(string directoryPath)
        {
            try
            {
                if (Directory.Exists(directoryPath))
                    Directory.Delete(directoryPath, true);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Deletes a directory and all its contents if it exists, and recreates it as an empty directory.
        /// </summary>
        public static bool DeleteAndRecreateDirectory(string directoryPath)
        {
            if (!DeleteDirectory(directoryPath))
                return false;

            try
            {
                Directory.CreateDirectory(directoryPath);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Deletes a file if it exists.
        /// </summary>
        public static bool DeleteFileIfExists(string fileName)
        {
            try
            {
                if (File.Exists(fileName))
                    File.Delete(fileName);
                return true;
            }
            catch
            {
                return false;
            }
        }

        #region Path Operations

        /// <summary>
        /// Gets the extension of a file, including the dot (e.g., ".txt").
        /// </summary>
        public static string GetFileExtension(string fileName)
        {
            return Path.GetExtension(fileName) ?? string.Empty;
        }

        /// <summary>
        /// Fetches the name of the last directory in the provided path.
        /// </summary>
        public static string GetDirectoryName(string directoryPath, bool includeSlash = false)
        {
            var directoryName = Path.GetFileName(Path.GetDirectoryName(directoryPath));
            return includeSlash ? $@"\{directoryName}" : directoryName;
        }

        /// <summary>
        /// Get the full directory path from a file path.
        /// </summary>
        public static string GetDirectoryFullPath(string filePath)
        {
            return Path.GetDirectoryName(filePath) ?? string.Empty;
        }

        /// <summary>
        /// Get a file name without its extension.
        /// </summary>
        public static string GetFileNameWithoutExtension(string fileName)
        {
            return Path.GetFileNameWithoutExtension(fileName);
        }

        /// <summary>
        /// Get the file name from a URI.
        /// </summary>
        public static string GetFileName(Uri uri)
        {
            return Path.GetFileName(uri.AbsolutePath) ?? string.Empty;
        }

        /// <summary>
        /// Get the file name from a URL string.
        /// </summary>
        public static string GetFileName(string url)
        {
            return Path.GetFileName(url) ?? string.Empty;
        }

        /// <summary>
        /// Get a unique file name in a directory by incrementing the name if it exists.
        /// </summary>
        public static string GetAvailableFileName(string folderPath, string suggestedFileName = "File.dat")
        {
            if (!Directory.Exists(folderPath))
                return string.Empty;

            var ext = Path.GetExtension(suggestedFileName);
            var nameWithoutExt = GetFileNameWithoutExtension(suggestedFileName);
            var fullPath = Path.Combine(folderPath, suggestedFileName);
            long count = 0;

            while (File.Exists(fullPath))
            {
                count++;
                fullPath = Path.Combine(folderPath, $"{nameWithoutExt}_{count}{ext}");
            }

            return fullPath;
        }

        /// <summary>
        /// Checks if the specified path is a file (not a directory).
        /// </summary>
        public static bool IsFile(string filePath)
        {
            try
            {
                return File.Exists(filePath);
            }
            catch
            {
                return false;
            }
        }

        #endregion

        #region Serialization

        /// <summary>
        /// Writes a serializable object to a binary file.
        /// </summary>
        public static bool WriteToFile(object obj, string filePath)
        {
            try
            {
                Directory.CreateDirectory(GetDirectoryFullPath(filePath));
                using (var fs = new FileStream(filePath, FileMode.Create))
                {
                    var formatter = new BinaryFormatter();
                    formatter.Serialize(fs, obj);
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Reads a serialized binary object from file.
        /// </summary>
        public static T ReadFromFile<T>(string filePath)
        {
            try
            {
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    var formatter = new BinaryFormatter();
                    return (T)formatter.Deserialize(fs);
                }
            }
            catch
            {
                return default(T);
            }
        }

        /// <summary>
        /// Writes a serializable object to an XML file.
        /// </summary>
        public static bool WriteToFileXML<T>(T obj, string filePath)
        {
            try
            {
                Directory.CreateDirectory(GetDirectoryFullPath(filePath));
                using (var fs = new FileStream(filePath, FileMode.Create))
                {
                    var serializer = new XmlSerializer(typeof(T));
                    serializer.Serialize(fs, obj);
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Reads a serialized XML object from file.
        /// </summary>
        public static T ReadFromFileXML<T>(string filePath)
        {
            try
            {
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    var serializer = new XmlSerializer(typeof(T));
                    return (T)serializer.Deserialize(fs);
                }
            }
            catch
            {
                return default(T);
            }
        }

        #endregion

        /// <summary>
        /// Safely copies a file from one destination to another by ensuring the destination folder exists.
        /// </summary>
        public static bool CopyFile(string src, string dst, bool overwrite = true)
        {
            try
            {
                Directory.CreateDirectory(GetDirectoryFullPath(dst));
                File.Copy(src, dst, overwrite);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Safely copies a file to a directory by ensuring the directory exists.
        /// </summary>
        public static bool CopyFileToDir(string src, string dstDir, bool overwrite = true)
        {
            try
            {
                Directory.CreateDirectory(dstDir);
                var destinationFile = Path.Combine(dstDir, GetFileName(src));
                return CopyFile(src, destinationFile, overwrite);
            }
            catch
            {
                return false;
            }
        }
    }
}
