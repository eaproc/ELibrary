using System;
using System.Drawing;
using static ELibrary.Standard.EIO;

namespace ELibrary.Standard.Objects
{
    public class EImage
    {
        public static Image valueOf(byte[] BytesArray)
        {
            try
            {
                if (BytesArray is null)
                    return null;
                using (var ms = new System.IO.MemoryStream(BytesArray, 0, BytesArray.Length))
                {
                    ms.Write(BytesArray, 0, BytesArray.Length);

                    // New Bitmap is to avoid "A generic error occurred in GDI+" when saving back
                    return new Bitmap(Image.FromStream(ms, true));
                }
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public static Image valueOf(DBNull SysDBNULL)
        {
            return null;
        }

        public static Image valueOf(object pVal)
        {
            if (pVal is null)
                return null;
            if (pVal is byte[])
                return valueOf((byte[])pVal);
            if (pVal is DBNull)
                return valueOf((DBNull)pVal);
            return null; // REM We don't understand the format
        }







        /// <summary>
        /// To solve the issue with image saving not being able to decern it's format
        /// </summary>
        /// <param name="Img"></param>
        /// <param name="FilePath">The real output is Bitmaps(bmp)</param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool SaveImage_To_File(Image Img, string FilePath)
        {
            try
            {
                var newBmp = new Bitmap(Img);
                newBmp.Save(FilePath, System.Drawing.Imaging.ImageFormat.Bmp);
                return true;
            }
            catch (Exception ex)
            {
            }

            return false;
        }










        /// <summary>
        /// Get Icon from application type
        /// </summary>
        /// <param name="objType"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static Icon get_Icon_From_Object_Type(Type objType)
        {
            try
            {
                return Icon.ExtractAssociatedIcon(get_Code_File_StartUp_File_FullPathName(objType.GetType()));
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public static byte[] toByteArray(Image pImage, System.Drawing.Imaging.ImageFormat pImageFormat)
        {
            try
            {
                byte[] data = null;
                var ms = new System.IO.MemoryStream();
                bool pMemoryFormatWorked = false;
                try
                {
                    pImage.Save(ms, pImageFormat);
                    pMemoryFormatWorked = true;
                }
                catch (Exception ex)
                {
                    Modules.basMain.MyLogFile.Print("toByteArray(ByVal pImage As Image, pImageFormat As Imaging.ImageFormat)", ex);
                }

                if (!pMemoryFormatWorked)
                {
                    // Try Bitmap
                    pImage.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
                }

                data = ms.GetBuffer();
                ms = null;
                return data;
            }
            catch (Exception ex)
            {
                Modules.basMain.MyLogFile.Print("toByteArray(ByVal pImage As Image, pImageFormat As Imaging.ImageFormat)", ex);
                return null;
            }
        }

        /// <summary>
        /// It doesnt throw exception
        /// </summary>
        /// <param name="pImage"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static byte[] toByteArray(Image pImage)
        {
            try
            {
                if (pImage is null)
                    return null;
                return toByteArray(pImage, pImage.RawFormat);
            }
            catch (Exception ex)
            {
                Modules.basMain.MyLogFile.Print("toByteArray(ByVal pImage As Image)", ex);
                Modules.basMain.MyLogFile.Print(ex);
                return null;
            }
        }







        #region Bas64 Conversions
        /// <summary>
        /// It doesnt throw exception
        /// </summary>
        /// <param name="pBytesFile"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string toBase64String(byte[] pBytesFile)
        {
            try
            {
                return Convert.ToBase64String(pBytesFile);
            }
            catch (Exception ex)
            {
                return string.Empty;
            }
        }
        /// <summary>
        /// It doesnt throw exception
        /// </summary>
        /// <param name="pImage"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string toBase64String(Image pImage)
        {
            return toBase64String(toByteArray(pImage));
        }


        /// <summary>
        /// It doesnt throw exception
        /// </summary>
        /// <param name="pBase64Data">Dont add the decoding data [data:image/gif;base64,]</param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static byte[] fromBase64String(string pBase64Data)
        {
            try
            {
                var p = Convert.FromBase64String(pBase64Data);
                return p;
            }
            catch (Exception ex)
            {
                return null;
            }
        }


        #endregion



    }
}