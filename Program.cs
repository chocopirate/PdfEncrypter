using System;
using System.IO;
using iText.Kernel.Pdf;
using System.Runtime.InteropServices;

namespace PdfEncrypter
{

    [Guid("471264F8-3DF4-4B1C-9811-B1C918CE2CC4")]
    [ComVisible(true)]
    public interface IProgram
    {
        void SetSourceFilePath(String path);
        void SetEncryptedFilePath(String path);
        void SetUserPassword(String password);
        int EncryptFile();
    }

    [Guid("366B47E2-FEC9-4E0D-BD33-6A812DEC6B3F")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("PdfEncrypter.Program")]
    [ComVisible(true)]
    public class Program : IProgram
    {
        public String sourceFilePath = "";
        public String saveFilePath = "";
        public byte[] userPassword = null;
        public byte[] masterPassword = System.Text.Encoding.Default.GetBytes("master");

        public void SetSourceFilePath(String path)
        {
            sourceFilePath = path.Replace(@"\", @"/");
        }

        public void SetEncryptedFilePath(String path)
        {
            saveFilePath = path.Replace(@"\", @"/");
        }

        public void SetUserPassword(String password)
        {
            userPassword = System.Text.Encoding.Default.GetBytes(password);
        }

        public int EncryptFile()
        {
            if (saveFilePath == sourceFilePath)
            {
                return -2;
            }
            try
            {
                PdfReader pdfReader = new PdfReader(sourceFilePath);
                WriterProperties writerProperties = new WriterProperties();
                writerProperties.SetStandardEncryption(userPassword, masterPassword, EncryptionConstants.ALLOW_PRINTING, EncryptionConstants.ENCRYPTION_AES_128);
                PdfWriter pdfWriter = new PdfWriter(new FileStream(saveFilePath, FileMode.Create), writerProperties);
                PdfDocument pdfDocument = new PdfDocument(pdfReader, pdfWriter);
                pdfDocument.Close();
                return 0;
            }
            catch (Exception)
            {
                return -1;
            }
        }
    }
}
