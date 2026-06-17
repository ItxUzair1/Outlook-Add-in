using System;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace KoyoLink
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                if (args.Length == 0) return;
                string filePath = args[0];

                if (string.IsNullOrEmpty(filePath)) return;

                // Resolve path to absolute path
                string absolutePath = Path.GetFullPath(filePath);

                // Resolve path name
                string name = Path.GetFileName(absolutePath);
                if (string.IsNullOrEmpty(name))
                {
                    name = absolutePath;
                }

                // Use the backend HTTP endpoint to open the file to bypass strict New Outlook file:/// blocking.
                // This is the only universal format that works in both New Outlook and Classic Outlook.
                string encodedPath = Uri.EscapeDataString(absolutePath);
                string fileUri = "https://localhost:4000/api/search/open-local?path=" + encodedPath;
                
                // Format as a standard HTML anchor link with explicit styling for Outlook
                string htmlFragment = string.Format("<a href=\"{0}\" style=\"color:blue; text-decoration:underline;\">{1}</a>", fileUri, name);

                // Wrap in Windows Clipboard CF_HTML format
                string htmlClipboardData = GetClipboardHtmlWrapper(htmlFragment, fileUri);

                // Create DataObject with multiple formats
                DataObject data = new DataObject();
                
                // 1. HTML Format for rich text editors (Outlook, Word, Teams, etc.)
                // Use MemoryStream to avoid .NET string to UTF-8 conversion altering the exact byte offsets
                byte[] htmlBytes = Encoding.UTF8.GetBytes(htmlClipboardData);
                byte[] htmlBytesNullTerminated = new byte[htmlBytes.Length + 1];
                Array.Copy(htmlBytes, htmlBytesNullTerminated, htmlBytes.Length);
                htmlBytesNullTerminated[htmlBytes.Length] = 0; // null terminator
                
                MemoryStream htmlStream = new MemoryStream(htmlBytesNullTerminated);
                data.SetData("HTML Format", htmlStream);
                
                // 2. Plain Text Format fallback for plain text editors (Notepad, search bars, etc.)
                data.SetText(absolutePath, TextDataFormat.UnicodeText);

                // Copy to clipboard and persist after application exits (retry 5 times, 200ms apart)
                Clipboard.SetDataObject(data, true, 5, 200);
            }
            catch (Exception ex)
            {
                // Log error for debugging if it fails silently
                try { File.WriteAllText(Path.Combine(Path.GetTempPath(), "koyolink_error.log"), ex.ToString()); } catch { }
            }
        }

        static string GetClipboardHtmlWrapper(string htmlFragment, string sourceUrl)
        {
            string HeaderTemplate =
                "Version:0.9\r\n" +
                "StartHTML:{0:D10}\r\n" +
                "EndHTML:{1:D10}\r\n" +
                "StartFragment:{2:D10}\r\n" +
                "EndFragment:{3:D10}\r\n" +
                "SourceURL:" + sourceUrl + "\r\n";

            string htmlStart = "<html>\r\n<body>\r\n<!--StartFragment-->";
            string htmlEnd = "<!--EndFragment-->\r\n</body>\r\n</html>";

            // Estimate header size with placeholder offsets (10 digits each)
            string dummyHeader = string.Format(HeaderTemplate, 0, 0, 0, 0);
            int headerByteCount = Encoding.UTF8.GetByteCount(dummyHeader);

            int startHtml = headerByteCount;
            int startFragment = startHtml + Encoding.UTF8.GetByteCount(htmlStart);
            int endFragment = startFragment + Encoding.UTF8.GetByteCount(htmlFragment);
            int endHtml = endFragment + Encoding.UTF8.GetByteCount(htmlEnd);

            string finalHeader = string.Format(HeaderTemplate, startHtml, endHtml, startFragment, endFragment);
            return finalHeader + htmlStart + htmlFragment + htmlEnd;
        }
    }
}
