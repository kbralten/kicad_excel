using System;
using System.IO;
using System.Threading;

namespace KiCadExcelBridge
{
    internal static class FileHelpers
    {
        // Try to open the file for shared read access with retries. If that fails,
        // attempt to copy to a temp file and open the temp copy.
        // Returns null if unable to open.
        public static Stream? OpenFileForReadWithFallback(string filePath, int maxRetries, out string actualPath, out bool usedFallback)
        {
            actualPath = filePath;
            usedFallback = false;

            if (!File.Exists(filePath)) return null;

            // Try to open directly with retries
            for (int attempt = 0; attempt < maxRetries; attempt++)
            {
                try
                {
                    var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    return fs;
                }
                catch (IOException)
                {
                    // wait a bit and retry
                    Thread.Sleep(150 * (attempt + 1));
                }
                catch (UnauthorizedAccessException)
                {
                    Thread.Sleep(150 * (attempt + 1));
                }
            }

            // If we couldn't open directly, try to copy to temp and open the copy
            try
            {
                var tempPath = Path.Combine(Path.GetTempPath(), "kicad_excel_" + Guid.NewGuid().ToString() + Path.GetExtension(filePath));
                File.Copy(filePath, tempPath, overwrite: true);
                // Try to open the temp copy
                var tempFs = new FileStream(tempPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                actualPath = tempPath;
                usedFallback = true;
                return tempFs;
            }
            catch
            {
                actualPath = string.Empty;
                usedFallback = false;
                return null;
            }
        }
    }
}
