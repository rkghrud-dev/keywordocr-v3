using System;
using System.IO;
using System.Threading;
using ClosedXML.Excel;

namespace KeywordOcr.App.Services;

internal static class WorkbookFileLoader
{
    public static XLWorkbook OpenReadOnly(string path, int maxAttempts = 5, int delayMilliseconds = 150)
    {
        Exception? lastException = null;

        for (var attempt = 1; attempt <= maxAttempts; attempt++)
        {
            try
            {
                using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
                var buffer = new MemoryStream();
                stream.CopyTo(buffer);
                buffer.Position = 0;
                return new XLWorkbook(buffer);
            }
            catch (IOException ex)
            {
                lastException = ex;
                if (attempt == maxAttempts)
                {
                    break;
                }

                Thread.Sleep(delayMilliseconds);
            }
        }

        throw new IOException($"Excel 파일을 읽는 중 잠금이 해제되지 않았습니다: {path}", lastException);
    }
}
