using System;

namespace Nedev.FileConverters.XlsToXlsx.Exceptions
{
    /// <summary>
    /// 处理样式相关的异常类
    /// </summary>
    public class StyleProcessingException : XlsToXlsxException
    {
        public StyleProcessingException(string message) 
            : base(message, 3004, "StyleProcessingError")
        {
        }

        public StyleProcessingException(string message, Exception innerException) 
            : base(message, 3004, "StyleProcessingError", innerException)
        {
        }
    }
}
