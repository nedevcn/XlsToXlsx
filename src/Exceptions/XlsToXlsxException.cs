using System;
using System.Runtime.Serialization;

namespace Nedev.XlsToXlsx.Exceptions
{
    /// <summary>
    /// XLS到XLSX转换异常类
    /// </summary>
    [Serializable]
    public class XlsToXlsxException : Exception
    {
        /// <summary>
        /// 错误代码
        /// </summary>
        public int ErrorCode { get; set; }
        
        /// <summary>
        /// 错误类型
        /// </summary>
        public string ErrorType { get; set; }
        
        /// <summary>
        /// 构造函数
        /// </summary>
        public XlsToXlsxException() : base() { }
        
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="message">错误消息</param>
        public XlsToXlsxException(string message) : base(message) { }
        
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="message">错误消息</param>
        /// <param name="innerException">内部异常</param>
        public XlsToXlsxException(string message, Exception innerException) : base(message, innerException) { }
        
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="message">错误消息</param>
        /// <param name="errorCode">错误代码</param>
        /// <param name="errorType">错误类型</param>
        public XlsToXlsxException(string message, int errorCode, string errorType) : base(message)
        {
            ErrorCode = errorCode;
            ErrorType = errorType;
        }
        
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="message">错误消息</param>
        /// <param name="errorCode">错误代码</param>
        /// <param name="errorType">错误类型</param>
        /// <param name="innerException">内部异常</param>
        public XlsToXlsxException(string message, int errorCode, string errorType, Exception innerException) : base(message, innerException)
        {
            ErrorCode = errorCode;
            ErrorType = errorType;
        }
        
        /// <summary>
        /// 序列化构造函数
        /// </summary>
        /// <param name="info">序列化信息</param>
        /// <param name="context">序列化上下文</param>
        protected XlsToXlsxException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
            if (info != null)
            {
                ErrorCode = info.GetInt32("ErrorCode");
                ErrorType = info.GetString("ErrorType");
            }
        }
        
        /// <summary>
        /// 序列化方法
        /// </summary>
        /// <param name="info">序列化信息</param>
        /// <param name="context">序列化上下文</param>
        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            base.GetObjectData(info, context);
            if (info != null)
            {
                info.AddValue("ErrorCode", ErrorCode);
                info.AddValue("ErrorType", ErrorType);
            }
        }
    }
    
    /// <summary>
    /// XLS解析异常类
    /// </summary>
    [Serializable]
    public class XlsParseException : XlsToXlsxException
    {
        /// <summary>
        /// 构造函数
        /// </summary>
        public XlsParseException() : base() { }
        
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="message">错误消息</param>
        public XlsParseException(string message) : base(message, 1001, "XlsParseError") { }
        
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="message">错误消息</param>
        /// <param name="innerException">内部异常</param>
        public XlsParseException(string message, Exception innerException) : base(message, 1001, "XlsParseError", innerException) { }
        
        /// <summary>
        /// 序列化构造函数
        /// </summary>
        /// <param name="info">序列化信息</param>
        /// <param name="context">序列化上下文</param>
        protected XlsParseException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
    
    /// <summary>
    /// XLSX生成异常类
    /// </summary>
    [Serializable]
    public class XlsxGenerateException : XlsToXlsxException
    {
        /// <summary>
        /// 构造函数
        /// </summary>
        public XlsxGenerateException() : base() { }
        
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="message">错误消息</param>
        public XlsxGenerateException(string message) : base(message, 2001, "XlsxGenerateError") { }
        
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="message">错误消息</param>
        /// <param name="innerException">内部异常</param>
        public XlsxGenerateException(string message, Exception innerException) : base(message, 2001, "XlsxGenerateError", innerException) { }
        
        /// <summary>
        /// 序列化构造函数
        /// </summary>
        /// <param name="info">序列化信息</param>
        /// <param name="context">序列化上下文</param>
        protected XlsxGenerateException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
    
    /// <summary>
    /// 文件格式异常类
    /// </summary>
    [Serializable]
    public class FileFormatException : XlsToXlsxException
    {
        /// <summary>
        /// 构造函数
        /// </summary>
        public FileFormatException() : base() { }
        
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="message">错误消息</param>
        public FileFormatException(string message) : base(message, 3001, "FileFormatError") { }
        
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="message">错误消息</param>
        /// <param name="innerException">内部异常</param>
        public FileFormatException(string message, Exception innerException) : base(message, 3001, "FileFormatError", innerException) { }
        
        /// <summary>
        /// 序列化构造函数
        /// </summary>
        /// <param name="info">序列化信息</param>
        /// <param name="context">序列化上下文</param>
        protected FileFormatException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
    
    /// <summary>
    /// 图片处理异常类
    /// </summary>
    [Serializable]
    public class ImageProcessingException : XlsToXlsxException
    {
        /// <summary>
        /// 构造函数
        /// </summary>
        public ImageProcessingException() : base() { }
        
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="message">错误消息</param>
        public ImageProcessingException(string message) : base(message, 4001, "ImageProcessingError") { }
        
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="message">错误消息</param>
        /// <param name="innerException">内部异常</param>
        public ImageProcessingException(string message, Exception innerException) : base(message, 4001, "ImageProcessingError", innerException) { }
        
        /// <summary>
        /// 序列化构造函数
        /// </summary>
        /// <param name="info">序列化信息</param>
        /// <param name="context">序列化上下文</param>
        protected ImageProcessingException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
    
    /// <summary>
    /// 图表处理异常类
    /// </summary>
    [Serializable]
    public class ChartProcessingException : XlsToXlsxException
    {
        /// <summary>
        /// 构造函数
        /// </summary>
        public ChartProcessingException() : base() { }
        
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="message">错误消息</param>
        public ChartProcessingException(string message) : base(message, 5001, "ChartProcessingError") { }
        
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="message">错误消息</param>
        /// <param name="innerException">内部异常</param>
        public ChartProcessingException(string message, Exception innerException) : base(message, 5001, "ChartProcessingError", innerException) { }
        
        /// <summary>
        /// 序列化构造函数
        /// </summary>
        /// <param name="info">序列化信息</param>
        /// <param name="context">序列化上下文</param>
        protected ChartProcessingException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
    
    /// <summary>
    /// 内存不足异常类
    /// </summary>
    [Serializable]
    public class OutOfMemoryException : XlsToXlsxException
    {
        /// <summary>
        /// 构造函数
        /// </summary>
        public OutOfMemoryException() : base() { }
        
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="message">错误消息</param>
        public OutOfMemoryException(string message) : base(message, 6001, "OutOfMemoryError") { }
        
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="message">错误消息</param>
        /// <param name="innerException">内部异常</param>
        public OutOfMemoryException(string message, Exception innerException) : base(message, 6001, "OutOfMemoryError", innerException) { }
        
        /// <summary>
        /// 序列化构造函数
        /// </summary>
        /// <param name="info">序列化信息</param>
        /// <param name="context">序列化上下文</param>
        protected OutOfMemoryException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
    
    /// <summary>
    /// 权限异常类
    /// </summary>
    [Serializable]
    public class PermissionException : XlsToXlsxException
    {
        /// <summary>
        /// 构造函数
        /// </summary>
        public PermissionException() : base() { }
        
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="message">错误消息</param>
        public PermissionException(string message) : base(message, 7001, "PermissionError") { }
        
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="message">错误消息</param>
        /// <param name="innerException">内部异常</param>
        public PermissionException(string message, Exception innerException) : base(message, 7001, "PermissionError", innerException) { }
        
        /// <summary>
        /// 序列化构造函数
        /// </summary>
        /// <param name="info">序列化信息</param>
        /// <param name="context">序列化上下文</param>
        protected PermissionException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}
