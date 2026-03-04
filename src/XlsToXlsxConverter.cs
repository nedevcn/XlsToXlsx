using System.IO;
using System.Threading.Tasks;
using Nedev.XlsToXlsx.Exceptions;

namespace Nedev.XlsToXlsx
{
    /// <summary>
    /// 进度更新委托，用于在转换过程中提供实时进度信息
    /// </summary>
    /// <param name="percentage">进度百分比 (0-100)</param>
    /// <param name="message">当前进度消息，描述正在执行的操作</param>
    public delegate void ProgressUpdateHandler(int percentage, string message);

    /// <summary>
    /// XLS到XLSX的转换器类
    /// 提供同步和异步方法将旧格式的Excel文件(.xls)转换为新格式(.xlsx)
    /// 支持数据验证、条件格式、图表、样式、图片和VBA等高级功能
    /// </summary>
    public class XlsToXlsxConverter
    {
        /// <summary>
        /// VBA项目大小限制（字节），默认50MB
        /// </summary>
        public static long VbaSizeLimit { get; set; } = 50 * 1024 * 1024;

        /// <summary>
        /// 将XLS文件转换为XLSX文件
        /// </summary>
        /// <param name="inputFilePath">输入XLS文件路径，必须是有效的.xls文件</param>
        /// <param name="outputFilePath">输出XLSX文件路径，将创建或覆盖该文件</param>
        /// <param name="progressUpdate">进度更新回调，可选参数，用于接收转换过程的进度信息</param>
        /// <param name="vbaSizeLimit">VBA项目大小限制（字节），默认使用全局设置</param>
        /// <exception cref="XlsToXlsxException">当输入或输出路径为空时抛出</exception>
        /// <exception cref="FileNotFoundException">当输入文件不存在时抛出</exception>
        /// <exception cref="FileFormatException">当输入文件不是XLS格式或输出文件不是XLSX格式时抛出</exception>
        /// <example>
        /// 示例用法：
        /// <code>
        /// // 基本转换
        /// XlsToXlsxConverter.Convert("input.xls", "output.xlsx");
        /// 
        /// // 带进度回调的转换
        /// XlsToXlsxConverter.Convert("input.xls", "output.xlsx", (percentage, message) =>
        /// {
        ///     Console.WriteLine($"转换进度: {percentage}% - {message}");
        /// });
        /// 
        /// // 带VBA大小限制的转换
        /// XlsToXlsxConverter.Convert("input.xls", "output.xlsx", vbaSizeLimit: 100 * 1024 * 1024); // 100MB
        /// </code>
        /// </example>
        public static void Convert(string inputFilePath, string outputFilePath, ProgressUpdateHandler? progressUpdate = null, long? vbaSizeLimit = null)
        {
            // 验证文件路径
            if (string.IsNullOrEmpty(inputFilePath))
            {
                throw new XlsToXlsxException("Input file path cannot be null or empty", 3001, "NullError");
            }
            if (string.IsNullOrEmpty(outputFilePath))
            {
                throw new XlsToXlsxException("Output file path cannot be null or empty", 3002, "NullError");
            }

            // 验证输入文件是否存在
            if (!File.Exists(inputFilePath))
            {
                throw new FileNotFoundException("Input file not found", inputFilePath);
            }

            // 验证输入文件扩展名
            string inputExtension = Path.GetExtension(inputFilePath).ToLower();
            if (inputExtension != ".xls")
            {
                throw new FileFormatException("Input file must be an XLS file");
            }

            // 验证输出文件扩展名
            string outputExtension = Path.GetExtension(outputFilePath).ToLower();
            if (outputExtension != ".xlsx")
            {
                throw new FileFormatException("Output file must be an XLSX file");
            }

            // 确保输出目录存在
            string outputDirectory = Path.GetDirectoryName(outputFilePath);
            if (!string.IsNullOrEmpty(outputDirectory) && !Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }

            try
            {
                using (var inputStream = new FileStream(inputFilePath, FileMode.Open, FileAccess.Read))
                using (var outputStream = new FileStream(outputFilePath, FileMode.Create, FileAccess.Write))
                {
                    Convert(inputStream, outputStream, progressUpdate, vbaSizeLimit);
                }
            }
            catch (XlsToXlsxException)
            {
                throw;
            }
            catch (Exception ex)
            {
                Logger.Error($"转换文件 {inputFilePath} 到 {outputFilePath} 时发生错误", ex);
                throw new XlsToXlsxException($"Error converting file {inputFilePath} to {outputFilePath}: {ex.Message}", 3003, "ConversionError", ex);
            }
        }

        /// <summary>
        /// 异步将XLS文件转换为XLSX文件
        /// </summary>
        /// <param name="inputFilePath">输入XLS文件路径，必须是有效的.xls文件</param>
        /// <param name="outputFilePath">输出XLSX文件路径，将创建或覆盖该文件</param>
        /// <param name="progressUpdate">进度更新回调，可选参数，用于接收转换过程的进度信息</param>
        /// <param name="vbaSizeLimit">VBA项目大小限制（字节），默认使用全局设置</param>
        /// <returns>异步任务，代表转换操作的完成</returns>
        /// <exception cref="ArgumentNullException">当输入或输出路径为空时抛出</exception>
        /// <exception cref="FileNotFoundException">当输入文件不存在时抛出</exception>
        /// <exception cref="FileFormatException">当输入文件不是XLS格式或输出文件不是XLSX格式时抛出</exception>
        /// <example>
        /// 示例用法：
        /// <code>
        /// // 异步转换
        /// await XlsToXlsxConverter.ConvertAsync("input.xls", "output.xlsx");
        /// 
        /// // 带进度回调的异步转换
        /// await XlsToXlsxConverter.ConvertAsync("input.xls", "output.xlsx", (percentage, message) =>
        /// {
        ///     Console.WriteLine($"转换进度: {percentage}% - {message}");
        /// });
        /// 
        /// // 带VBA大小限制的异步转换
        /// await XlsToXlsxConverter.ConvertAsync("input.xls", "output.xlsx", vbaSizeLimit: 100 * 1024 * 1024); // 100MB
        /// </code>
        /// </example>
        public static async Task ConvertAsync(string inputFilePath, string outputFilePath, ProgressUpdateHandler? progressUpdate = null, long? vbaSizeLimit = null)
        {
            // 验证文件路径
            if (string.IsNullOrEmpty(inputFilePath))
            {
                throw new XlsToXlsxException("Input file path cannot be null or empty", 3001, "NullError");
            }
            if (string.IsNullOrEmpty(outputFilePath))
            {
                throw new XlsToXlsxException("Output file path cannot be null or empty", 3002, "NullError");
            }

            // 验证输入文件是否存在
            if (!File.Exists(inputFilePath))
            {
                throw new FileNotFoundException("Input file not found", inputFilePath);
            }

            // 验证输入文件扩展名
            string inputExtension = Path.GetExtension(inputFilePath).ToLower();
            if (inputExtension != ".xls")
            {
                throw new FileFormatException("Input file must be an XLS file");
            }

            // 验证输出文件扩展名
            string outputExtension = Path.GetExtension(outputFilePath).ToLower();
            if (outputExtension != ".xlsx")
            {
                throw new FileFormatException("Output file must be an XLSX file");
            }

            // 确保输出目录存在
            string outputDirectory = Path.GetDirectoryName(outputFilePath);
            if (!string.IsNullOrEmpty(outputDirectory) && !Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }

            try
            {
                using (var inputStream = new FileStream(inputFilePath, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, FileOptions.Asynchronous))
                using (var outputStream = new FileStream(outputFilePath, FileMode.Create, FileAccess.Write, FileShare.None, 4096, FileOptions.Asynchronous))
                {
                    await ConvertAsync(inputStream, outputStream, progressUpdate, vbaSizeLimit);
                }
            }
            catch (XlsToXlsxException)
            {
                throw;
            }
            catch (Exception ex)
            {
                Logger.Error($"异步转换文件 {inputFilePath} 到 {outputFilePath} 时发生错误", ex);
                throw new XlsToXlsxException($"Error converting file {inputFilePath} to {outputFilePath}: {ex.Message}", 3003, "ConversionError", ex);
            }
        }

        /// <summary>
        /// 将XLS流转换为XLSX流
        /// </summary>
        /// <param name="inputStream">输入XLS流</param>
        /// <param name="outputStream">输出XLSX流</param>
        /// <param name="progressUpdate">进度更新回调</param>
        /// <param name="vbaSizeLimit">VBA项目大小限制（字节），默认使用全局设置</param>
        public static void Convert(Stream inputStream, Stream outputStream, ProgressUpdateHandler? progressUpdate = null, long? vbaSizeLimit = null)
        {
            progressUpdate?.Invoke(0, "开始解析XLS文件");
            
            // 解析XLS文件
            var xlsParser = new Formats.Xls.XlsParser(inputStream);
            xlsParser.VbaSizeLimit = vbaSizeLimit ?? VbaSizeLimit;
            var workbook = xlsParser.Parse();
            
            progressUpdate?.Invoke(50, "XLS文件解析完成，开始生成XLSX文件");

            // 生成XLSX文件
            var xlsxGenerator = new Formats.Xlsx.XlsxGenerator(outputStream);
            xlsxGenerator.VbaSizeLimit = vbaSizeLimit ?? VbaSizeLimit;
            xlsxGenerator.Generate(workbook);
            
            progressUpdate?.Invoke(100, "XLSX文件生成完成");
        }

        /// <summary>
        /// 异步将XLS流转换为XLSX流
        /// </summary>
        /// <param name="inputStream">输入XLS流</param>
        /// <param name="outputStream">输出XLSX流</param>
        /// <param name="progressUpdate">进度更新回调</param>
        /// <param name="vbaSizeLimit">VBA项目大小限制（字节），默认使用全局设置</param>
        /// <returns>异步任务</returns>
        public static async Task ConvertAsync(Stream inputStream, Stream outputStream, ProgressUpdateHandler? progressUpdate = null, long? vbaSizeLimit = null)
        {
            progressUpdate?.Invoke(0, "开始解析XLS文件");
            
            // 解析XLS文件
            var xlsParser = new Formats.Xls.XlsParser(inputStream);
            xlsParser.VbaSizeLimit = vbaSizeLimit ?? VbaSizeLimit;
            var workbook = await xlsParser.ParseAsync();
            
            progressUpdate?.Invoke(50, "XLS文件解析完成，开始生成XLSX文件");

            // 生成XLSX文件
            var xlsxGenerator = new Formats.Xlsx.XlsxGenerator(outputStream);
            xlsxGenerator.VbaSizeLimit = vbaSizeLimit ?? VbaSizeLimit;
            await xlsxGenerator.GenerateAsync(workbook);
            
            progressUpdate?.Invoke(100, "XLSX文件生成完成");
        }

        /// <summary>
        /// 批量转换XLS文件到XLSX文件
        /// </summary>
        /// <param name="inputFilePaths">输入XLS文件路径列表，每个路径必须是有效的.xls文件</param>
        /// <param name="outputFilePaths">输出XLSX文件路径列表，长度必须与输入列表相同</param>
        /// <param name="progressUpdate">进度更新回调，可选参数，用于接收批量转换过程的进度信息</param>
        /// <param name="vbaSizeLimit">VBA项目大小限制（字节），默认使用全局设置</param>
        /// <exception cref="ArgumentNullException">当输入或输出文件路径列表为空时抛出</exception>
        /// <exception cref="ArgumentException">当输入和输出文件路径列表长度不匹配时抛出</exception>
        /// <example>
        /// 示例用法：
        /// <code>
        /// // 批量转换
        /// string[] inputFiles = { "file1.xls", "file2.xls" };
        /// string[] outputFiles = { "file1.xlsx", "file2.xlsx" };
        /// XlsToXlsxConverter.BatchConvert(inputFiles, outputFiles);
        /// 
        /// // 带进度回调的批量转换
        /// XlsToXlsxConverter.BatchConvert(inputFiles, outputFiles, (percentage, message) =>
        /// {
        ///     Console.WriteLine($"批量转换进度: {percentage}% - {message}");
        /// });
        /// 
        /// // 带VBA大小限制的批量转换
        /// XlsToXlsxConverter.BatchConvert(inputFiles, outputFiles, vbaSizeLimit: 100 * 1024 * 1024); // 100MB
        /// </code>
        /// </example>
        public static void BatchConvert(string[] inputFilePaths, string[] outputFilePaths, ProgressUpdateHandler? progressUpdate = null, long? vbaSizeLimit = null)
        {
            // 验证参数
            if (inputFilePaths == null)
            {
                throw new ArgumentNullException(nameof(inputFilePaths), "Input file paths cannot be null");
            }
            if (outputFilePaths == null)
            {
                throw new ArgumentNullException(nameof(outputFilePaths), "Output file paths cannot be null");
            }
            if (inputFilePaths.Length == 0)
            {
                // 允许空文件列表，直接返回
                return;
            }
            if (inputFilePaths.Length != outputFilePaths.Length)
            {
                throw new System.ArgumentException("输入文件路径和输出文件路径数量必须相同");
            }

            var errors = new List<Exception>();

            for (int i = 0; i < inputFilePaths.Length; i++)
            {
                int fileIndex = i;
                int totalFiles = inputFilePaths.Length;
                
                int fileProgressWeight = 100 / totalFiles;
                int startPercentage = fileIndex * fileProgressWeight;
                progressUpdate?.Invoke(startPercentage, $"开始处理文件 {fileIndex + 1}/{totalFiles}: {Path.GetFileName(inputFilePaths[fileIndex])}");
                
                try
                {
                    Convert(inputFilePaths[fileIndex], outputFilePaths[fileIndex], (percentage, message) =>
                    {
                        int overallPercentage = startPercentage + (percentage * fileProgressWeight) / 100;
                        progressUpdate?.Invoke(overallPercentage, $"处理文件 {fileIndex + 1}/{totalFiles}: {message}");
                    }, vbaSizeLimit);
                }
                catch (Exception ex)
                {
                    Logger.Error($"批量转换中处理文件 {inputFilePaths[fileIndex]} 时发生错误", ex);
                    errors.Add(new Exception($"处理文件 {inputFilePaths[fileIndex]} 时发生错误: {ex.Message}", ex));
                }
            }

            progressUpdate?.Invoke(100, "批量转换完成");

            // 如果有错误，抛出汇总异常
            if (errors.Count > 0)
            {
                throw new AggregateException("批量转换过程中发生错误", errors);
            }
        }

        /// <summary>
        /// 异步批量转换XLS文件到XLSX文件
        /// </summary>
        /// <param name="inputFilePaths">输入XLS文件路径列表，每个路径必须是有效的.xls文件</param>
        /// <param name="outputFilePaths">输出XLSX文件路径列表，长度必须与输入列表相同</param>
        /// <param name="progressUpdate">进度更新回调，可选参数，用于接收批量转换过程的进度信息</param>
        /// <returns>异步任务，代表批量转换操作的完成</returns>
        /// <exception cref="ArgumentNullException">当输入或输出文件路径列表为空时抛出</exception>
        /// <exception cref="ArgumentException">当输入和输出文件路径列表长度不匹配时抛出</exception>
        /// <example>
        /// 示例用法：
        /// <code>
        /// // 异步批量转换
        /// string[] inputFiles = { "file1.xls", "file2.xls" };
        /// string[] outputFiles = { "file1.xlsx", "file2.xlsx" };
        /// await XlsToXlsxConverter.BatchConvertAsync(inputFiles, outputFiles);
        /// 
        /// // 带进度回调的异步批量转换
        /// await XlsToXlsxConverter.BatchConvertAsync(inputFiles, outputFiles, (percentage, message) =>
        /// {
        ///     Console.WriteLine($"批量转换进度: {percentage}% - {message}");
        /// });
        /// </code>
        /// </example>
        public static async Task BatchConvertAsync(string[] inputFilePaths, string[] outputFilePaths, ProgressUpdateHandler? progressUpdate = null)
        {
            // 验证参数
            if (inputFilePaths == null)
            {
                throw new XlsToXlsxException("Input file paths cannot be null", 3004, "NullError");
            }
            if (outputFilePaths == null)
            {
                throw new XlsToXlsxException("Output file paths cannot be null", 3005, "NullError");
            }
            if (inputFilePaths.Length == 0)
            {
                // 允许空文件列表，返回完成的任务
                await Task.CompletedTask;
                return;
            }
            if (inputFilePaths.Length != outputFilePaths.Length)
            {
                throw new XlsToXlsxException("输入文件路径和输出文件路径数量必须相同", 3006, "ArgumentException");
            }

            int totalFiles = inputFilePaths.Length;
            var tasks = new List<Task>();
            var progressCounts = new int[totalFiles];
            var progressLock = new object();
            var errors = new List<Exception>();
            var errorsLock = new object();

            for (int i = 0; i < inputFilePaths.Length; i++)
            {
                int fileIndex = i;
                int fileProgressWeight = 100 / totalFiles;
                int startPercentage = fileIndex * fileProgressWeight;
                
                tasks.Add(Task.Run(async () =>
                {
                    try
                    {
                        progressUpdate?.Invoke(startPercentage, $"开始处理文件 {fileIndex + 1}/{totalFiles}: {Path.GetFileName(inputFilePaths[fileIndex])}");
                        
                        await ConvertAsync(inputFilePaths[fileIndex], outputFilePaths[fileIndex], (percentage, message) =>
                        {
                            lock (progressLock)
                            {
                                progressCounts[fileIndex] = percentage;
                                int overallPercentage = startPercentage + (percentage * fileProgressWeight) / 100;
                                progressUpdate?.Invoke(overallPercentage, $"处理文件 {fileIndex + 1}/{totalFiles}: {message}");
                            }
                        });
                    }
                    catch (Exception ex)
                    {
                        Logger.Error($"异步批量转换中处理文件 {inputFilePaths[fileIndex]} 时发生错误", ex);
                        lock (errorsLock)
                        {
                            errors.Add(new Exception($"处理文件 {inputFilePaths[fileIndex]} 时发生错误: {ex.Message}", ex));
                        }
                    }
                }));
            }

            await Task.WhenAll(tasks);
            progressUpdate?.Invoke(100, "批量转换完成");

            // 如果有错误，抛出汇总异常
            if (errors.Count > 0)
            {
                throw new AggregateException("批量转换过程中发生错误", errors);
            }
        }

        /// <summary>
        /// 计算整体进度
        /// </summary>
        /// <param name="progressCounts">每个文件的进度</param>
        /// <param name="totalFiles">总文件数</param>
        /// <returns>整体进度百分比</returns>
        private static int CalculateOverallProgress(int[] progressCounts, int totalFiles)
        {
            if (progressCounts == null || progressCounts.Length == 0)
                return 0;

            int sum = 0;
            foreach (int progress in progressCounts)
            {
                sum += progress;
            }

            return sum / totalFiles;
        }
    }
}