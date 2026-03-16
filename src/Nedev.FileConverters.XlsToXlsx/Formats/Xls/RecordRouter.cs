using System;
using System.Collections.Generic;

namespace Nedev.FileConverters.XlsToXlsx.Formats.Xls
{
    /// <summary>
    /// BIFF记录路由处理器 - 使用字典映射替代switch-case
    /// </summary>
    public class RecordRouter
    {
        private readonly Dictionary<ushort, Action<BiffRecord>> _handlers;

        public RecordRouter()
        {
            _handlers = new Dictionary<ushort, Action<BiffRecord>>();
        }

        /// <summary>
        /// 注册记录处理器
        /// </summary>
        /// <param name="recordType">记录类型ID</param>
        /// <param name="handler">处理委托</param>
        public void Register(ushort recordType, Action<BiffRecord> handler)
        {
            _handlers[recordType] = handler;
        }

        /// <summary>
        /// 批量注册记录处理器
        /// </summary>
        /// <param name="recordTypes">记录类型ID数组</param>
        /// <param name="handler">处理委托</param>
        public void RegisterRange(ushort[] recordTypes, Action<BiffRecord> handler)
        {
            foreach (var type in recordTypes)
            {
                _handlers[type] = handler;
            }
        }

        /// <summary>
        /// 路由记录到对应的处理器
        /// </summary>
        /// <param name="record">BIFF记录</param>
        /// <returns>是否找到对应的处理器</returns>
        public bool Route(BiffRecord record)
        {
            if (_handlers.TryGetValue(record.Id, out var handler))
            {
                handler(record);
                return true;
            }
            return false;
        }

        /// <summary>
        /// 尝试路由记录，如果没有找到处理器则执行默认操作
        /// </summary>
        /// <param name="record">BIFF记录</param>
        /// <param name="defaultAction">默认操作</param>
        public void RouteOrDefault(BiffRecord record, Action<BiffRecord> defaultAction)
        {
            if (!Route(record))
            {
                defaultAction(record);
            }
        }

        /// <summary>
        /// 清空所有处理器
        /// </summary>
        public void Clear()
        {
            _handlers.Clear();
        }

        /// <summary>
        /// 检查是否注册了指定类型的处理器
        /// </summary>
        /// <param name="recordType">记录类型ID</param>
        /// <returns>是否已注册</returns>
        public bool IsRegistered(ushort recordType)
        {
            return _handlers.ContainsKey(recordType);
        }
    }
}
