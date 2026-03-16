using System.Collections.Generic;
using System.Linq;

namespace Nedev.FileConverters.XlsToXlsx.Formats.Xls
{
    /// <summary>
    /// 工作簿样式构建器 - 将各工作表的 XF 转换为全局 Workbook.Styles，并调整单元格 StyleId 为全局索引
    /// </summary>
    public class WorkbookStylesBuilder
    {
        /// <summary>
        /// 构建工作簿的全局样式列表
        /// </summary>
        /// <param name="workbook">工作簿对象</param>
        public void BuildWorkbookStyles(Workbook workbook)
        {
            // 合并所有工作表的调色板，以便后续查找颜色
            MergePalettes(workbook);

            // 获取要使用的 XF 列表（优先使用工作簿级别的 XfList）
            List<Xf> allXfs = GetAllXfs(workbook);

            // 从 XF 列表构建全局样式
            BuildStylesFromXfs(workbook, allXfs);

            // 更新单元格的 StyleId 值，使其指向全局样式列表
            UpdateCellStyleIds(workbook);
        }

        /// <summary>
        /// 合并所有工作表的调色板到工作簿级别
        /// </summary>
        private void MergePalettes(Workbook workbook)
        {
            foreach (var sheet in workbook.Worksheets)
            {
                foreach (var kv in sheet.Palette)
                {
                    if (!workbook.Palette.ContainsKey(kv.Key))
                        workbook.Palette[kv.Key] = kv.Value;
                }
            }
        }

        /// <summary>
        /// 获取所有 XF 对象列表
        /// </summary>
        private List<Xf> GetAllXfs(Workbook workbook)
        {
            List<Xf> allXfs = new List<Xf>();

            // 如果工作簿级别的 XfList 存在且不为空，则优先使用
            if (workbook.XfList != null && workbook.XfList.Count > 0)
            {
                allXfs.AddRange(workbook.XfList);
            }
            else
            {
                // 否则，按顺序收集所有工作表的 XF
                foreach (var sheet in workbook.Worksheets)
                {
                    allXfs.AddRange(sheet.Xfs);
                }
            }

            return allXfs;
        }

        /// <summary>
        /// 从 XF 列表构建全局样式
        /// </summary>
        private void BuildStylesFromXfs(Workbook workbook, List<Xf> allXfs)
        {
            foreach (var xf in allXfs)
            {
                var style = new Style();
                int fontIndex = NormalizeBiffFontIndex(xf.FontIndex);

                if (fontIndex >= 0 && fontIndex < workbook.Fonts.Count)
                    style.Font = workbook.Fonts[fontIndex];

                // future: copy other XF properties
                workbook.Styles.Add(style);
            }
        }

        /// <summary>
        /// 更新单元格的 StyleId 值，使其指向全局样式列表
        /// </summary>
        private void UpdateCellStyleIds(Workbook workbook)
        {
            foreach (var sheet in workbook.Worksheets)
            {
                foreach (var sheetRow in sheet.Rows)
                {
                    foreach (var cell in sheetRow.Cells ?? new List<Cell>())
                    {
                        if (!string.IsNullOrEmpty(cell.StyleId) && int.TryParse(cell.StyleId, out int idx))
                        {
                            // 使用 workbook.XfList 时索引已经是全局的，不需要偏移
                            cell.StyleId = idx.ToString();
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 标准化 BIFF 字体索引
        /// </summary>
        private static int NormalizeBiffFontIndex(int fontIndex)
        {
            return ParsingHelpers.NormalizeBiffFontIndex(fontIndex);
        }
    }
}
