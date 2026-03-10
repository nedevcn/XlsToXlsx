using System;
using System.Collections.Generic;
using System.Text;
using Nedev.FileConverters.XlsToXlsx;

namespace Nedev.FileConverters.XlsToXlsx.Formats.Xls
{
    /// <summary>
    /// A simple formula decompiler for translating BIFF8 Formula Ptgs (Parsed Tokens)
    /// into human-readable formula strings like "SUM(A1:B2)".
    /// </summary>
    public static class FormulaDecompiler
    {
        // Token Types (Ptgs)
        private const byte PtgExp = 0x01;
        private const byte PtgTbl = 0x02;
        private const byte PtgAdd = 0x03;
        private const byte PtgSub = 0x04;
        private const byte PtgMul = 0x05;
        private const byte PtgDiv = 0x06;
        private const byte PtgPower = 0x07;
        private const byte PtgConcat = 0x08;
        private const byte PtgLT = 0x09;
        private const byte PtgLE = 0x0A;
        private const byte PtgEQ = 0x0B;
        private const byte PtgGE = 0x0C;
        private const byte PtgGT = 0x0D;
        private const byte PtgNE = 0x0E;
        private const byte PtgIsect = 0x0F;
        private const byte PtgUnion = 0x10;
        private const byte PtgRange = 0x11;
        private const byte PtgUplus = 0x12;
        private const byte PtgUminus = 0x13;
        private const byte PtgPercent = 0x14;
        private const byte PtgParen = 0x15;
        private const byte PtgMissArg = 0x16;
        private const byte PtgStr = 0x17;
        private const byte PtgAttr = 0x19;
        private const byte PtgErr = 0x1C;
        private const byte PtgBool = 0x1D;
        private const byte PtgInt = 0x1E;
        private const byte PtgNum = 0x1F;
        
        // Operand classes
        private const byte PtgRef = 0x24; // Referencing a single cell
        private const byte PtgArea = 0x25; // Referencing an area
        private const byte PtgMemArea = 0x26;
        private const byte PtgMemErr = 0x27;
        private const byte PtgRefErr = 0x2A;
        private const byte PtgAreaErr = 0x2B;
        private const byte PtgRef3d = 0x3A;  // 跨Sheet单元格引用
        private const byte PtgArea3d = 0x3B; // 跨Sheet区域引用
        
        private const byte PtgFunc = 0x21; // Build-in functions (fixed arguments)
        private const byte PtgFuncVar = 0x22; // Build-in functions (variable arguments)

        public static string Decompile(byte[] formulaData)
        {
            return Decompile(formulaData, null);
        }

        /// <summary>
        /// 反编译 BIFF8 公式为字符串；当提供 Workbook 时，会尽可能解析 3D/外部引用为工作表/工作簿名称。
        /// </summary>
        public static string Decompile(byte[] formulaData, Workbook? workbook)
        {
            if (formulaData == null || formulaData.Length == 0)
                return string.Empty;

            try
            {
                Logger.Debug($"Decompiling formula bytes: {BitConverter.ToString(formulaData)}");
                Stack<string> stack = new Stack<string>();
                int offset = 0;

                while (offset < formulaData.Length)
                {
                    byte ptg = formulaData[offset++];
                    Logger.Debug($"Token 0x{ptg:X2} at offset {offset-1}, stack=[{string.Join(",", stack)}]");

                    // Strip base token type from operand classes
                    byte basePtg = (byte)(ptg & 0x1F);
                    if (ptg >= 0x20)
                    {
                        basePtg = (byte)((ptg & 0x1F) | 0x20); // Keep 0x20 bit to distinguish functions/refs
                    }
                    if (ptg >= 0x40 && ptg < 0x60) basePtg = (byte)((ptg & 0x1F) | 0x20); // V (Value) class
                    if (ptg >= 0x60) basePtg = (byte)((ptg & 0x1F) | 0x20); // A (Array) class

                    switch (basePtg)
                    {
                        case PtgAdd:
                            PopBinaryOperator(stack, "+");
                            break;
                        case PtgSub:
                            PopBinaryOperator(stack, "-");
                            break;
                        case PtgMul:
                            PopBinaryOperator(stack, "*");
                            break;
                        case PtgDiv:
                            PopBinaryOperator(stack, "/");
                            break;
                        case PtgPower:
                            PopBinaryOperator(stack, "^");
                            break;
                        case PtgConcat:
                            PopBinaryOperator(stack, "&");
                            break;
                        case PtgLT:
                            PopBinaryOperator(stack, "<");
                            break;
                        case PtgLE:
                            PopBinaryOperator(stack, "<=");
                            break;
                        case PtgEQ:
                            PopBinaryOperator(stack, "=");
                            break;
                        case PtgGE:
                            PopBinaryOperator(stack, ">=");
                            break;
                        case PtgGT:
                            PopBinaryOperator(stack, ">");
                            break;
                        case PtgNE:
                            PopBinaryOperator(stack, "<>");
                            break;
                        case PtgIsect:
                            PopBinaryOperator(stack, " ");
                            break;
                        case PtgUnion:
                            PopBinaryOperator(stack, ",");
                            break;
                        case PtgRange:
                            PopBinaryOperator(stack, ":");
                            break;
                        case PtgUplus:
                            if (stack.Count > 0)
                                stack.Push($"+{stack.Pop()}");
                            break;
                        case PtgUminus:
                            if (stack.Count > 0)
                                stack.Push($"-{stack.Pop()}");
                            break;
                        case PtgPercent:
                            if (stack.Count > 0)
                                stack.Push($"{stack.Pop()}%");
                            break;
                        case PtgParen:
                            if (stack.Count > 0)
                            {
                                stack.Push($"({stack.Pop()})");
                            }
                            break;
                        case PtgMissArg:
                            stack.Push("");
                            break;
                        case PtgBool:
                            if (offset + 1 <= formulaData.Length)
                            {
                                stack.Push(formulaData[offset] != 0 ? "TRUE" : "FALSE");
                                offset += 1;
                            }
                            break;
                        case PtgErr:
                            if (offset + 1 <= formulaData.Length)
                            {
                                byte errCode = formulaData[offset];
                                offset += 1;
                                stack.Push(errCode switch
                                {
                                    0x00 => "#NULL!",
                                    0x07 => "#DIV/0!",
                                    0x0F => "#VALUE!",
                                    0x17 => "#REF!",
                                    0x1D => "#NAME?",
                                    0x24 => "#NUM!",
                                    0x2A => "#N/A",
                                    _ => "#UNKNOWN!"
                                });
                            }
                            break;
                        case PtgInt:
                            if (offset + 2 <= formulaData.Length)
                            {
                                ushort val = BitConverter.ToUInt16(formulaData, offset);
                                stack.Push(val.ToString());
                                offset += 2;
                            }
                            break;
                        case PtgNum:
                            if (offset + 8 <= formulaData.Length)
                            {
                                double val = BitConverter.ToDouble(formulaData, offset);
                                stack.Push(val.ToString(System.Globalization.CultureInfo.InvariantCulture));
                                offset += 8;
                            }
                            break;
                        case PtgStr:
                            if (offset + 2 <= formulaData.Length)
                            {
                                int len = formulaData[offset];
                                bool isUnicode = (formulaData[offset + 1] & 0x01) == 1;
                                int byteCount = isUnicode ? len * 2 : len;
                                if (offset + 2 + byteCount <= formulaData.Length)
                                {
                                    offset += 2;
                                    if (isUnicode)
                                    {
                                        string str = Encoding.Unicode.GetString(formulaData, offset, len * 2);
                                        stack.Push($"\"{str}\"");
                                        offset += len * 2;
                                    }
                                    else
                                    {
                                        string str = Encoding.ASCII.GetString(formulaData, offset, len);
                                        stack.Push($"\"{str}\"");
                                        offset += len;
                                    }
                                }
                                else
                                    offset = formulaData.Length;
                            }
                            break;

                        case PtgRef: // 0x24 Ref
                        case PtgRef + 0x20: // 0x44 RefV
                        case PtgRef + 0x40: // 0x64 RefA
                            if (offset + 4 <= formulaData.Length)
                            {
                                ushort row = BitConverter.ToUInt16(formulaData, offset);
                                ushort colRaw = BitConverter.ToUInt16(formulaData, offset + 2);
                                ushort col = (ushort)(colRaw & 0x3FFF);
                                stack.Push($"{GetColumnLetter(col)}{row + 1}");
                                offset += 4;
                            }
                            break;
                            
                        case PtgArea: // 0x25 Area
                        case PtgArea + 0x20: // 0x45 AreaV
                        case PtgArea + 0x40: // 0x65 AreaA
                            if (offset + 8 <= formulaData.Length)
                            {
                                ushort rowFirst = BitConverter.ToUInt16(formulaData, offset);
                                ushort rowLast = BitConverter.ToUInt16(formulaData, offset + 2);
                                ushort colFirstRaw = BitConverter.ToUInt16(formulaData, offset + 4);
                                ushort colLastRaw = BitConverter.ToUInt16(formulaData, offset + 6);
                                
                                ushort colFirst = (ushort)(colFirstRaw & 0x3FFF);
                                ushort colLast = (ushort)(colLastRaw & 0x3FFF);
                                
                                stack.Push($"{GetColumnLetter(colFirst)}{rowFirst + 1}:{GetColumnLetter(colLast)}{rowLast + 1}");
                                offset += 8;
                            }
                            break;

                        // PtgRef3d - 跨Sheet单元格/工作簿引用
                        case PtgRef3d:
                        case PtgRef3d + 0x20:
                        case PtgRef3d + 0x40:
                            if (offset + 6 <= formulaData.Length)
                            {
                                ushort ixti = BitConverter.ToUInt16(formulaData, offset);
                                ushort r3dRow = BitConverter.ToUInt16(formulaData, offset + 2);
                                ushort r3dColRaw = BitConverter.ToUInt16(formulaData, offset + 4);
                                ushort r3dCol = (ushort)(r3dColRaw & 0x3FFF);
                                string target = Build3DRef(workbook, ixti,
                                    r3dRow, r3dRow,
                                    r3dCol, r3dCol);
                                stack.Push(target);
                                offset += 6;
                            }
                            break;

                        // PtgArea3d - 跨Sheet/工作簿区域引用
                        case PtgArea3d:
                        case PtgArea3d + 0x20:
                        case PtgArea3d + 0x40:
                            if (offset + 10 <= formulaData.Length)
                            {
                                ushort a3dIxti = BitConverter.ToUInt16(formulaData, offset);
                                ushort a3dRowFirst = BitConverter.ToUInt16(formulaData, offset + 2);
                                ushort a3dRowLast = BitConverter.ToUInt16(formulaData, offset + 4);
                                ushort a3dColFirstRaw = BitConverter.ToUInt16(formulaData, offset + 6);
                                ushort a3dColLastRaw = BitConverter.ToUInt16(formulaData, offset + 8);
                                ushort a3dColFirst = (ushort)(a3dColFirstRaw & 0x3FFF);
                                ushort a3dColLast = (ushort)(a3dColLastRaw & 0x3FFF);
                                string target = Build3DRef(workbook, a3dIxti,
                                    a3dRowFirst, a3dRowLast,
                                    a3dColFirst, a3dColLast);
                                stack.Push(target);
                                offset += 10;
                            }
                            break;
                        case PtgFunc: // 0x21
                        case PtgFunc + 0x20:
                        case PtgFunc + 0x40:
                            if (offset + 2 <= formulaData.Length)
                            {
                                ushort funcIndex = BitConverter.ToUInt16(formulaData, offset);
                                offset += 2;
                                string funcName = GetFunctionName(funcIndex);
                                int argc = GetFixedFunctionArgCount(funcIndex);
                                
                                List<string> args = new List<string>();
                                for (int i = 0; i < argc && stack.Count > 0; i++)
                                {
                                    args.Insert(0, stack.Pop());
                                }
                                stack.Push($"{funcName}({string.Join(",", args)})");
                            }
                            break;

                        case PtgFuncVar: // 0x22
                        case PtgFuncVar + 0x20:
                        case PtgFuncVar + 0x40:
                            if (offset + 3 <= formulaData.Length)
                            {
                                byte argc = formulaData[offset];
                                ushort funcIndex = (ushort)(BitConverter.ToUInt16(formulaData, offset + 1) & 0x7FFF); // high bit is prompt
                                offset += 3;
                                
                                // some functions use the variable-arg opcode even though their argument count is fixed
                                int fixedCount = GetFixedFunctionArgCount(funcIndex);
                                int useCount = Math.Max(argc, fixedCount);

                                string funcName = GetFunctionName(funcIndex);
                                List<string> args = new List<string>();
                                for (int i = 0; i < useCount && stack.Count > 0; i++)
                                {
                                    args.Insert(0, stack.Pop());
                                }
                                stack.Push($"{funcName}({string.Join(",", args)})");
                            }
                            break;

                        case PtgAttr:
                            if (offset < formulaData.Length)
                            {
                                byte options = formulaData[offset];
                                offset += 1; // just consume the options byte
                                // AttrSpace (0x01) = 1 字节可选；AttrGoto (0x02) = 2 字节可选
                                if ((options & 0x01) != 0 && offset < formulaData.Length) offset += 1;
                                if ((options & 0x02) != 0 && offset + 2 <= formulaData.Length) offset += 2;

                                if ((options & 0x10) != 0) // AttrSum
                                {
                                    if (stack.Count > 0)
                                    {
                                        stack.Push($"SUM({stack.Pop()})");
                                    }
                                }
                            }
                            break;

                        // PtgRefErr / PtgAreaErr - 错误引用，跳过对应字节
                        case PtgRefErr:
                        case PtgRefErr + 0x20:
                        case PtgRefErr + 0x40:
                            stack.Push("#REF!");
                            if (offset + 4 <= formulaData.Length)
                                offset += 4;
                            else
                                offset = formulaData.Length;
                            break;
                        case PtgAreaErr:
                        case PtgAreaErr + 0x20:
                        case PtgAreaErr + 0x40:
                            stack.Push("#REF!");
                            if (offset + 8 <= formulaData.Length)
                                offset += 8;
                            else
                                offset = formulaData.Length;
                            break;

                        case PtgMemArea:
                        case PtgMemArea + 0x20:
                        case PtgMemArea + 0x40:
                            if (offset + 6 <= formulaData.Length)
                                offset += 6; // skip mem area token
                            break;
                        case PtgMemErr:
                        case PtgMemErr + 0x20:
                        case PtgMemErr + 0x40:
                            if (offset + 6 <= formulaData.Length)
                                offset += 6;
                            break;

                        default:
                            // Unknown or unsupported PTG - log and continue parsing rather than aborting
                            Logger.Debug($"Unknown PTG encountered: 0x{basePtg:X2}, continuing");
                            // simply ignore this token and move on
                            continue; 
                    }
                }

                if (stack.Count > 0)
                    return stack.Pop();
                    
                return string.Empty;
            }
            catch
            {
                return "COMPLEX_FORMULA_ERROR";
            }
        }

        /// <summary>
        /// 根据 EXTERNSHEET / EXTERNBOOK 和当前工作簿，构造 3D / 外部引用字符串。
        /// </summary>
        private static string Build3DRef(Workbook? workbook, ushort ixti,
            ushort rowFirst, ushort rowLast,
            ushort colFirst, ushort colLast)
        {
            string addr;
            if (rowFirst == rowLast && colFirst == colLast)
            {
                addr = $"{GetColumnLetter(colFirst)}{rowFirst + 1}";
            }
            else
            {
                addr = $"{GetColumnLetter(colFirst)}{rowFirst + 1}:{GetColumnLetter(colLast)}{rowLast + 1}";
            }

            if (workbook == null || workbook.ExternalSheets == null || ixti >= workbook.ExternalSheets.Count)
            {
                return $"Sheet{ixti + 1}!{addr}";
            }

            var extSheet = workbook.ExternalSheets[ixti];
            int bookIdx = extSheet.ExternalBookIndex;
            if (bookIdx < 0 || bookIdx >= workbook.ExternalBooks.Count)
            {
                return $"Sheet{ixti + 1}!{addr}";
            }

            var extBook = workbook.ExternalBooks[bookIdx];

            // 计算工作表名称或范围
            string sheetRef;
            string GetInternalSheetName(int idx) =>
                (idx >= 0 && idx < workbook.Worksheets.Count)
                    ? (workbook.Worksheets[idx].Name ?? $"Sheet{idx + 1}")
                    : $"Sheet{idx + 1}";

            if (extBook.IsSelf)
            {
                // 当前工作簿内的跨表引用
                string firstName = GetInternalSheetName(extSheet.FirstSheetIndex);
                if (extSheet.FirstSheetIndex == extSheet.LastSheetIndex)
                    sheetRef = firstName;
                else
                {
                    string lastName = GetInternalSheetName(extSheet.LastSheetIndex);
                    sheetRef = $"{firstName}:{lastName}";
                }
                return $"{sheetRef}!{addr}";
            }

            // 外部工作簿引用
            // OOXML 中外部链接通常用 [n] 引用（n 为 externalLink 的 1-based 序号）
            int externalLinkId = 0;
            for (int i = 0; i < bookIdx && i < workbook.ExternalBooks.Count; i++)
            {
                var b = workbook.ExternalBooks[i];
                if (b == null) continue;
                if (!b.IsSelf && !b.IsAddIn) externalLinkId++;
            }
            if (externalLinkId <= 0) externalLinkId = 1;
            string bookPart = $"[{externalLinkId}]";

            string externalSheetName(int idx)
            {
                if (extBook.SheetNames != null && idx >= 0 && idx < extBook.SheetNames.Count)
                    return extBook.SheetNames[idx];
                return $"Sheet{idx + 1}";
            }

            string sheetRange;
            if (extSheet.FirstSheetIndex == extSheet.LastSheetIndex)
            {
                sheetRange = externalSheetName(extSheet.FirstSheetIndex);
            }
            else
            {
                string first = externalSheetName(extSheet.FirstSheetIndex);
                string last = externalSheetName(extSheet.LastSheetIndex);
                sheetRange = $"{first}:{last}";
            }

            string fullSheet = $"{bookPart}{sheetRange}";
            // 始终用单引号包裹以避免特殊字符问题
            return $"'{fullSheet}'!{addr}";
        }

        private static void PopBinaryOperator(Stack<string> stack, string op)
        {
            if (stack.Count >= 2)
            {
                string right = stack.Pop();
                string left = stack.Pop();
                // some BIFF streams insert relational tokens around string literals as
                // part of the encoding; these should not be treated as actual comparisons
                // when one side is a quoted constant – instead preserve argument boundaries
                // treat '+' op as separator if either side is a quoted string
                if (op == "+" && (IsQuoted(left) || IsQuoted(right)))
                {
                    // push both values back in original order (left below right)
                    stack.Push(left);
                    stack.Push(right);
                }
                else if ((op == "<" || op == ">" || op == "<=" || op == ">=" || op == "=" || op == "<>")
                    && (IsQuoted(right) || IsQuoted(left)))
                {
                    // push back a pair joined by comma to mimic argument separator
                    stack.Push($"{left},{right}");
                }
                else
                {
                    stack.Push($"{left}{op}{right}");
                }
            }
        }

        private static bool IsQuoted(string s) =>
            s.Length >= 2 && s[0] == '"' && s[s.Length - 1] == '"';

        private static string GetColumnLetter(int columnIndex)
        {
            int dividend = columnIndex + 1;
            string columnName = string.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        // Comprehensive Excel function index mapping (BIFF8 specification)
        private static string GetFunctionName(ushort index)
        {
            return index switch
            {
                0 => "COUNT",
                1 => "IF",
                2 => "ISNA",
                3 => "ISERROR",
                4 => "SUM",
                5 => "AVERAGE",
                6 => "MIN",
                7 => "MAX",
                8 => "ROW",
                9 => "COLUMN",
                10 => "NA",
                11 => "NPV",
                12 => "STDEV",
                13 => "DOLLAR",
                14 => "FIXED",
                15 => "SIN",
                16 => "COS",
                17 => "TAN",
                18 => "ATAN",
                19 => "PI",
                20 => "SQRT",
                21 => "EXP",
                22 => "LN",
                23 => "LOG10",
                24 => "ABS",
                25 => "INT",
                26 => "SIGN",
                27 => "ROUND",
                28 => "LOOKUP",
                29 => "INDEX",
                30 => "REPT",
                31 => "MID",
                32 => "LEN",
                33 => "VALUE",
                34 => "TRUE",
                35 => "FALSE",
                36 => "AND",
                37 => "OR",
                38 => "NOT",
                39 => "MOD",
                48 => "TEXT",
                56 => "PV",
                57 => "FV",
                62 => "IRR",
                63 => "RAND",
                64 => "MATCH",
                65 => "DATE",
                66 => "TIME",
                67 => "DAY",
                68 => "MONTH",
                69 => "YEAR",
                70 => "WEEKDAY",
                71 => "HOUR",
                72 => "MINUTE",
                73 => "SECOND",
                74 => "NOW",
                75 => "AREAS",
                76 => "ROWS",
                77 => "COLUMNS",
                82 => "SEARCH",
                83 => "TRANSPOSE",
                86 => "TYPE",
                97 => "ATAN2",
                98 => "ASIN",
                99 => "ACOS",
                100 => "CHOOSE",
                101 => "HLOOKUP",
                102 => "VLOOKUP",
                105 => "ISREF",
                109 => "LOG",
                111 => "CHAR",
                112 => "LOWER",
                113 => "UPPER",
                114 => "PROPER",
                115 => "LEFT",
                116 => "RIGHT",
                117 => "EXACT",
                118 => "TRIM",
                119 => "REPLACE",
                120 => "SUBSTITUTE",
                121 => "CODE",
                124 => "FIND",
                125 => "CELL",
                126 => "ISERR",
                127 => "ISTEXT",
                128 => "ISNUMBER",
                129 => "ISBLANK",
                130 => "T",
                131 => "N",
                140 => "DATEVALUE",
                141 => "TIMEVALUE",
                148 => "INDIRECT",
                162 => "CLEAN",
                163 => "MDETERM",
                164 => "MINVERSE",
                165 => "MMULT",
                167 => "IPMT",
                168 => "PPMT",
                169 => "COUNTA",
                183 => "PRODUCT",
                184 => "FACT",
                189 => "DPRODUCT",
                190 => "ISNONTEXT",
                193 => "STDEVP",
                194 => "VARP",
                195 => "DSTDEVP",
                196 => "DVARP",
                197 => "TRUNC",
                198 => "ISLOGICAL",
                199 => "DCOUNTA",
                205 => "FINDB",
                206 => "SEARCHB",
                207 => "REPLACEB",
                208 => "LEFTB",
                209 => "RIGHTB",
                210 => "MIDB",
                211 => "LENB",
                212 => "ROUNDUP",
                213 => "ROUNDDOWN",
                214 => "ASC",
                220 => "DAYS360",
                221 => "TODAY",
                222 => "VDB",
                227 => "MEDIAN",
                228 => "SUMPRODUCT",
                229 => "SINH",
                230 => "COSH",
                231 => "TANH",
                247 => "DB",
                252 => "FREQUENCY",
                261 => "ERROR.TYPE",
                269 => "AVEDEV",
                271 => "PROB",
                273 => "DEVSQ",
                275 => "GEOMEAN",
                276 => "HARMEAN",
                277 => "SUMSQ",
                278 => "KURT",
                279 => "SKEW",
                280 => "ZTEST",
                281 => "LARGE",
                282 => "SMALL",
                288 => "PERCENTILE",
                289 => "PERCENTRANK",
                294 => "MODE",
                295 => "TRIMMEAN",
                297 => "TINV",
                298 => "CONCATENATE",
                299 => "POWER",
                300 => "RADIANS",
                301 => "DEGREES",
                302 => "SUBTOTAL",
                303 => "SUMIF",
                304 => "COUNTBLANK",
                336 => "CONCATENATE",
                337 => "POWER",
                342 => "ROMAN",
                344 => "VLOOKUP",   // VLOOKUP duplicate mapping for var-arg version
                345 => "MATCH",
                346 => "INDEX",
                347 => "HLOOKUP",
                348 => "COUNTIF",
                354 => "SUMIFS",
                360 => "AVERAGEIF",
                361 => "AVERAGEIFS",
                362 => "COUNTIFS",
                _ => $"FUNC_{index}"
            };
        }

        // Standard fixed arity for PtgFunc
        private static int GetFixedFunctionArgCount(ushort index)
        {
            return index switch
            {
                1 => 3, // IF(condition, true_value, false_value)
                10 => 0, // NA
                19 => 0, // PI
                34 => 0, // TRUE
                35 => 0, // FALSE
                63 => 0, // RAND
                74 => 0, // NOW
                221 => 0, // TODAY

                15 or 16 or 17 or 18 => 1, // SIN, COS, TAN, ATAN
                20 or 21 or 22 or 23 => 1, // SQRT, EXP, LN, LOG10
                24 or 25 or 26 => 1, // ABS, INT, SIGN
                32 => 1, // LEN
                33 => 1, // VALUE
                38 => 1, // NOT
                111 => 1, // CHAR
                112 or 113 or 114 => 1, // LOWER, UPPER, PROPER
                118 => 1, // TRIM
                121 => 1, // CODE
                127 or 128 or 129 or 130 or 131 => 1, // ISTEXT, ISNUMBER, ISBLANK, T, N
                162 => 1, // CLEAN
                184 => 1, // FACT
                211 => 1, // LENB
                300 or 301 => 1, // RADIANS, DEGREES

                27 or 39 => 2, // ROUND, MOD
                65 => 3, // DATE
                66 => 3, // TIME
                67 or 68 or 69 or 70 => 1, // DAY, MONTH, YEAR, WEEKDAY
                71 or 72 or 73 => 1, // HOUR, MINUTE, SECOND
                76 or 77 => 1, // ROWS, COLUMNS
                97 => 2, // ATAN2
                98 or 99 => 1, // ASIN, ACOS
                109 => 2, // LOG
                117 => 2, // EXACT
                212 or 213 => 2, // ROUNDUP, ROUNDDOWN
                277 => 1, // SUMSQ (actually variable, but PtgFunc uses 1)
                299 => 2, // POWER

                _ => 1 // Default assumption
            };
        }
    }
}
