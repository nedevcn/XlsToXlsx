using System;
using System.Security.Cryptography;
using System.Text;
using Nedev.FileConverters.XlsToXlsx.Exceptions;

namespace Nedev.FileConverters.XlsToXlsx.Formats.Xls
{
    /// <summary>
    /// 提供对加密的 XLS (BIFF8) 记录进行处理的解密工具
    /// </summary>
    public class XlsDecryptor
    {
        private byte[] _baseKey;
        private byte[] _activeKey;
        private int _currentBlock = -1;
        private byte[] _sBox = new byte[256];
        
        public XlsDecryptor(byte[] encryptionData, string password)
        {
            // 解析 FILEPASS 记录数据
            // BIFF8 Standard Encryption:
            // 0-1: 1 (RC4)
            // 2-3: 1 (Standard)
            // 4-19: Salt
            // 20-35: Encrypted Verifier
            // 36-51: Encrypted Verifier Hash
            
            if (encryptionData.Length < 52)
                throw new XlsParseException("Invalid FILEPASS record data");
                
            byte[] salt = new byte[16];
            Array.Copy(encryptionData, 4, salt, 0, 16);
            
            _baseKey = DeriveKey(password, salt);
            _activeKey = new byte[_baseKey.Length];
        }

        private byte[] DeriveKey(string password, byte[] salt)
        {
            // BIFF8 Standard Encryption Key Derivation:
            // 1. Password to UTF-16LE
            // 2. MD5 hash of password
            // 3. MD5 hash of (hash + salt) multiple times? 
            // 实际上基础版是: MD5(PasswordBytes + Salt)
            
            using (MD5 md5 = MD5.Create())
            {
                byte[] passwordBytes = Encoding.Unicode.GetBytes(password);
                byte[] buffer = new byte[passwordBytes.Length + salt.Length];
                Buffer.BlockCopy(passwordBytes, 0, buffer, 0, passwordBytes.Length);
                Buffer.BlockCopy(salt, 0, buffer, passwordBytes.Length, salt.Length);
                
                byte[] hash = md5.ComputeHash(buffer);
                
                // 只取前 5 字节 (40-bit key) 或根据版本取更多
                byte[] key = new byte[16];
                Array.Copy(hash, 0, key, 0, Math.Min(hash.Length, 16));
                return key;
            }
        }

        /// <summary>
        /// 解密特定位置的数据块
        /// </summary>
        /// <param name="data">要解密的字节数组</param>
        /// <param name="streamPosition">该数据在 Workbook 流中的起始位置</param>
        public void Decrypt(byte[] data, long streamPosition)
        {
            for (int i = 0; i < data.Length; i++)
            {
                long currentPos = streamPosition + i;
                int currentBlock = (int)(currentPos / 1024);
                int offsetInBlock = (int)(currentPos % 1024);

                if (currentBlock != _currentBlock)
                {
                    ResetRC4(currentBlock);
                    _lastOffset = 0;
                }
                else if (offsetInBlock < _lastOffset)
                {
                    // 如果偏移往回走了（虽然在这个逻辑下不太可能，除非跨 block 处理有问题），也重置
                    ResetRC4(currentBlock);
                    _lastOffset = 0;
                }

                // 跳到当前字节的偏移
                if (offsetInBlock > _lastOffset)
                {
                    SkipRC4(offsetInBlock - _lastOffset);
                    _lastOffset = offsetInBlock;
                }
                
                data[i] ^= GetNextRC4Byte();
                _lastOffset++; // 保持偏移更新
            }
        }

        private int _lastOffset = 0;

        private void ResetRC4(int blockIndex)
        {
            _currentBlock = blockIndex;
            
            // 衍生当前 Block 的 Key: MD5(BaseKey + BlockIndex)
            using (MD5 md5 = MD5.Create())
            {
                byte[] blockBytes = BitConverter.GetBytes(blockIndex);
                byte[] buffer = new byte[_baseKey.Length + 4];
                Buffer.BlockCopy(_baseKey, 0, buffer, 0, _baseKey.Length);
                Buffer.BlockCopy(blockBytes, 0, buffer, _baseKey.Length, 4);
                
                _activeKey = md5.ComputeHash(buffer);
            }
            
            // 初始化 S-Box
            for (int i = 0; i < 256; i++) _sBox[i] = (byte)i;
            int j = 0;
            for (int i = 0; i < 256; i++)
            {
                j = (j + _sBox[i] + _activeKey[i % _activeKey.Length]) & 0xFF;
                byte temp = _sBox[i];
                _sBox[i] = _sBox[j];
                _sBox[j] = temp;
            }
            
            _rc4I = 0;
            _rc4J = 0;
        }

        private int _rc4I, _rc4J;

        private void SkipRC4(int count)
        {
            for (int i = 0; i < count; i++) GetNextRC4Byte();
        }

        private byte GetNextRC4Byte()
        {
            _rc4I = (_rc4I + 1) & 0xFF;
            _rc4J = (_rc4J + _sBox[_rc4I]) & 0xFF;
            
            byte temp = _sBox[_rc4I];
            _sBox[_rc4I] = _sBox[_rc4J];
            _sBox[_rc4J] = temp;
            
            return _sBox[(_sBox[_rc4I] + _sBox[_rc4J]) & 0xFF];
        }
    }
}
