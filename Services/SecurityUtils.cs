using System;
using System.Text;

namespace CostAnalysis.Services
{
    public static class SecurityUtils
    {
        // A simple key for XOR - keep this private
        private const int Key = 0xBDD; 

        /// <summary>
        /// Simple XOR "encryption" to hide strings from plain-text scanners.
        /// </summary>
        public static string Protect(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            var result = new StringBuilder();
            for (int i = 0; i < text.Length; i++)
            {
                result.Append((char)(text[i] ^ Key));
            }
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(result.ToString()));
        }

        /// <summary>
        /// Decodes the string back to plain text at runtime.
        /// </summary>
        public static string Unprotect(string protectedText)
        {
            if (string.IsNullOrEmpty(protectedText)) return string.Empty;
            try
            {
                var decoded = Encoding.UTF8.GetString(Convert.FromBase64String(protectedText));
                var result = new StringBuilder();
                for (int i = 0; i < decoded.Length; i++)
                {
                    result.Append((char)(decoded[i] ^ Key));
                }
                return result.ToString();
            }
            catch { return string.Empty; }
        }
    }
}
