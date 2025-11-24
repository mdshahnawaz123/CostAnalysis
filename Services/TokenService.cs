using Newtonsoft.Json;
using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace CostAnalysis.Services
{
    public class LocalAuthToken
    {
        public string Username { get; set; }
        public string MachineId { get; set; }
        public DateTime ExpiresUtc { get; set; }
    }

    public static class TokenService
    {
        private static readonly string TokenFolder =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CostAnalysis");
        private static readonly string TokenPath = Path.Combine(TokenFolder, "auth.token");

        private static readonly DataProtectionScope Scope = DataProtectionScope.CurrentUser;

        public static void SaveToken(LocalAuthToken token)
        {
            Directory.CreateDirectory(TokenFolder);
            var json = JsonConvert.SerializeObject(token);
            var bytes = Encoding.UTF8.GetBytes(json);
            var protectedBytes = ProtectedData.Protect(bytes, null, Scope);
            File.WriteAllBytes(TokenPath, protectedBytes);
        }

        public static LocalAuthToken LoadToken()
        {
            try
            {
                if (!File.Exists(TokenPath)) return null;
                var protectedBytes = File.ReadAllBytes(TokenPath);
                var bytes = ProtectedData.Unprotect(protectedBytes, null, Scope);
                var json = Encoding.UTF8.GetString(bytes);
                return JsonConvert.DeserializeObject<LocalAuthToken>(json);
            }
            catch
            {
                return null;
            }
        }

        public static void DeleteToken()
        {
            try { if (File.Exists(TokenPath)) File.Delete(TokenPath); } catch { }
        }
    }
}
