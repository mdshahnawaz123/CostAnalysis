using Microsoft.Win32;
using System;

namespace CostAnalysis.Services
{
    public static class MachineHelper
    {
        public static string GetMachineId()
        {
            try
            {
                using (var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Cryptography"))
                {
                    var guid = key?.GetValue("MachineGuid") as string;
                    if (!string.IsNullOrWhiteSpace(guid)) return guid;
                }
            }
            catch { /* ignore */ }

            return Environment.MachineName ?? "unknown-machine";
        }
    }
}
