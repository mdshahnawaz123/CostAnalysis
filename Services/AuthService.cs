using CostAnalysis.Model;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;

namespace CostAnalysis.Services
{
    public class AuthService
    {
        private readonly string _source;
        private List<UserRecord> _users = new List<UserRecord>();
        private readonly int _httpTimeoutSeconds = 8;

        public UserRecord CurrentUser { get; set; }

        public AuthService(string source)
        {
            _source = source ?? throw new ArgumentNullException(nameof(source));
            TryLoadRemote().GetAwaiter().GetResult();
        }

        private bool IsUrl(string s)
        {
            return !string.IsNullOrWhiteSpace(s) &&
                   (s.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                    s.StartsWith("https://", StringComparison.OrdinalIgnoreCase));
        }

        public async Task<bool> TryLoadRemote()
        {
            try
            {
                if (!IsUrl(_source))
                {
                    if (!File.Exists(_source)) return false;
                    var jsonLocal = File.ReadAllText(_source);
                    _users = JsonConvert.DeserializeObject<List<UserRecord>>(jsonLocal) ?? new List<UserRecord>();
                    return true;
                }

                using (var http = new HttpClient())
                {
                    http.Timeout = TimeSpan.FromSeconds(_httpTimeoutSeconds);

                    var pat = Environment.GetEnvironmentVariable("COSTANALYSIS_GITHUB_PAT");
                    if (!string.IsNullOrWhiteSpace(pat))
                    {
                        http.DefaultRequestHeaders.Authorization =
                            new System.Net.Http.Headers.AuthenticationHeaderValue("token", pat);
                    }

                    http.DefaultRequestHeaders.UserAgent.TryParseAdd("CostAnalysisAddin/1.0");

                    var resp = await http.GetAsync(_source).ConfigureAwait(false);
                    if (!resp.IsSuccessStatusCode) return false;
                    var json = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
                    _users = JsonConvert.DeserializeObject<List<UserRecord>>(json) ?? new List<UserRecord>();
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        public bool ValidateCredentials(string username, string password, out UserRecord matched, out string error)
        {
            matched = null; error = null;
            if (string.IsNullOrWhiteSpace(username) || string.IsNullOrWhiteSpace(password))
            {
                error = "Enter username and password.";
                return false;
            }

            matched = _users.FirstOrDefault(u => string.Equals(u.Username, username, StringComparison.OrdinalIgnoreCase));
            if (matched == null)
            {
                error = "Unknown user.";
                return false;
            }

            if (matched.Password != password)
            {
                error = "Invalid password.";
                return false;
            }

            if (!matched.Active)
            {
                error = "Account inactive.";
                return false;
            }

            if (matched.Expires.Date < DateTime.UtcNow.Date)
            {
                error = "Account expired.";
                return false;
            }

            CurrentUser = matched;
            return true;
        }

        public UserRecord GetUser(string username)
        {
            return _users.FirstOrDefault(u => string.Equals(u.Username, username, StringComparison.OrdinalIgnoreCase));
        }
    }
}
