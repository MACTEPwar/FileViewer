using IdentityModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace FileViewer.Helpers
{
    public static class Sha256Helper
    {
        public static string ComputeHash(string codeVerifier)
        {
            using var sha256 = SHA256.Create();
            var challengeBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(codeVerifier));
            var codeChallenge = Base64Url.Encode(challengeBytes);
            return codeChallenge;
        }
    }
}
