using System.Text;
using System.Security.Cryptography;

namespace Web_QM.Common
{
    public static class CheckPass
    {
        public static string ComputeSha256Hash(this string username, string salt, string pass)
        {
            var crypt = new SHA256Managed();
            var hash = new StringBuilder();
            var crypto = crypt.ComputeHash(Encoding.UTF8.GetBytes($"{username} - {salt} - {pass}"));
            foreach (var theByte in crypto) hash.Append(theByte.ToString("x2"));
            return hash.ToString();
        }

        public static bool ValidPassword(this string username, string salt, string pass, string passHash)
        {
            var isValid = ComputeSha256Hash(username, salt, pass).Equals(passHash);
            return isValid;
        }
    }
}
