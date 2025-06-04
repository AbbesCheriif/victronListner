using System;
using System.Linq;

namespace victronListner
{
    /// <summary>
    /// Utilitaires pour les opérations réseau
    /// </summary>
    public static class NetworkHelper
    {
        /// <summary>
        /// Récupère l'adresse IP à partir d'une adresse MAC via ARP
        /// </summary>
        public static string GetIpFromMac(string macAddress)
        {
            try
            {
                var process = new System.Diagnostics.Process
                {
                    StartInfo = new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = "arp",
                        Arguments = "-a",
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        CreateNoWindow = true
                    }
                };

                process.Start();
                string output = process.StandardOutput.ReadToEnd();
                process.WaitForExit();

                foreach (string line in output.Split('\n'))
                {
                    if (line.ToLower().Contains(macAddress.ToLower()))
                    {
                        var parts = line.Trim().Split(' ');
                        if (parts.Length > 0)
                        {
                            return parts[0].Trim('(', ')');
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erreur lors de la détection IP: {ex.Message}");
            }

            return null;
        }

        /// <summary>
        /// Obtient le service générique à partir d'un nom de service complet
        /// </summary>
        public static string GetGenericService(string service)
        {
            if (service.Split('.').Length > 3)
            {
                int lastDotIndex = service.LastIndexOf('.');
                return service.Substring(0, lastDotIndex);
            }
            return service;
        }

        /// <summary>
        /// Extrait les nombres d'une chaîne de caractères
        /// </summary>
        public static int ExtractNumbers(string input)
        {
            string numbers = new string(input.Where(char.IsDigit).ToArray());
            return int.TryParse(numbers, out int result) ? result : 0;
        }
    }
}