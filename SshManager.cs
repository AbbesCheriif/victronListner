using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Renci.SshNet;

namespace victronListner
{
    /// <summary>
    /// Gestionnaire pour les opérations SSH
    /// </summary>
    public class SshManager
    {
        /// <summary>
        /// Récupère la liste des services Modbus via SSH
        /// </summary>
        public async Task<List<string>> GetModbusServices()
        {
            var services = new List<string>();

            try
            {
                using (var client = new SshClient(Configuration.SSH_HOST, Configuration.SSH_USERNAME, Configuration.SSH_PASSWORD))
                {
                    await Task.Run(() => client.Connect());

                    var command = client.CreateCommand("dbus -y | grep com.victronenergy");
                    string result = command.Execute();

                    services = result.Split('\n')
                        .Where(line => !string.IsNullOrWhiteSpace(line))
                        .ToList();

                    client.Disconnect();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erreur SSH: {ex.Message}");
            }

            return services;
        }

        /// <summary>
        /// Récupère l'instance d'un appareil via SSH
        /// </summary>
        public async Task<string> GetDeviceInstance(string service)
        {
            try
            {
                using (var client = new SshClient(Configuration.SSH_HOST, Configuration.SSH_USERNAME, Configuration.SSH_PASSWORD))
                {
                    await Task.Run(() => client.Connect());

                    var command = client.CreateCommand($"dbus -y {service} /DeviceInstance GetValue");
                    string result = command.Execute();

                    client.Disconnect();
                    return result.Trim();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erreur SSH pour {service}: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Récupère tous les appareils connectés avec leurs instances
        /// </summary>
        public async Task<Dictionary<string, int>> GetConnectedDevices()
        {
            var devices = new Dictionary<string, int>();
            var services = await GetModbusServices();

            foreach (string service in services)
            {
                string deviceInstance = await GetDeviceInstance(service);
                if (!string.IsNullOrEmpty(deviceInstance))
                {
                    int instanceId = NetworkHelper.ExtractNumbers(deviceInstance);
                    devices[service] = instanceId;
                }
            }

            return devices;
        }
    }
}