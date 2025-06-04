using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Text.Json;
using System.IO;
using static OfficeOpenXml.ExcelErrorValue;
using System.Diagnostics;

namespace victronListner
{
    /// <summary>
    /// Gestionnaire du menu principal et des interactions utilisateur
    /// </summary>
    public class MenuManager
    {
        private readonly ExcelManager _excelManager;
        private readonly ModbusManager _modbusManager;
        private readonly DeviceManager _deviceManager;

        public string json { get; private set; }

        public MenuManager(ExcelManager excelManager, ModbusManager modbusManager, DeviceManager deviceManager)
        {
            _excelManager = excelManager;
            _modbusManager = modbusManager;
            _deviceManager = deviceManager;
        }

        /// <summary>
        /// Affiche et gère le menu principal
        /// </summary>
        public async Task readAll()
        {
            try
            {     
                await StartAutoPolling();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erreur: {ex.Message}");
            }
        }

        /// <summary>
        /// Teste la connexion Modbus
        /// </summary>
        private async Task TestConnection()
        {
            Console.Write("IP du Cerbo (laisser vide pour auto-détection): ");
            string ip = Console.ReadLine();

            if (string.IsNullOrEmpty(ip))
            {
                ip = NetworkHelper.GetIpFromMac(Configuration.VICTRON_MAC);
                if (string.IsNullOrEmpty(ip))
                {
                    Console.WriteLine("Impossible de détecter l'IP automatiquement");
                    return;
                }
            }

            Configuration.MODBUS_IP = ip;

            try
            {
                var modbusClient = await _modbusManager.ConnectModbus(ip, Configuration.MODBUS_PORT);
                if (modbusClient != null)
                {
                    Console.WriteLine($"Connexion réussie à {ip}:{Configuration.MODBUS_PORT}");
                }
                else
                {
                    Console.WriteLine("Échec de la connexion");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erreur de connexion: {ex.Message}");
            }
        }

        /// <summary>
        /// Affiche les services disponibles
        /// </summary>
        private void ShowAvailableServices()
        {
            Console.WriteLine("\n=== SERVICES DISPONIBLES ===");
            for (int i = 0; i < _excelManager.ServiceNames.Count; i++)
            {
                Console.WriteLine($"{i + 1}. {_excelManager.ServiceNames[i]}");
            }
        }

        /// <summary>
        /// Affiche les appareils connectés
        /// </summary>
        private void ShowConnectedDevices()
        {
            Console.WriteLine("\n=== APPAREILS CONNECTÉS ===");
            for (int i = 0; i < _deviceManager.AvailableServices.Count; i++)
            {
                Console.WriteLine($"{i + 1}. {_deviceManager.AvailableServices[i]}");
            }
        }

        /// <summary>
        /// Interface pour lire un paramètre
        /// </summary>
        private async Task ReadSetting()
        {
            Console.WriteLine("\n=== LECTURE D'UN PARAMÈTRE ===");

            // Sélection du service
            Console.WriteLine("Services disponibles:");
            for (int i = 0; i < _deviceManager.AvailableServices.Count; i++)
            {
                Console.WriteLine($"{i + 1}. {_deviceManager.AvailableServices[i]}");
            }

            Console.Write("Sélectionnez un service (numéro ou nom): ");
            string serviceInput = Console.ReadLine();

            string selectedService = GetSelectedService(serviceInput);
            if (selectedService == null) return;

            // Sélection du paramètre
            if (_excelManager.Settings.ContainsKey(selectedService))
            {
                var serviceSettings = _excelManager.Settings[selectedService];
                Console.WriteLine($"\nParamètres disponibles pour {selectedService}:");
                for (int i = 0; i < serviceSettings.Count; i++)
                {
                    Console.WriteLine($"{i + 1}. {serviceSettings[i]}");
                }

                Console.Write("Sélectionnez un paramètre (numéro ou nom): ");
                string settingInput = Console.ReadLine();

                string selectedSetting = GetSelectedSetting(serviceSettings, settingInput);
                if (selectedSetting == null) return;

                // Lecture du paramètre
                var result = await _deviceManager.ReadWriteSetting(selectedService, selectedSetting, "read", null);
            //  return $"Service: {service}, Paramètre: {setting}, Valeur écrite: {value.Value}";
                Console.WriteLine($"Résultat: Service: {selectedService}, Paramètre: {selectedSetting}, Valeur écrite: {result}");
            }
            else
            {
                Console.WriteLine("Service non trouvé dans les paramètres disponibles");
            }
        }

        /// <summary>
        /// Interface pour écrire un paramètre
        /// </summary>
        private async Task WriteSetting()
        {
            Console.WriteLine("\n=== ÉCRITURE D'UN PARAMÈTRE ===");

            Console.WriteLine("Services disponibles:");
            for (int i = 0; i < _deviceManager.AvailableServices.Count; i++)
            {
                Console.WriteLine($"{i + 1}. {_deviceManager.AvailableServices[i]}");
            }

            Console.Write("Sélectionnez un service: ");
            string serviceInput = Console.ReadLine();

            string selectedService = GetSelectedService(serviceInput);
            if (selectedService == null) return;

            if (_excelManager.Settings.ContainsKey(selectedService))
            {
                var serviceSettings = _excelManager.Settings[selectedService];
                Console.WriteLine($"\nParamètres disponibles pour {selectedService}:");
                for (int i = 0; i < serviceSettings.Count; i++)
                {
                    Console.WriteLine($"{i + 1}. {serviceSettings[i]}");
                }

                Console.Write("Sélectionnez un paramètre: ");
                string settingInput = Console.ReadLine();

                string selectedSetting = GetSelectedSetting(serviceSettings, settingInput);
                if (selectedSetting == null) return;

                Console.Write("Entrez la valeur à écrire: ");
                if (int.TryParse(Console.ReadLine(), out int value))
                {
                    var result = await _deviceManager.ReadWriteSetting(selectedService, selectedSetting, "write", value);
                    Console.WriteLine($"Résultat: Service: {selectedService}, Paramètre: {selectedSetting}, Valeur écrite: {result}");
                }
                else
                {
                    Console.WriteLine("Valeur invalide");
                }
            }
        }

        /// <summary>
        /// Interface pour lire un registre directement
        /// </summary>
        private async Task ReadRegisterDirect()
        {
            Console.WriteLine("\n=== LECTURE DIRECTE D'UN REGISTRE ===");

            Console.Write("Unit ID: ");
            if (!int.TryParse(Console.ReadLine(), out int unitId))
            {
                Console.WriteLine("Unit ID invalide");
                return;
            }

            Console.Write("Adresse du registre: ");
            if (!int.TryParse(Console.ReadLine(), out int register))
            {
                Console.WriteLine("Adresse invalide");
                return;
            }

            Console.Write("Nombre de registres (défaut: 1): ");
            string countInput = Console.ReadLine();
            int count = string.IsNullOrEmpty(countInput) ? 1 : int.Parse(countInput);

            try
            {
                var client = await _modbusManager.ConnectModbus(NetworkHelper.GetIpFromMac(Configuration.VICTRON_MAC), Configuration.MODBUS_PORT);
                if (client != null)
                {
                    var result = await _modbusManager.ReadRegisters(client, register, count, unitId);
                    Console.WriteLine($"Valeurs lues: [{string.Join(", ", result)}]");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erreur: {ex.Message}");
            }
        }


        // boucle pour lire les données à chaque seconde
        private async Task StartAutoPolling()
        {
            Console.WriteLine("Démarrage de la récupération automatique...");

                try
                {
                    var result = new Dictionary<string, Dictionary<string, object>>();

                    // Rafraîchir la liste des services
                    var availableServices = _deviceManager.AvailableServices;

                    foreach (var service in availableServices)
                    {
                        if (!_excelManager.Settings.ContainsKey(service)) continue;

                        var serviceSettings = _excelManager.Settings[service];
                        var settingValues = new Dictionary<string, object>();

                        foreach (var setting in serviceSettings)
                        {
                            try
                            {
                                var value = await _deviceManager.ReadWriteSetting(service, setting, "read", null);
                                settingValues[setting] = value;
                            }
                            catch (Exception ex)
                            {
                                settingValues[setting] = $"Erreur: {ex.Message}";
                            }
                        }

                        result[service] = settingValues;
                    }

                    // Afficher le JSON dans la console
                    json = JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
                    /*Console.WriteLine(json);

                    Console.WriteLine($"Données mises à jour à {DateTime.Now}");*/
                    //await Task.Delay(1000); // Attendre 1 seconde
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Erreur dans la boucle : {ex.Message}");
                    await Task.Delay(1000); // attendre même en cas d'erreur
                }
        }


        /// <summary>
        /// Obtient le service sélectionné par l'utilisateur
        /// </summary>
        private string GetSelectedService(string serviceInput)
        {
            if (int.TryParse(serviceInput, out int serviceIndex) && serviceIndex > 0 && serviceIndex <= _deviceManager.AvailableServices.Count)
            {
                return _deviceManager.AvailableServices[serviceIndex - 1];
            }
            else
            {
                return serviceInput;
            }
        }

        /// <summary>
        /// Obtient le paramètre sélectionné par l'utilisateur
        /// </summary>
        private string GetSelectedSetting(System.Collections.Generic.List<string> serviceSettings, string settingInput)
        {
            if (int.TryParse(settingInput, out int settingIndex) && settingIndex > 0 && settingIndex <= serviceSettings.Count)
            {
                return serviceSettings[settingIndex - 1];
            }
            else
            {
                return settingInput;
            }
        }
    }
}