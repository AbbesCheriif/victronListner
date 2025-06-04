using System;
using System.Collections.Generic;
using System.Net;
using System.Text;
using System.Text.Json.Serialization;
using System.Text.Json;
using System.Threading.Tasks;
using System.Xml.Linq;
using Microsoft.Azure.Devices.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Renci.SshNet.Messages;
using System.Diagnostics;

namespace victronListner
{
    /// <summary>
    /// Classe principale de l'application avec intégration Azure IoT Hub
    /// </summary>
    internal class Program
    {
        // Gestionnaires principaux
        private static ExcelManager _excelManager;
        private static ModbusManager _modbusManager;
        private static SshManager _sshManager;
        private static DeviceManager _deviceManager;
        private static MenuManager _menuManager;

        // Configuration Azure IoT Hub
        private static DeviceClient deviceClient;
        private const string deviceConnectionString = "HostName=azureHubIot.azure-devices.net;DeviceId=victronTwin;SharedAccessKey=BPYOtr0LOwAlPiXiqOb465iYPFovowNLejEatNLlIwU=";

        static async Task Main(string[] args)
        {
            Console.WriteLine("=== Cerbo GX Console Application ===");
            try
            {
                // Initialisation
                await InitializeApplication();

                // Connexion à Azure IoT Hub
                Console.WriteLine("Connexion à Azure IoT Hub...");
                deviceClient = DeviceClient.CreateFromConnectionString(deviceConnectionString, Microsoft.Azure.Devices.Client.TransportType.Mqtt);
                Console.WriteLine("✅ Connexion à Azure IoT Hub établie");

                while (true)
                {
                    // 1. Déclarer et démarrer le Stopwatch
                    Stopwatch stopwatch = new Stopwatch();
                    stopwatch.Start(); // Démarre le chronomètre

                    // Menu principal - récupération des données
                    await _menuManager.readAll();
                    Console.WriteLine("📋 Données récupérées:");
                    Console.WriteLine(_menuManager.json);

                    // Traitement et envoi vers IoT Hub
                    Console.WriteLine("\n🚀 Traitement et envoi vers Azure IoT Hub...");
                    await ProcessAndSendToIoTHub(_menuManager.json);

                    Console.WriteLine("\n✅ Cycle terminé. Attente avant le prochain cycle...");
                    // 2. Arrêter le Stopwatch
                    stopwatch.Stop(); // Arrête le chronomètre

                    // 3. Récupérer et afficher le temps écoulé
                    // Temps total écoulé en millisecondes
                    long milliseconds = stopwatch.ElapsedMilliseconds;
                    Console.WriteLine($"\nTemps d'exécution du bloc de code : {milliseconds} ms");


                    await Task.Delay(120000); // Attendre 120 secondes avant le prochain cycle
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erreur lors de l'exécution: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
            }
            finally
            {
                // Nettoyage des ressources
                if (deviceClient != null)
                {
                    await deviceClient.CloseAsync();
                    deviceClient.Dispose();
                }
            }

            Console.WriteLine("Appuyez sur une touche pour quitter...");
            Console.ReadKey();
        }

        /// <summary>
        /// Initialise tous les composants de l'application
        /// </summary>
        private static async Task InitializeApplication()
        {
            Console.WriteLine("Initialisation de l'application...");

            // Création des gestionnaires
            _excelManager = new ExcelManager();
            _modbusManager = new ModbusManager();
            _sshManager = new SshManager();
            _deviceManager = new DeviceManager(_sshManager, _excelManager, _modbusManager);
            _menuManager = new MenuManager(_excelManager, _modbusManager, _deviceManager);

            // Chargement des données Excel
            _excelManager.LoadExcelData();

            // Récupération des appareils connectés
            await _deviceManager.LoadConnectedDevices();

            Console.WriteLine("✅ Initialisation terminée.");
        }

        /// <summary>
        /// Traite le JSON et envoie chaque paramètre vers Azure IoT Hub
        /// </summary>
        /// <param name="jsonString">Le JSON contenant les services et leurs paramètres</param>
        private static async Task ProcessAndSendToIoTHub(string jsonString)
        {
            try
            {
                if (string.IsNullOrEmpty(jsonString))
                {
                    Console.WriteLine("⚠️ Aucune donnée JSON à traiter");
                    return;
                }

                // Parser le JSON
                var jsonData = JObject.Parse(jsonString);
                int totalSent = 0;

                // Parcourir chaque service
                foreach (var serviceProperty in jsonData.Properties())
                {
                    string serviceName = serviceProperty.Name;
                    var serviceData = serviceProperty.Value as JObject;

                    if (serviceData == null) continue;

                    Console.WriteLine($"\n--- Traitement du service: {serviceName} ---");

                    int settingIndex = 1;
                    // Parcourir chaque setting du service
                    foreach (var settingProperty in serviceData.Properties())
                    {
                        string settingName = settingProperty.Name;
                        var settingValue = settingProperty.Value;

                        // Nettoyer le nom du service (enlever com.victronenergy. si présent)
                        var cleanedService = serviceName.Replace("com.victronenergy.", "").Replace(".", "_");
                        var settingId = $"Setting_{cleanedService}_{settingIndex}";

                        // Créer l'objet de données à envoyer
                        var settingData = new
                        {
                            Service = $"Service_{serviceName.Replace("com.victronenergy.", "")}",
                            SettingId = settingId,
                            //SettingName = settingName,
                            Value = settingValue,
                            Timestamp = DateTime.UtcNow
                        };

                        // Convertir en JSON et envoyer
                        string messageString = JsonConvert.SerializeObject(settingData);
                        var message = new Microsoft.Azure.Devices.Client.Message(Encoding.UTF8.GetBytes(messageString)); 

                        Console.WriteLine($"📤 Envoi: {settingId} = {settingValue}");

                        try
                        {
                            await deviceClient.SendEventAsync(message);
                            Console.WriteLine($"✅ Envoyé avec succès: {settingId}");
                            totalSent++;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"❌ Erreur lors de l'envoi de {settingId}: {ex.Message}");
                        }

                        settingIndex++;

                        // Petit délai pour éviter de surcharger IoT Hub
                        await Task.Delay(100);
                    }
                }

                Console.WriteLine($"\n📊 Résumé: {totalSent} paramètres envoyés vers Azure IoT Hub");
            }
            catch (Newtonsoft.Json.JsonException ex)
            {
                Console.WriteLine($"❌ Erreur lors du parsing JSON: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Erreur lors du traitement: {ex.Message}");
            }
        }
    }
}