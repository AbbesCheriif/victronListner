using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace victronListner
{
    /// <summary>
    /// Gestionnaire pour les appareils connectés
    /// </summary>
    public class DeviceManager
    {
        private readonly SshManager _sshManager;
        private readonly ExcelManager _excelManager;
        private readonly ModbusManager _modbusManager;

        public Dictionary<string, int> AvailableServicesWithInstanceIds { get; private set; }
        public List<string> AvailableServices { get; private set; }

        public DeviceManager(SshManager sshManager, ExcelManager excelManager, ModbusManager modbusManager)
        {
            _sshManager = sshManager;
            _excelManager = excelManager;
            _modbusManager = modbusManager;
        }

        /// <summary>
        /// Charge la liste des appareils connectés
        /// </summary>
        public async Task LoadConnectedDevices()
        {
            try
            {
                AvailableServicesWithInstanceIds = await _sshManager.GetConnectedDevices();
                AvailableServices = AvailableServicesWithInstanceIds.Keys
                    .Select(NetworkHelper.GetGenericService)
                    .Distinct()
                    .ToList();
                AvailableServices = new List<string> { "com.victronenergy.solarcharger", "com.victronenergy.system", "com.victronenergy.vebus" };
                AvailableServicesWithInstanceIds = new Dictionary<string, int>
                {
                    { "com.victronenergy.solarcharger", 279 },
                    { "com.victronenergy.system", 0 },
                    { "com.victronenergy.vebus", 276 }
                };

                Console.WriteLine($"Appareils connectés trouvés: {AvailableServices.Count}");
                Console.WriteLine("######Contenu du dictionnaire :");
                foreach (KeyValuePair<string, int> paire in AvailableServicesWithInstanceIds)
                {
                    Console.WriteLine($"Clé : {paire.Key}, Valeur : {paire.Value}");
                }

                Console.WriteLine("######\n---");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erreur lors de la récupération des appareils: {ex.Message}");
                AvailableServices = new List<string>();
            }
        }

        /// <summary>
        /// Lit ou écrit un paramètre d'appareil
        /// </summary>
        public async Task<string> ReadWriteSetting(string service, string setting, string operation, int? value)
        {
            try
            {
                // Connexion Modbus
                var modbusClient = _modbusManager.GetClient();
                if (modbusClient == null)
                {
                    string ip = NetworkHelper.GetIpFromMac(Configuration.VICTRON_MAC);
                    modbusClient = await _modbusManager.ConnectModbus(ip, Configuration.MODBUS_PORT);
                }

                if (modbusClient == null)
                    throw new Exception("Impossible de se connecter au Cerbo GX");

                // Récupération des informations du paramètre
                if (!AvailableServicesWithInstanceIds.ContainsKey(service))
                    throw new Exception($"Service {service} non connecté");

                int instanceId = AvailableServicesWithInstanceIds[service];

                // Récupération de l'Unit ID depuis le mapping
                int? unitId = _excelManager.GetUnitId(instanceId);
                if (!unitId.HasValue)
                    throw new Exception($"Unit ID non trouvé pour l'instance {instanceId}");

                // Récupération de l'adresse du registre
                var settingRow = _excelManager.GetParameterInfo(service, setting);
                if (settingRow == null)
                    throw new Exception($"Paramètre {setting} non trouvé pour le service {service}");

                int register = int.Parse(settingRow.Field<string>("Address"));

                // Exécution de l'opération
                if (operation == "read")
                {
                    var result = await _modbusManager.ReadRegisters(modbusClient, register, 1, unitId.Value);
                    return $"{result[0]}";
                }
                else if (operation == "write" && value.HasValue)
                {
                    await _modbusManager.WriteRegister(modbusClient, register, value.Value, unitId.Value);
                    return $"{value.Value}";
                }
                else
                {
                    throw new Exception("Opération non valide");
                }
            }
            catch (Exception ex)
            {
                return $"Erreur: {ex.Message}";
            }
        }
    }
}