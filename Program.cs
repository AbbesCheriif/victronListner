using System;
using System.Threading.Tasks;

namespace victronListner
{
    /// <summary>
    /// Classe principale de l'application
    /// </summary>
    internal class Program
    {
        // Gestionnaires principaux
        private static ExcelManager _excelManager;
        private static ModbusManager _modbusManager;
        private static SshManager _sshManager;
        private static DeviceManager _deviceManager;
        private static MenuManager _menuManager;

        static async Task Main(string[] args)
        {
            Console.WriteLine("=== Cerbo GX Console Application ===");

            try
            {
                // Initialisation
                await InitializeApplication();

                // Menu principal
                await _menuManager.ShowMainMenu();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erreur lors de l'initialisation: {ex.Message}");
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

            Console.WriteLine("Initialisation terminée.");
        }
    }
}