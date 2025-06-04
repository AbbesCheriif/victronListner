namespace victronListner
{
    /// <summary>
    /// Classe contenant toutes les constantes de configuration
    /// </summary>
    public static class Configuration
    {
        // Configuration Modbus
        public static string MODBUS_IP = "192.168.1.127";
        public const int MODBUS_PORT = 502;

        // Configuration Victron
        public const string VICTRON_MAC = "c0-61-9a-b3-15-47";

        // Configuration SSH
        public const string SSH_HOST = "192.168.1.127";
        public const string SSH_USERNAME = "root";
        public const string SSH_PASSWORD = "00000000";

        // Configuration Excel
        public const string EXCEL_FILE_PATH = "CCGX-Modbus-TCP-register-list_editable.xlsx";
        public const int EXCEL_DATA_START_ROW = 2;
        public const int EXCEL_MAPPING_START_ROW = 1;
    }
}