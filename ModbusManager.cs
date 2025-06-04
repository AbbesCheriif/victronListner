using System;
using System.Threading.Tasks;
using NModbus;

namespace victronListner
{
    /// <summary>
    /// Gestionnaire pour les opérations Modbus
    /// </summary>
    public class ModbusManager
    {
        private IModbusMaster _modbusClient;

        /// <summary>
        /// Établit une connexion Modbus
        /// </summary>
        public async Task<IModbusMaster> ConnectModbus(string ip, int port)
        {
            try
            {
                var tcpClient = new System.Net.Sockets.TcpClient();
                await tcpClient.ConnectAsync(ip, port);

                var factory = new ModbusFactory();
                var master = factory.CreateMaster(tcpClient);

                Console.WriteLine($"Connecté au serveur Modbus {ip}:{port}");
                _modbusClient = master;
                return master;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erreur de connexion Modbus: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Lit des registres Modbus
        /// </summary>
        public async Task<ushort[]> ReadRegisters(IModbusMaster client, int register, int count, int unit)
        {
            try
            {
                var result = await Task.Run(() => client.ReadHoldingRegisters((byte)unit, (ushort)register, (ushort)count));
                Console.WriteLine($"Registre {register}: [{string.Join(", ", result)}]");
                return result;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erreur: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Écrit dans un registre Modbus
        /// </summary>
        public async Task<bool> WriteRegister(IModbusMaster client, int register, int value, int unit)
        {
            try
            {
                await Task.Run(() => client.WriteSingleRegister((byte)unit, (ushort)register, (ushort)value));
                Console.WriteLine($"Valeur {value} écrite avec succès dans le registre {register}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erreur: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Obtient le client Modbus actuel
        /// </summary>
        public IModbusMaster GetClient()
        {
            return _modbusClient;
        }

        /// <summary>
        /// Définit le client Modbus
        /// </summary>
        public void SetClient(IModbusMaster client)
        {
            _modbusClient = client;
        }
    }
}