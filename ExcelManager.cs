using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace victronListner
{
    /// <summary>
    /// Gestionnaire pour les opérations Excel
    /// </summary>
    public class ExcelManager
    {
        public DataTable ExcelData { get; private set; }
        public DataTable UnitIdMapping { get; private set; }
        public List<string> ServiceNames { get; private set; }
        public Dictionary<string, List<string>> Settings { get; private set; }

        public ExcelManager()
        {
            // Configuration EPPlus pour éviter les licences
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }

        /// <summary>
        /// Charge les données depuis le fichier Excel
        /// </summary>
        public void LoadExcelData()
        {
            try
            {
                if (!File.Exists(Configuration.EXCEL_FILE_PATH))
                {
                    throw new FileNotFoundException($"Le fichier Excel {Configuration.EXCEL_FILE_PATH} est introuvable.");
                }

                using (var package = new ExcelPackage(new FileInfo(Configuration.EXCEL_FILE_PATH)))
                {
                    Console.WriteLine($"Nombre de feuilles dans le fichier: {package.Workbook.Worksheets.Count}");

                    // Afficher les noms des feuilles pour diagnostic
                    for (int i = 0; i < package.Workbook.Worksheets.Count; i++)
                    {
                        Console.WriteLine($"Feuille {i}: {package.Workbook.Worksheets[i].Name}");
                    }

                    // Vérifier qu'il y a au moins une feuille
                    if (package.Workbook.Worksheets.Count == 0)
                    {
                        throw new Exception("Le fichier Excel ne contient aucune feuille.");
                    }

                    // Feuille principale (première feuille disponible)
                    var worksheet = package.Workbook.Worksheets[0];
                    Console.WriteLine($"Traitement de la feuille principale: {worksheet.Name}");
                    ExcelData = WorksheetToDataTable(worksheet, true, Configuration.EXCEL_DATA_START_ROW);

                    // Recherche de la feuille Unit ID mapping
                    var unitMappingWorksheet = package.Workbook.Worksheets
                        .FirstOrDefault(ws => ws.Name.Contains("Unit ID") || ws.Name.Contains("mapping"));

                    if (unitMappingWorksheet != null)
                    {
                        Console.WriteLine($"Feuille Unit ID mapping trouvée: {unitMappingWorksheet.Name}");
                        UnitIdMapping = WorksheetToDataTable(unitMappingWorksheet, true, Configuration.EXCEL_MAPPING_START_ROW);
                    }
                    else
                    {
                        Console.WriteLine("Attention: Feuille Unit ID mapping non trouvée. Création d'une table vide.");
                        UnitIdMapping = new DataTable();
                        UnitIdMapping.Columns.Add("/DeviceInstance");
                        UnitIdMapping.Columns.Add("Unit ID");
                    }
                }

                ProcessExcelData();
                Console.WriteLine($"Données Excel chargées: {ServiceNames.Count} services trouvés");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erreur lors du chargement Excel: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Convertit une feuille Excel en DataTable
        /// </summary>
        private DataTable WorksheetToDataTable(OfficeOpenXml.ExcelWorksheet worksheet, bool hasHeader, int startRow = 1)
        {
            DataTable dt = new DataTable();

            // Vérifier que la feuille n'est pas vide
            if (worksheet.Dimension == null)
            {
                Console.WriteLine("Attention: La feuille est vide ou ne contient pas de données.");
                return dt;
            }

            Console.WriteLine($"Dimensions de la feuille: {worksheet.Dimension.Address}");
            Console.WriteLine($"Lignes: {worksheet.Dimension.Start.Row} à {worksheet.Dimension.End.Row}");
            Console.WriteLine($"Colonnes: {worksheet.Dimension.Start.Column} à {worksheet.Dimension.End.Column}");

            // Ajout des colonnes
            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                string columnName = hasHeader ?
                    worksheet.Cells[startRow, col].Value?.ToString()?.Trim() ?? $"Column{col}" :
                    $"Column{col}";
                dt.Columns.Add(columnName);
            }

            // Ajout des données
            int dataStartRow = hasHeader ? startRow + 1 : startRow;
            for (int row = dataStartRow; row <= worksheet.Dimension.End.Row; row++)
            {
                DataRow dr = dt.NewRow();
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    dr[col - 1] = worksheet.Cells[row, col].Value?.ToString()?.Trim() ?? "";
                }
                dt.Rows.Add(dr);
            }

            Console.WriteLine($"DataTable créée avec {dt.Rows.Count} lignes et {dt.Columns.Count} colonnes");
            return dt;
        }

        /// <summary>
        /// Traite les données Excel pour extraire les services et paramètres
        /// </summary>
        private void ProcessExcelData()
        {
            // Extraction des noms de services
            ServiceNames = ExcelData.AsEnumerable()
                .Select(row => row.Field<string>("dbus-service-name"))
                .Where(name => !string.IsNullOrEmpty(name))
                .Distinct()
                .ToList();

            // Création du dictionnaire des paramètres
            Settings = new Dictionary<string, List<string>>();
            foreach (string service in ServiceNames)
            {
                var serviceSettings = ExcelData.AsEnumerable()
                    .Where(row => row.Field<string>("dbus-service-name") == service)
                    .Select(row => row.Field<string>("description"))
                    .Where(desc => !string.IsNullOrEmpty(desc))
                    .ToList();

                Settings[service] = serviceSettings;
            }
        }

        /// <summary>
        /// Obtient les informations d'un paramètre
        /// </summary>
        public DataRow GetParameterInfo(string service, string setting)
        {
            return ExcelData.AsEnumerable()
                .FirstOrDefault(row => row.Field<string>("dbus-service-name") == service &&
                                      row.Field<string>("description") == setting);
        }

        /// <summary>
        /// Obtient l'Unit ID pour une instance donnée
        /// </summary>
        public int? GetUnitId(int instanceId)
        {
            var unitIdRow = UnitIdMapping.AsEnumerable()
                .FirstOrDefault(row => row.Field<string>("/DeviceInstance") == instanceId.ToString());

            if (unitIdRow != null && int.TryParse(unitIdRow.Field<string>("Unit ID"), out int unitId))
            {
                return unitId;
            }

            return null;
        }
    }
}