using CsvHelper;
using CsvHelper.Configuration;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.Security;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.Windows.Forms;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace GraphReports
{
    public static class ServicePlanMapping
    {
        private static Dictionary<string, string> ServicePlanMap = new Dictionary<string, string>();

        static ServicePlanMapping()
        {
            LoadMapping("SKUMapping.csv");
        }
        public static void LoadMapping(string csvFilePath)
        {
            csvFilePath = @"C:\dev\GraphExplorer\GraphReports\ServicePlanMapping.csv";
            using (var reader = new StreamReader(csvFilePath))
            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                var records = csv.GetRecords<SkuRecord>();
                foreach (var record in records)
                {
                    ServicePlanMap[record.ServicePlanName] = record.ServicePlanID;
                }
            }
        }
        public static string GetServicePlanById(string skuId)
        {
            return ServicePlanMap.FirstOrDefault(x => x.Value == skuId).Key;
        }

        private class SkuRecord
        {
            public string ServicePlanName { get; set; }
            public string ServicePlanID { get; set; }
        }


    }
}
