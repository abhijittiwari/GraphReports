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

public static class Mapping
{
    private static Dictionary<string, string> productSkuMapping = new Dictionary<string, string>();

    static Mapping()
    {
        LoadMapping("SKUMapping.csv");
    }

    public static void LoadMapping(string csvFilePath)
    {
        csvFilePath = @"C:\dev\GraphExplorer\GraphReports\SKUMapping.csv";
        using (var reader = new StreamReader(csvFilePath))
        using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
        {
            var records = csv.GetRecords<SkuRecord>();
            foreach (var record in records)
            {
                productSkuMapping[record.ProductName] = record.SKU_ID;
            }
        }
    }

    public static string GetProductNameBySkuId(string skuId)
    {
        return productSkuMapping.FirstOrDefault(x => x.Value == skuId).Key;
    }

    private class SkuRecord
    {
        public string ProductName { get; set; }
        public string SKU_ID { get; set; }
    }
}
