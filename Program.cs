using Microsoft.Extensions.Configuration;
using System;
using System.IO;
using ConsoleApp1.Apps;
using ConsoleApp1.Utils;
using ClassLibrary1.Utils.Persistence;

namespace ConsoleApp1
{
     class Program
    {
        static void Main(string[] args)
        {
            //Console.WriteLine("Hello World!");
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            ConfigurationManager.Build(new string[] {  "appsettings.json", "appsettings.development.json" },Directory.GetCurrentDirectory());
            OA.newStartOAHelper();
            //OA.StartOAHelper();
        }
    }
}
