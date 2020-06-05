using System;
using BenchmarkDotNet.Running;

namespace IEnumerable.ExportExtension.BenchmarkTest
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Begin!");
            BenchmarkRunner.Run<IEnumerableExtensionTest>();
            Console.WriteLine("Done!");
            Console.ReadKey();
        }
    }
}
