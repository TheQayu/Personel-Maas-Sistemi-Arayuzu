using System;
using System.Reflection;
using System.Threading;
using BenchmarkDotNet.Attributes;

namespace BenchmarkSuite1
{
    public class PuantajViewBenchmark
    {
        private MethodInfo loadMethod;
        [GlobalSetup]
        public void Setup()
        {
            loadMethod = typeof(denemelikimid.Form1).GetMethod("LoadPuantajView", BindingFlags.Instance | BindingFlags.NonPublic);
        }

        [Benchmark]
        public void LoadPuantajView()
        {
            Exception exception = null;
            var thread = new Thread(() =>
            {
                try
                {
                    using (var form = new denemelikimid.Form1())
                    {
                        loadMethod.Invoke(form, null);
                    }
                }
                catch (Exception ex)
                {
                    exception = ex;
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
            if (exception != null)
            {
                throw new InvalidOperationException("Benchmark failed.", exception);
            }
        }
    }
}