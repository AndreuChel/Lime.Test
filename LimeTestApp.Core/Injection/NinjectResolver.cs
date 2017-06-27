using Ninject;
using Ninject.Modules;
using Ninject.Syntax;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LimeTestApp.Core.Injection
{
    /// <summary>
    /// IoC-контейнер
    /// </summary>
    public static class NinjectResolver
    {
        private static readonly IKernel kernel;

        static NinjectResolver()
        {
            kernel = new StandardKernel();
        }

        public static T Get<T>()
        {
            return (T)Get(typeof(T));
        }

        public static object Get(Type type)
        {
            return kernel.Get(type);
        }

        public static T Get<T>(string name)
        {
            return (T)Get(typeof(T), name);
        }

        public static object Get(Type type, string name)
        {
            return kernel.Get(type, name);
        }

        public static IBindingToSyntax<T1> Bind<T1>()
        {
            return kernel.Bind<T1>();
        }

        public static IBindingWhenInNamedWithOrOnSyntax<T2> Bind<T1, T2>() where T2 : T1
        {
            return kernel.Bind<T1>().To<T2>();
        }

        public static T LoadModule<T>() where T : INinjectModule, new()
        {
            kernel.Load<T>();
            return kernel.GetModules().OfType<T>().Single();
        }

        public static void LoadModule(INinjectModule module)
        {
            kernel.Load(module);
        }
    }
}
