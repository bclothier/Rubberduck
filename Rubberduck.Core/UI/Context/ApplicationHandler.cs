using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Windows;

namespace Rubberduck.UI.Context
{
    /// <remarks>
    /// Because we have a hosted WPF scenario, we need to manually run the <see cref="Application"/> so that we can leverage
    /// some of the features that usually comes with the object. This was based on the Dr. WPF article:
    /// http://drwpf.com/blog/2007/10/05/managing-application-resources-when-wpf-is-hosted/
    /// 
    /// Note that we already tried the last suggestion in the article which would avoid needing an application object at all but 
    /// in practice this caused more problems/headaches so having an application is actually simpler and enables us to solve any
    /// other problems in manners more akin to a traditional WPF application because we now have an application we can use. 
    /// 
    /// BUT!
    /// 
    /// A WPF application cannot be ever fully unloaded once it has been loaded. See the SO thread for discussion:
    /// https://stackoverflow.com/questions/52894181/wpf-cannot-close-application-instance-for-running-it-a-second-time
    /// 
    /// In light of this restriction and the fact that we don't want to futz with AppDomains, the best thing to do is to
    /// create the <see cref="Application"/> object if it doesn't exist and simply never call the <see cref="Application.Shutdown"/>
    /// method. This will work even when we unload &amp; reload the Rubberduck add-in as the object will remain alive and we can still
    /// customize it as we need.
    /// </remarks>
    public static class ApplicationHandler
    {
        public static void WpfStartup(CultureInfo cultureInfo, string theme)
        {
            if (Application.Current == null)
            {
                // create the WPF Application object
                // we do not want to run but simply provide the object to contain
                // our application-level settings & resources
                new Application();
            }

            ChangeTheme(theme);
            ChangeLanguage(cultureInfo);

            // Avoid prematurely releasing resources since we are hosted
            Application.Current.ShutdownMode = ShutdownMode.OnExplicitShutdown;
        }

        public static void ChangeTheme(string theme)
        {
            //TODO: If we ever get the theme to change sanely, we can uncomment the next line and delete the if block

            // Application.Current.Resources.MergedDictionaries.Clear();
            if(Application.Current.Resources.MergedDictionaries.Any())
            {
                return;
            }

            // merge in the resource dictionary
            Application.Current.Resources.MergedDictionaries.Add(
                Application.LoadComponent(
                    new Uri($"/Rubberduck.Core;component/UI/Styles/{theme}{(theme.ToLowerInvariant().EndsWith("theme") ? string.Empty : "theme")}.xaml",
                    UriKind.Relative)) as ResourceDictionary);
        }

        public static IEnumerable<string> GetThemes()
        {
            var strings = new List<string>();
            Assembly asm = Assembly.GetExecutingAssembly();
            Stream stream = asm.GetManifestResourceStream(asm.GetName().Name + ".g.resources");
            using (ResourceReader reader = new ResourceReader(stream))
            {
                foreach (DictionaryEntry entry in reader)
                {
                    var key = ((string)entry.Key).ToLowerInvariant();
                    if(key.EndsWith("theme.xaml"))
                    {
                        var pos = key.LastIndexOf("/") + 1;
                        strings.Add(key.Substring(pos, key.Length - pos)
                            .Replace("theme.xaml", string.Empty));
                    }
                }
            }

            return strings;
        }

        public static void ChangeLanguage(CultureInfo cultureInfo)
        {
            try
            {
                Application.Current.Dispatcher.Thread.CurrentUICulture = cultureInfo;
            }
            catch (CultureNotFoundException)
            {
            }
        }

        public static CultureInfo GetLanguage()
        {
            return Application.Current.Dispatcher.Thread.CurrentUICulture;
        }
    }
}