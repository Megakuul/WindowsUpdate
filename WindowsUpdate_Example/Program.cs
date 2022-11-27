using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using WUApiLib;

namespace WindowsUpdate_Example
{
    class Program
    {
        private static UpdateSession wuSession { get; set; }



        static void Main(string[] args)
        {
            Console.WriteLine("Start windows updates (press enter)");
            Console.ReadLine();
            Console.WriteLine("Searching for updates...");

            UpdateCollection updates = searchUpdates().Updates;

            if (updates.Count == 0) {
                Console.WriteLine("All updates are already installed");
                Console.ReadLine();
                return;
            }    

            Console.WriteLine($"Downloading the updates ended with code: {downloadUpdates(updates)}");

            Console.WriteLine($"Installing the updates ended up with code: {installUpdates(updates)}");

            Console.WriteLine("View all installed updates (press any key)");

            foreach(IUpdate update in updates)
            {
                if (!update.IsInstalled)
                {
                    Console.WriteLine($"Installed: {update.Title}");
                } else
                {
                    Console.WriteLine($"Not installed: {update.Title}");
                }
            }

            while (true);
        }

        private static ISearchResult searchUpdates()
        {
            wuSession = new UpdateSession();
            IUpdateSearcher wuSearcher = wuSession.CreateUpdateSearcher();

            return wuSearcher.Search("IsInstalled=0");
        }

        private static int downloadUpdates(UpdateCollection updates)
        {
            UpdateDownloader wuDownloader = wuSession.CreateUpdateDownloader();

            wuDownloader.Updates = new UpdateCollection();

            foreach (IUpdate update in updates)
            {
                Console.WriteLine($"Update found: {update.Title}");
                wuDownloader.Updates.Add(update);
            }
            try {
                wuDownloader.Download();
                return 0;
            } catch (Exception ex) {
                Console.WriteLine(ex.Message);
                return 1;
            }
        }

        private static int installUpdates(UpdateCollection updates)
        {
            IUpdateInstaller wuInstaller = wuSession.CreateUpdateInstaller();

            wuInstaller.Updates = new UpdateCollection();

            wuInstaller.IsForced = true;

            foreach (IUpdate update in updates)
            {
                if (update.IsDownloaded)
                {
                    wuInstaller.Updates.Add(update);
                    Console.WriteLine($"Installing update: {update.Title}");
                } 
            }

            try {
                wuInstaller.Install();
                return 0;
            }
            catch (Exception ex) {
                Console.WriteLine(ex.Message);
                return 1;
                
            }
        }
    }
}
