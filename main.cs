﻿using System.Threading;

namespace office_rpc {
    class Program {
        static void Main(string[] args) {
            var client = new DiscordClient();
            client.Init();

            while (true) {
                OfficeWindow officeWindow = OfficeManager.GetWindow();
                
                // update presence when the window has changed
                if (officeWindow != null && !officeWindow.Equals(client.GetOfficeWindow())) {
                    client.UpdatePresence(officeWindow);
                }
                
                // when a presence is set and now window is found, clear it
                if (officeWindow == null && client.GetOfficeWindow() != null) {
                    client.Reset();
                }
                
                // test again every 2 seconds
                Thread.Sleep(2000);
            }
        }
    }
}