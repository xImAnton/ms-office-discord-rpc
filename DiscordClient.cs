using DiscordRPC;
using System;
using System.Collections.Generic;
using DiscordRPC.Logging;

namespace office_rpc {
    public class DiscordClient {
        private static readonly Dictionary<OfficeWindowType, string> WINDOW_NAMES = new Dictionary<OfficeWindowType, string>() {
            {OfficeWindowType.WORD, "Word"},
            {OfficeWindowType.EXCEL, "Excel"}
        };

        private static readonly Dictionary<OfficeWindowType, string> LARGE_IMAGE_KEYS = new Dictionary<OfficeWindowType, string>() {
            {OfficeWindowType.WORD, "word"},
            {OfficeWindowType.EXCEL, "excel"}
        };
        
        private DiscordRpcClient Client;
        private OfficeWindow CurrentOfficeWindow;
        private Timestamps TypeSwitchTime = Timestamps.Now;

        public void Init() {
            Client = new DiscordRpcClient("851465420980092939");
            Client.Logger = new ConsoleLogger(LogLevel.None);
            Client.OnReady += (_, e) => {
                Console.WriteLine("ready. user: {0}", e.User.Username);
            };
            
            Client.Initialize();
        }

        public void UpdatePresence(OfficeWindow officeWindow) {
            if (CurrentOfficeWindow != null && CurrentOfficeWindow.WindowType != officeWindow.WindowType) {
                TypeSwitchTime = Timestamps.Now;
            }
            CurrentOfficeWindow = officeWindow;
            RichPresence presence = new RichPresence {
                Details = "Using " + WINDOW_NAMES[officeWindow.WindowType],
                State = officeWindow.DocumentName != null ? "Editing " + officeWindow.DocumentName : "Idle",
                Assets = new Assets {
                    LargeImageKey = LARGE_IMAGE_KEYS[officeWindow.WindowType],
                    LargeImageText = "Microsoft " + WINDOW_NAMES[officeWindow.WindowType],
                    SmallImageKey = "pp",
                    SmallImageText = "Made by xImAnton_#2013"
                },
                Timestamps = TypeSwitchTime
            };
            Client.SetPresence(presence);
        }

        public OfficeWindow GetOfficeWindow() {
            return CurrentOfficeWindow;
        }

        public void Close() {
            Client.Deinitialize();
        }

        public void Reset() {
            Client.ClearPresence();
        }
    }
}