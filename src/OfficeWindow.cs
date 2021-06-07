namespace OfficeRPC {
    public class OfficeWindow {
        public string DocumentName { get; }
        public OfficeWindowType WindowType { get; }

        public OfficeWindow(string documentName, OfficeWindowType type) {
            DocumentName = documentName;
            WindowType = type;
        }

        public override bool Equals(object? obj) {
            if (obj == null) return false;
            if (obj as OfficeWindow == null) return false;
            return ((OfficeWindow)obj).WindowType == WindowType && ((OfficeWindow)obj).DocumentName == DocumentName;
        }
    }
}