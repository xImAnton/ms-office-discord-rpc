namespace OfficeRPC {
    public class OfficeWindow {
        public string DocumentName { get; }
        public OfficeWindowType WindowType { get; }

        public OfficeWindow(string documentName, OfficeWindowType type) {
            DocumentName = documentName;
            WindowType = type;
        }
    }
}