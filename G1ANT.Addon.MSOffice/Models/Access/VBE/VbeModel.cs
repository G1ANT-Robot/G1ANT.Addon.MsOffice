namespace G1ANT.Addon.MSOffice.Models.Access.VBE
{
    //internal class VbeEventCollectionModel : List<VbeEventModel>
    //{
    //    public VbeEventCollectionModel(Events events)
    //    {
    //        AddRange(events.ReferencesEvents.Cast<Event>().Select(e => new VbeAddinModel(e)));
    //    }
    //}

    internal class VbeModel
    {
        public VbeWindowCollectionModel Windows { get; }
        public string Version { get; }
        public VbeProjectCollectionModel Projects { get; }
        public VbeProjectModel ActiveVBProject { get; }
        public VbeAddinCollectionModel Addins { get; }
        public VbeWindowModel MainWindow { get; }
        //public VbeEventCollectionModel Events { get; }

        public VbeModel(Microsoft.Vbe.Interop.VBE vbe)
        {
            Version = vbe.Version;
            Windows = new VbeWindowCollectionModel(vbe.Windows);
            Projects = new VbeProjectCollectionModel(vbe.VBProjects);
            ActiveVBProject = new VbeProjectModel(vbe.ActiveVBProject);
            Addins = new VbeAddinCollectionModel(vbe.Addins);
            MainWindow = new VbeWindowModel(vbe.MainWindow);
            //Events = new VbeEventCollectionModel(vbe.Events);
        }
    }
}
