namespace MsCrmTools.Translator.AppCode
{
    internal class Language
    {
        public Language(int lcid, string name)
        {
            Lcid = lcid;
            Name = name;
        }

        public int Lcid { get; }
        public string Name { get; }

        public override string ToString()
        {
            return Name;
        }
    }
}