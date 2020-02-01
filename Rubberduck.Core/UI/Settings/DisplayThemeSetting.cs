using System.Xml.Serialization;
using Rubberduck.Resources;

namespace Rubberduck.Settings
{
    [XmlType(AnonymousType = true)]
    public class DisplayThemeSetting
    {
        [XmlAttribute]
        public string Code { get; set; }

        /// <Summary>
        /// Default constructor required for XML serialization.
        /// </Summary>
        public DisplayThemeSetting()
        {
        }

        public DisplayThemeSetting(string code)
        {
            Code = code;

            var resource = "Theme_" + Code;
            try
            {
                _name = RubberduckUI.ResourceManager.GetString(resource);
            }
            catch
            {
                _name = null;
            }

            IsDefined = !string.IsNullOrWhiteSpace(_name);
        }

        private string _name;
        [XmlIgnore]
        public string Name => string.IsNullOrWhiteSpace(_name) ? Code : _name;

        [XmlIgnore]
        public bool IsDefined { get; }

        public override bool Equals(object obj)
        {
            return obj is DisplayThemeSetting other && Code.Equals(other.Code);
        }

        public override int GetHashCode()
        {
            return Code.GetHashCode();
        }
    }
}
