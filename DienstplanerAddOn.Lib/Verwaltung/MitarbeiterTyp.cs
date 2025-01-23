using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DienstplanerAddOn.Lib.Verwaltung
{
    public class MitarbeiterTyp
    {
        public string Name { get; set; }
        public int BenötigtProSchicht { get; set; }
        public int BenötigtProTag { get; set; }
        public int MaxSchichtenAmStück { get; set; }
        public int FreiNachBlock { get; set; }
        public int MaxSchichtenProMonat { get; set; }
        public int MinSchichtenProMonat { get; set; }

        public MitarbeiterTyp(string name, int benötigtProSchicht, int benötigtProTag, 
            int maxSchichtenAmStück, int freiNachBlock, int maxSchichtenProMonat, int minSchichtenProMonat)
        {
            Name = name;
            BenötigtProSchicht = benötigtProSchicht;
            BenötigtProTag = benötigtProTag;
            MaxSchichtenAmStück = maxSchichtenAmStück;
            FreiNachBlock = freiNachBlock;
            MaxSchichtenProMonat = maxSchichtenProMonat;
            MinSchichtenProMonat = minSchichtenProMonat;
        }

        public override string ToString()
        {
            return $"{nameof(Name)}:{Name}\n{nameof(BenötigtProTag)}:{BenötigtProTag}\n{nameof(BenötigtProSchicht)}:{BenötigtProSchicht}";
        }



    }
}
