using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Bananagrams2
{
    public class BananagramsPlayer
    {
        public BananagramsPlayer(Bananagrams game, string name)
        {
            this.game = game;
            this.name = name;
            words = new List<string>();
        }

        public bool Equals(BananagramsPlayer p)
        {
            if (p == null)
            {
                return false;
            }

            return (name == p.name);
        }

        public void Reset()
        {
            words.Clear();
        }

        public string name { get; private set; }

        //public int score { get; set; }
        public List<string> words { get; set; }

        [Newtonsoft.Json.JsonIgnore]
        public Bananagrams game { get; set; }
    }
}