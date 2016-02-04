using System.Collections.Generic;

namespace Bananagrams2
{
    public class WebRole
    {
        public static Dictionary<string, BananagramsPlayer> bananagramsPlayers;
        public static Dictionary<int, Bananagrams> bananagramsGames;

        public static Dictionary<string, HanabiPlayer> hanabiPlayers;
        public static Dictionary<int, Hanabi> hanabiGames;

        public static void InitDB()
        {
            // TODO: Turn this into a DB or something...
            if (WebRole.bananagramsGames == null)
            {
                WebRole.bananagramsGames = new Dictionary<int, Bananagrams>();
            }
            if (WebRole.bananagramsPlayers == null)
            {
                WebRole.bananagramsPlayers = new Dictionary<string, BananagramsPlayer>();
            }

            if (WebRole.hanabiGames == null)
            {
                WebRole.hanabiGames = new Dictionary<int, Hanabi>();
            }
            if (WebRole.hanabiPlayers == null)
            {
                WebRole.hanabiPlayers = new Dictionary<string, HanabiPlayer>();
            }
        }
    }
}
