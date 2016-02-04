using System;
using System.Threading.Tasks;

// Add reference
using Microsoft.AspNet.SignalR;

namespace Bananagrams2
{
    // Derive from Hub, which is a server-side class
    // and a client side proxy.
    public class HanabiHub : Hub
    {
        public Hanabi Create(string gameType)
        {
            WebRole.InitDB();
            Random r = new Random();
            int gameNumber = r.Next(1, 10000);

            Hanabi game = new Hanabi(gameNumber, gameType, this);
            WebRole.hanabiGames.Add(gameNumber, game);

            Groups.Add(Context.ConnectionId, gameNumber.ToString());

            return game;
        }
        public Hanabi Join(string gameString)
        {
            WebRole.InitDB();
            Hanabi game;

            int gameNumber = -1;

            try
            {
                gameNumber = int.Parse(gameString);
            }
            catch
            {
                gameNumber = -1;
            }

            if (WebRole.hanabiGames.ContainsKey(gameNumber))
            {
                game = WebRole.hanabiGames[gameNumber];
            }
            else
            {
                game = Create("medium");
                gameNumber = game.gameNumber;
            }

            Groups.Add(Context.ConnectionId, gameNumber.ToString());
            return game;
        }

        public int JoinAsPlayer(string gameNumberString, string playerName)
        {
            int gameNumber = int.Parse(gameNumberString);
            Hanabi game = WebRole.hanabiGames[gameNumber];
            HanabiPlayer player = null;
            int playerNumber = -1;

            for (int i = 0; i < game.players.Count; i++)
            {
                if (game.players[i].name == playerName)
                {
                    player = game.players[i];
                    playerNumber = i;
                }
            }

            if (player == null)
            {
                player = new HanabiPlayer(game, playerName);
                WebRole.hanabiPlayers.Add(Context.ConnectionId, player);
                game.AddPlayer(player);
                playerNumber = game.players.Count - 1;
            }
            else
            {
                WebRole.hanabiPlayers.Add(Context.ConnectionId, player);
            }

            return playerNumber;
        }

        public void Reset()
        {
            HanabiPlayer p = WebRole.hanabiPlayers[Context.ConnectionId];
            Hanabi game = p.game;
            game.Reset("");
            RequestGameUpdate();
        }

        public void RequestGameUpdate()
        {
            Hanabi game = WebRole.hanabiPlayers[Context.ConnectionId].game;
            Clients.Group(game.gameNumber.ToString()).broadcastGame(game);
        }

        /*
        public override Task OnDisconnected()
        {
            // TODO: fix this up
            //Groups.Remove(Context.ConnectionId, WebRole.users[Context.ConnectionId].ToString());
            return base.OnDisconnected();
        }
        */

        public void DiscardCard(int cardIndex)
        {
            HanabiPlayer p = WebRole.hanabiPlayers[Context.ConnectionId];
            Hanabi game = p.game;
            string error = null;
            if (!game.DiscardCard(p, cardIndex, ref error))
            {
                Clients.Caller.broadcastMessage(error);
            }
            BroadcastGame();
        }

        public void PlayCard(int cardIndex)
        {
            HanabiPlayer p = WebRole.hanabiPlayers[Context.ConnectionId];
            Hanabi game = p.game;
            string error = null;
            if (!game.PlayCard(p, cardIndex, ref error))
            {
                Clients.Caller.broadcastMessage(error);
            }
            BroadcastGame();
        }

        public void HintColor(int target, int cardIndex)
        {
            HanabiPlayer p = WebRole.hanabiPlayers[Context.ConnectionId];
            Hanabi game = p.game;
            string error = null;
            if (!game.HintColor(p, target, cardIndex, ref error))
            {
                Clients.Caller.broadcastMessage(error);
            }
            BroadcastGame();
        }
        public void HintNumber(int target, int cardIndex)
        {
            HanabiPlayer p = WebRole.hanabiPlayers[Context.ConnectionId];
            Hanabi game = p.game;
            string error = null;
            if (!game.HintNumber(p, target, cardIndex, ref error))
            {
                Clients.Caller.broadcastMessage(error);
            }
            BroadcastGame();
        }

        // Called when a user types something in the input box
        public void Send(String message)
        {
            HanabiPlayer p = WebRole.hanabiPlayers[Context.ConnectionId];
            Hanabi game = p.game;
            bool validAction = game.Act(p, message);
            if (validAction)
            {
                BroadcastGame();
            }
            else
            {
                if (!String.IsNullOrEmpty(message))
                {
                    Clients.Group(game.gameNumber.ToString()).broadcastMessage(p.name, message);
                }
            }
        }

        public void BroadcastGame()
        {
            Hanabi game = WebRole.hanabiPlayers[Context.ConnectionId].game;
            Clients.Group(game.gameNumber.ToString()).broadcastGame(game);
        }

        public void EndGame(string message)
        {
            Hanabi game = WebRole.hanabiPlayers[Context.ConnectionId].game;
            Clients.Group(game.gameNumber.ToString()).broadcastMessage(message);
        }
    }
}