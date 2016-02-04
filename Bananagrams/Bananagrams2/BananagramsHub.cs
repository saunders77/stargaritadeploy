using System;
using System.Threading.Tasks;

// Add reference
using Microsoft.AspNet.SignalR;

namespace Bananagrams2
{
    // Derive from Hub, which is a server-side class
    // and a client side proxy.
    public class BananagramsHub : Hub
    {
        public Bananagrams Create(string gameType)
        {
            WebRole.InitDB();
            Random r = new Random();
            int gameNumber = r.Next(1, 10000);

            Bananagrams game = new Bananagrams(gameNumber, gameType, this);
            WebRole.bananagramsGames.Add(gameNumber, game);

            Groups.Add(Context.ConnectionId, gameNumber.ToString());

            return game;
        }
        public Bananagrams Join(string gameString)
        {
            WebRole.InitDB();
            Bananagrams game;

            int gameNumber = -1;

            try
            {
                gameNumber = int.Parse(gameString);
            }
            catch
            {
                gameNumber = -1;
            }

            if (WebRole.bananagramsGames.ContainsKey(gameNumber))
            {
                game = WebRole.bananagramsGames[gameNumber];
            }
            else
            {
                game = Create("medium");
                gameNumber = game.gameNumber;
            }

            Groups.Add(Context.ConnectionId, gameNumber.ToString());
            return game;
        }

        public void JoinAsPlayer(string gameNumberString, string playerName)
        {
            int gameNumber = int.Parse(gameNumberString);
            Bananagrams game = WebRole.bananagramsGames[gameNumber];
            BananagramsPlayer player = null;

            foreach (BananagramsPlayer p in game.players)
            {
                if (p.name == playerName)
                {
                    player = p;
                }
            }

            if (player == null)
            {
                player = new BananagramsPlayer(game, playerName);
                WebRole.bananagramsPlayers.Add(Context.ConnectionId, player);
                game.players.Add(player);
            }
            else
            {
                WebRole.bananagramsPlayers.Add(Context.ConnectionId, player);
            }

            BroadcastGame();
        }

        /*
        public override Task OnDisconnected()
        {
            // TODO: fix this up
            //Groups.Remove(Context.ConnectionId, WebRole.users[Context.ConnectionId].ToString());
            return base.OnDisconnected();
        }
        */

        public void Flip()
        {
            BananagramsPlayer p = WebRole.bananagramsPlayers[Context.ConnectionId];
            Bananagrams game = p.game;
            game.Flip(p);
            BroadcastGame();
        }

        public void Challenge()
        {
            BananagramsPlayer p = WebRole.bananagramsPlayers[Context.ConnectionId];
            Bananagrams game = p.game;
            game.ReverseMove();
            BroadcastGame();
        }

        // Called when a user types something in the input box
        public void Send(String message)
        {
            BananagramsPlayer p = WebRole.bananagramsPlayers[Context.ConnectionId];
            Bananagrams game = p.game;
            bool validWord = game.Guess(p, message);
            if (validWord)
            {
                Bananagrams.LastMove lastMove = game.lastMove;
                if (lastMove.isChanged)
                {
                    String winningPlayer = game.players[lastMove.winningPlayerNumber].name;
                    String lastMoveMessage;
                    if (lastMove.losingPlayerNumber >= 0)
                    {
                        if (lastMove.losingPlayerNumber == lastMove.winningPlayerNumber)
                        {
                            lastMoveMessage = String.Format("{0} added to {1} to spell {2}.", winningPlayer, lastMove.losingWord, lastMove.word);
                        }
                        else
                        {
                            lastMoveMessage = String.Format("{0} stole {1} from {2} to spell {3}.",
                                                                winningPlayer,
                                                                lastMove.losingWord,
                                                                game.players[lastMove.losingPlayerNumber].name,
                                                                lastMove.word);
                        }
                    }
                    else
                    {
                        lastMoveMessage = String.Format("{0} spelled {1}.",
                                                                winningPlayer,
                                                                lastMove.word);
                    }
                    Clients.Group(game.gameNumber.ToString()).broadcastMessage("Game", lastMoveMessage);
                }

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
            Bananagrams game = WebRole.bananagramsPlayers[Context.ConnectionId].game;
            Clients.Group(game.gameNumber.ToString()).broadcastGame(game);
        }
    }
}