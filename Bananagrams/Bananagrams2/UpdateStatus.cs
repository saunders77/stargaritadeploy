using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;

// Add reference
using Microsoft.AspNet.SignalR;
using Microsoft.AspNet.SignalR.Hubs;

namespace Bananagrams2
{
    // Derive from Hub, which is a server-side class
    // and a client side proxy.
    public class UpdateStatus : Hub
    {
        public Game Create(string gameType)
        {
            WebRole.InitDB();
            Random r = new Random();
            int gameNumber = r.Next(1, 10000);
            
            Game game = new Game(gameNumber, gameType, this);
            WebRole.games.Add(gameNumber, game);

            Groups.Add(Context.ConnectionId, gameNumber.ToString());

            return game;
        }
        public Game Join(string gameString)
        {
            WebRole.InitDB();
            Game game;

            int gameNumber = -1;

            try
            {
                gameNumber = int.Parse(gameString);
            }
            catch
            {
                gameNumber = -1;
            }

            if (WebRole.games.ContainsKey(gameNumber))
            {
                game = WebRole.games[gameNumber];
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
            Game game = WebRole.games[gameNumber];
            Player player = null;

            foreach (Player p in game.players)
            {
                if (p.name == playerName)
                {
                    player = p;
                }
            }

            if (player == null)
            {
                player = new Player(game, playerName);
                WebRole.users.Add(Context.ConnectionId, player);
                game.players.Add(player);
            }
            else
            {
                WebRole.users.Add(Context.ConnectionId, player);
            }

            BroadcastGame();
        }

        public override Task OnDisconnected()
        {
            // TODO: fix this up
            //Groups.Remove(Context.ConnectionId, WebRole.users[Context.ConnectionId].ToString());
            return base.OnDisconnected();
        }

        public void Flip()
        {
            Player p = WebRole.users[Context.ConnectionId];
            Game game = p.game;
            game.Flip(p);
            BroadcastGame();
        }

        public void Challenge()
        {
            Player p = WebRole.users[Context.ConnectionId];
            Game game = p.game;
            game.ReverseMove();
            BroadcastGame();
        }

        // Called when a user types something in the input box
        public void Send(String message)
        {
            Player p = WebRole.users[Context.ConnectionId];
            Game game = p.game;
            bool validWord = game.Guess(p, message);
            if (validWord)
            {
                Game.LastMove lastMove = game.lastMove;
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
            Game game = WebRole.users[Context.ConnectionId].game;
            Clients.Group(game.gameNumber.ToString()).broadcastGame(game);
        }
    }
}