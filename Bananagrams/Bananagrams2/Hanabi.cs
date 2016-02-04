using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System;
using System.Collections;
using System.Collections.Generic;

namespace Bananagrams2
{
    public class Hanabi
    {
        public enum CardColor
        {
            Unknown,
            Grey,
            Red,
            Blue,
            Green,
            Orange,
            Rainbow
        }

        public class CardComparer : IComparer<Card>
        {
            int IComparer<Card>.Compare(Card x, Card y)
            {
                if (x.color != y.color)
                {
                    return (x.color - y.color);
                }
                return (x.number - y.number);
            }
        }

        public class Card
        {
            [JsonConverter(typeof(StringEnumConverter))]
            public CardColor color { get; private set; }
            public int number { get; private set; }
            public bool revealedColor { get; private set; }
            public bool revealedNumber { get; private set; }

            public Card(CardColor color, int number, bool revealed)
            {
                this.color = color;
                this.number = number;
                revealedColor = revealed;
                revealedNumber = revealed;
            }

            public bool CanBePlacedOn(Card t)
            {
                return this.MatchesColor(t) && (this.number == t.number + 1);
            }

            public bool MatchesColor(Card t)
            {
                return (this.color == t.color);
            }

            public bool MatchesNumber(Card t)
            {
                return (this.number == t.number);
            }

            public bool RevealColor(CardColor color)
            {
                if (this.color == color)
                {
                    RevealColor();
                    return true;
                }
                return false;
            }
            public bool RevealNumber(int number)
            {
                if (this.number == number)
                {
                    RevealNumber();
                    return true;
                }
                return false;
            }
            public void RevealColor()
            {
                revealedColor = true;
            }
            public void RevealNumber()
            {
                revealedNumber = true;
            }

            public void Reveal()
            {
                RevealColor();
                RevealNumber();
            }
        }

        List<CardColor> colorsInGame;
        List<Card> cardsDeck;
        public int numCardsRemaining { get; private set; }
        public List<Card> cardsDiscarded { get; private set; }
        public Dictionary<CardColor, Card> cardsPlayed { get; private set; }
        public int numClues { get; private set; }
        int maxClues;
        public int numFuses { get; private set; }
        int maxFuses;
        int maxCardsInHand;
        int maxCardNumber;
        int numFinalFlipsRemaining;
        public int score { get; private set; } 
        int maxScore;

        public int gameNumber { get; private set; }
        private Random random;

        public int nextPlayer { get; private set; }

        public List<HanabiPlayer> players { get; set; }

        // Could be an action instead of a card...
        public Card lastAction { get; private set; }

        private CardComparer cc = new CardComparer();
        // Todo: this should be a subscriber/listener model?
        private HanabiHub myHub;

        bool IsValidColor(CardColor cardColor)
        {
            return (colorsInGame.Contains(cardColor));
        }

        void InitializeMembers()
        {
            cardsDeck = new List<Card>();
            cardsDiscarded = new List<Card>();
            cardsPlayed = new Dictionary<CardColor, Card>();
            lastAction = null;
            score = 0;
        }

        // Currently everything is hardcoded, but can be modified here.
        void ChooseParameters()
        {
            maxClues = 8;
            maxFuses = 4;
            maxCardsInHand = 5;
            maxCardNumber = 5;
            
            // Choose the colors to play with. Currently hardcoded.
            colorsInGame = new List<CardColor>();
            foreach (CardColor cardColor in Enum.GetValues(typeof(CardColor)))
            {
                if (cardColor != CardColor.Unknown && cardColor != CardColor.Rainbow)
                {
                    colorsInGame.Add(cardColor);
                }
            }

            maxScore = maxCardNumber * colorsInGame.Count;
        }

        void SetupGame()
        {
            numClues = maxClues;
            numFuses = maxFuses;

            foreach (CardColor cardColor in colorsInGame)
            {
                cardsPlayed[cardColor] = new Card(cardColor, 0, true);

                cardsDeck.Add(new Card(cardColor, 1, false));
                cardsDeck.Add(new Card(cardColor, 1, false));
                cardsDeck.Add(new Card(cardColor, 1, false));
                cardsDeck.Add(new Card(cardColor, 2, false));
                cardsDeck.Add(new Card(cardColor, 2, false));
                cardsDeck.Add(new Card(cardColor, 3, false));
                cardsDeck.Add(new Card(cardColor, 3, false));
                cardsDeck.Add(new Card(cardColor, 4, false));
                cardsDeck.Add(new Card(cardColor, 4, false));
                cardsDeck.Add(new Card(cardColor, 5, false));
            }

            foreach (HanabiPlayer player in players)
            {
                for (int i = 0; i < maxCardsInHand; i++)
                {
                    AddCardToPlayer(player);
                }
            }
            numFinalFlipsRemaining = players.Count;
        }


        public Hanabi(int gameNumber, string gameType, HanabiHub hub)
        {
            InitializeMembers();
            random = new Random();

            players = new List<HanabiPlayer>();
            this.gameNumber = gameNumber;
            this.myHub = hub;

            ChooseParameters();
            //SetupGame();
        }

        Card RemoveCardFromPlayer(HanabiPlayer player, int cardIndex)
        {
            Card card = player.cards[cardIndex];
            player.cards.RemoveAt(cardIndex);
            card.Reveal();
            return card;
        }

        void AddCardToPlayer(HanabiPlayer player)
        {
            if (cardsDeck.Count == 0)
            {
                if (numFinalFlipsRemaining <= 0)
                {
                    EndGame("No more turns left!");
                }
                else
                {
                    numFinalFlipsRemaining--;
                }
                return;
            }

            int nextInt = random.Next(cardsDeck.Count - 1);
            Card card = cardsDeck[nextInt];
            cardsDeck.RemoveAt(nextInt);
            numCardsRemaining = cardsDeck.Count;
            player.cards.Add(card);
        }

        void IncrementClues()
        {
            if (numClues < maxClues)
            {
                numClues++;
            }
        }

        bool DecrementClues(ref string error)
        {
            if (numClues > 0)
            {
                numClues--;
                return true;
            }
            else
            {
                error = "No clues left!";
                return false;
            }
        }

        void IncrementFuses()
        {
            if (numFuses < maxFuses)
            {
                numFuses++;
            }
        }

        void DecrementFuses()
        {
            if (numFuses > 0)
            {
                numFuses--;
            }
            if (numFuses <= 0)
            {
                EndGame("No fuses left!");
            }
        }

        bool GameWon()
        {
            return score == maxScore;
        }

        void EndGame(string message)
        {
            if (GameWon())
            {
                message += " You win!";
            }
            else
            {
                message += " Your score was " + score.ToString() + ". ";
                if (score == maxScore - 1)
                {
                    message += "So close!";
                }
                else if (score >= maxScore - 3)
                {
                    message += "Not bad!";
                }
                else if (score >= maxScore - 5)
                {
                    message += "\nFireworks are expensive. If at first you don't succeed, maybe pyrotechnics aren't for you.";
                }
                else if (score >= maxScore / 2)
                {
                    message += "Better luck next time.";
                }
                else
                {
                    message += "Were you even trying?";
                }
            }
            myHub.EndGame(message);
        }

        public void AddPlayer(HanabiPlayer p)
        {
            players.Add(p);       
        }

        public void Reset(string gameType)
        {
            InitializeMembers();
            foreach (HanabiPlayer p in players)
            {
                p.Reset();
            }
            SetupGame();
        }

        public bool PlayCard(HanabiPlayer player, int cardIndex, ref string error)
        {
            if (!IsNextPlayer(player, ref error) || cardIndex >= maxCardsInHand)
            {
                return false;
            }

            Card card = RemoveCardFromPlayer(player, cardIndex);
            lastAction = card;

            if (card.CanBePlacedOn(cardsPlayed[card.color]))
            {
                cardsPlayed[card.color] = card;
                score++;
                if (card.number == maxCardNumber)
                {
                    IncrementClues();
                    IncrementFuses();
                }
            }
            else
            {
                cardsDiscarded.Add(card);
                cardsDiscarded.Sort(cc);
                DecrementFuses();
            }

            if (GameWon())
            {
                EndGame("You win!");
            }
            else
            {
                AddCardToPlayer(player);
            }

            return EndTurn();
        }

        public bool DiscardCard(HanabiPlayer player, int cardIndex, ref string error)
        {
            if (!IsNextPlayer(player, ref error) || cardIndex >= maxCardsInHand)
            {
                return false;
            }

            Card card = RemoveCardFromPlayer(player, cardIndex);
            lastAction = card;
            cardsDiscarded.Add(card);
            cardsDiscarded.Sort(cc);
            IncrementClues();
            AddCardToPlayer(player);

            return EndTurn();
        }

        public bool HintColor(HanabiPlayer player, int target, int cardIndex, ref string error)
        {
            HanabiPlayer targetPlayer = players[target];
            if (!IsNextPlayer(player, ref error) || player == targetPlayer || cardIndex >= targetPlayer.cards.Count)
            {
                return false;
            }

            if (!DecrementClues(ref error))
            {
                return false;
            }

            CardColor color = targetPlayer.cards[cardIndex].color;
            foreach (Card card in targetPlayer.cards)
            {
                card.RevealColor(color);
            }
            lastAction = new Card(color, 0, false);
            lastAction.RevealColor();
            return EndTurn();
        }

        public bool HintNumber(HanabiPlayer player, int target, int cardIndex, ref string error)
        {
            HanabiPlayer targetPlayer = players[target];
            if (!IsNextPlayer(player, ref error) || player == targetPlayer || cardIndex >= targetPlayer.cards.Count)
            {
                return false;
            }

            if (!DecrementClues(ref error))
            {
                return false;
            }

            int number = targetPlayer.cards[cardIndex].number;
            foreach (Card card in targetPlayer.cards)
            {
                card.RevealNumber(number);
            }
            lastAction = new Card(CardColor.Unknown, number, false);
            lastAction.RevealNumber();
            return EndTurn();
        }

        /// <summary>
        /// Processes a text input from the user
        /// </summary>
        /// <param name="player">The player whose input is being processed</param>
        /// <param name="s">The string that the player typed</param>
        /// <returns>Whether the string is a valid command, or a valid word</returns>
        public bool Act(HanabiPlayer player, string s)
        {
            s = s.ToUpper();
            if (s.Equals("RESET"))
            {
                Reset("");
                return true;
            }

            string error = null;
            if (!IsNextPlayer(player, ref error))
            {
                return false;
            }

            bool success = false;
            if (s.StartsWith("PLAY "))
            {
                success = PlayCard(player, int.Parse(s.Substring(5)), ref error);
            }
            else if (s.StartsWith("DISCARD "))
            {
                success = DiscardCard(player, int.Parse(s.Substring(8)), ref error);
            }
            else if (s.StartsWith("HINT COLOR "))
            {
                success = HintColor(player, int.Parse(s.Substring(11,1)), int.Parse(s.Substring(13)), ref error);
            }
            else if (s.StartsWith("HINT NUMBER "))
            {
                success = HintNumber(player, int.Parse(s.Substring(12,1)), int.Parse(s.Substring(14)), ref error);
            }
            return success;
        }

        private bool IsNextPlayer(HanabiPlayer player, ref string error)
        {
            if (players[nextPlayer] == player)
            {
                return true;
            }
            else
            {
                error = "It's not your turn!";
                return false;
            }
        }

        private bool EndTurn()
        {
            nextPlayer = (nextPlayer + 1) % players.Count;
            return true;
        }

        public enum ActionType
        {
            NullAction,
            PlayCard,
            DiscardCard,
            HintColor,
            HintNumber
        }
        public class Action
        {
            public ActionType actionType { get; set; }
            public Card card { get; set; }
            public HanabiPlayer sourcePlayer { get; set; }
            public HanabiPlayer targetPlayer { get; set; }

            public Action(ActionType actionType)
            {
                this.actionType = actionType;
            }
        }
    }

    public class HanabiPlayer
    {
        public string name { get; private set; }

        public List<Hanabi.Card> cards { get; set; }

        [Newtonsoft.Json.JsonIgnore]
        public Hanabi game { get; set; }
        public HanabiPlayer(Hanabi game, string name)
        {
            this.game = game;
            this.name = name;
            cards = new List<Hanabi.Card>();
        }

        public bool Equals(HanabiPlayer p)
        {
            if (p == null)
            {
                return false;
            }

            return (name == p.name);
        }

        public void Reset()
        {
            cards.Clear();
        }
    }
}