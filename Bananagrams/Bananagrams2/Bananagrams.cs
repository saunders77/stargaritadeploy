using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Timers;
using System.Web;

namespace Bananagrams2
{
    public class Bananagrams
    {
        static int slowTimerInterval = 10; // seconds
        static int mediumTimerInterval = 7; // seconds
        static int fastTimerInterval = 4; // seconds

        static double maxLetters = 20;

        static List<char> shortLetterFrequency = ("EEEEEEEEEEEEAAAAAAAAA" +
                "IIIIIIIIIOOOOOOOONNNNNNRRRRRRTTTTTTT" +
                "LLLLSSSSSUUUUDDDDGGGBBCCMMPPFFHHVVWWYYKJXQZ").ToCharArray().ToList();

        static List<char> mediumLetterFrequency = ("EEEEEEEEEEEEEEEEEE" +
                "AAAAAAAAAAAAAIIIIIIIIIIIIOOOOOOOOOOO" +
                "TTTTTTTTTRRRRRRRRRNNNNNNNNDDDDDD" +
                "SSSSSSUUUUUULLLLLGGGGBBBCCCFFFHHHMMMPPPVVVWWWYYY" +
                "JJKKQQXXZZ").ToCharArray().ToList();

        static List<char> longLetterFrequency = ("EEEEEEEEEEEEEEEEEEEEEEEE" +
                "AAAAAAAAAAAAAAAAOOOOOOOOOOOOOOO" +
                "TTTTTTTTTTTTTTTIIIIIIIIIIIII" +
                "NNNNNNNNNNNNNRRRRRRRRRRRRR" +
                "SSSSSSSSSSLLLLLLLUUUUUUUDDDDDDDD" +
                "GGGGGCCCCCCMMMMMMBBBBPPPPHHHHH" +
                "FFFFWWWWYYYYVVVKKJJXXQQZZ").ToCharArray().ToList();

        static HashSet<String> wordList;

        public Bananagrams(int gameNumber, string gameType, BananagramsHub hub)
        {
            InitializeWordList();

            this.gameNumber = gameNumber;
            this.myHub = hub;
            players = new List<BananagramsPlayer>();
            letters = new List<char>();
            lastMove = new LastMove();
            random = new Random();
            timer = new Timer();
            timer.Elapsed += (sender, e) => Tick(sender, e);

            SetLetterFrequency(gameType);

            numLettersRemaining = letterFrequency.Count;
        }

        public int gameNumber { get; private set; }

        public List<char> letters { get; private set; }

        public int numLettersRemaining { get; private set; }

        private List<char> letterFrequency;
        private Random random;

        public int nextPlayer { get; private set; }

        public List<BananagramsPlayer> players { get; set; }

        public LastMove lastMove { get; private set; }

        public bool timerEnabled { get; private set; }

        public int timerInterval { get; private set; }

        private Timer timer;

        // Todo: this should be a subscriber/listener model?
        private BananagramsHub myHub;

        public void Reset(string gameType)
        {
            SetLetterFrequency(gameType);
            letters.Clear();
            numLettersRemaining = letterFrequency.Count;
            lastMove = new LastMove();
            foreach (BananagramsPlayer p in players)
            {
                p.Reset();
            }
            timer.Stop();
        }

        private void SetLetterFrequency(string gameType)
        {
            letterFrequency = new List<char>(mediumLetterFrequency);
            gameType = gameType.ToUpper();
            switch (gameType)
            {
                case "LONG":
                case "SLOW":
                    //letterFrequency = new List<char>(longLetterFrequency);
                    timerInterval = slowTimerInterval;
                    break;
                case "FAST":
                case "SHORT":
                    //letterFrequency = new List<char>(shortLetterFrequency);
                    timerInterval = fastTimerInterval;
                    break;
                case "MEDIUM":
                default:
                    //letterFrequency = new List<char>(mediumLetterFrequency);
                    timerInterval = mediumTimerInterval;
                    break;
            }
        }

        private void Tick(object sender, ElapsedEventArgs e)
        {
            if (numLettersRemaining > 0)
            {
                Flip(players[nextPlayer]);
                myHub.BroadcastGame();
            }
        }

        private void ResetFlipTimer()
        {
            if (!timer.Enabled)
            {
                timer.Start();
                timerEnabled = true;
            }
            timer.Interval = timerInterval * 1000; // milliseconds
        }

        public void Flip(BananagramsPlayer p)
        {
            //if (players[nextPlayer].Equals(p) && letterFrequency.Count > 0)
            if (letterFrequency.Count > 0)
            {
                int randomNumber = random.Next(0, letterFrequency.Count);
                char letter = letterFrequency[randomNumber];
                letterFrequency.RemoveAt(randomNumber);
                numLettersRemaining = letterFrequency.Count;
                letters.Add(letter);

                if (letters.Count > maxLetters)
                {
                    letters.RemoveAt(0);
                }

                nextPlayer = (nextPlayer + 1) % players.Count;
                ResetFlipTimer();

                return;
            }
        }

        public int GetPlayerNumber(BananagramsPlayer player)
        {
            for (int i = 0; i < players.Count; i++)
            {
                if (players[i].Equals(player))
                {
                    return i;
                }
            }
            return -1;
        }

        private bool FormWordByStealing(BananagramsPlayer guesser, string s)
        {
            // See if we can form s
            int guesserNumber = GetPlayerNumber(guesser);
            bool canFormWord = false;
            for (int i = 1; i <= players.Count; i++)
            {
                int playerNumber = (guesserNumber + i) % players.Count;
                BananagramsPlayer player = players[playerNumber];
                foreach (string word in player.words)
                {
                    if (s.Length <= word.Length)
                    {
                        continue;
                    }

                    bool ok = true;
                    List<char> leftoverChars = new List<char>(letters);
                    List<char> wordChars = word.ToCharArray().ToList();

                    char[] needChars = s.ToCharArray();

                    foreach (char c in needChars)
                    {
                        if (wordChars.Contains(c))
                        {
                            wordChars.Remove(c);
                        }
                        else if (leftoverChars.Contains(c))
                        {
                            leftoverChars.Remove(c);
                        }
                        else
                        {
                            ok = false;
                            break;
                        }
                    }
                    if (ok && wordChars.Count == 0)
                    {
                        guesser.words.Add(s);
                        player.words.Remove(word);
                        letters = leftoverChars;

                        lastMove.word = s;
                        lastMove.winningPlayerNumber = guesserNumber;
                        lastMove.losingWord = word;
                        lastMove.losingPlayerNumber = playerNumber;

                        canFormWord = true;
                        break;
                    }
                }
                if (canFormWord)
                {
                    break;
                }
            }
            return canFormWord;
        }

        private bool FormWordFromPool(BananagramsPlayer guesser, string s)
        {
            int guesserNumber = GetPlayerNumber(guesser);
            bool canFormWord = true;
            List<char> leftoverChars = new List<char>(letters);

            char[] needChars = s.ToCharArray();

            foreach (char c in needChars)
            {
                if (leftoverChars.Contains(c))
                {
                    leftoverChars.Remove(c);
                }
                else
                {
                    canFormWord = false;
                    break;
                }
            }
            if (canFormWord)
            {
                guesser.words.Add(s);
                letters = leftoverChars;

                lastMove.word = s;
                lastMove.winningPlayerNumber = guesserNumber;
                lastMove.losingPlayerNumber = -1;
                lastMove.losingWord = null;
            }

            return canFormWord;
        }

        /// <summary>
        /// Processes a text input from the user
        /// </summary>
        /// <param name="guesser">The player whose input is being processed</param>
        /// <param name="s">The string that the player typed</param>
        /// <returns>Whether the string is a valid command, or a valid word</returns>
        public bool Guess(BananagramsPlayer guesser, string s)
        {
            lastMove.isChanged = false;
            s = s.ToUpper();
            if (s.Equals("F"))
            {
                Flip(guesser);
                return true;
            }
            else if (s.Equals("C"))
            {
                ReverseMove();
                return true;
            }
            else if (s.Equals("P"))
            {
                if (timer.Enabled)
                {
                    timer.Stop();
                    timerEnabled = false;
                }
                else
                {
                    timer.Start();
                    timerEnabled = true;
                }
                return true;
            }
            else if (s.Equals("NEW SHORT") || s.Equals("NEW MEDIUM") || s.Equals("NEW LONG"))
            {
                Reset(s.Substring(4));
                return true;
            }
            else if (s.Length < 3)
            {
                return false;
            }
            else if (!wordList.Contains(s))
            {
                return false;
            }

            bool canFormWord = FormWordByStealing(guesser, s);

            // TODO: Make this a function?
            if (!canFormWord)
            {
                canFormWord = FormWordFromPool(guesser, s);
            }

            if (canFormWord)
            {
                lastMove.isChanged = true;
                ResetFlipTimer();
            }

            return canFormWord;
        }

        public void ReverseMove()
        {
            if (lastMove.word != null)
            {
                List<char> leftoverChars = lastMove.word.ToCharArray().ToList();

                players[lastMove.winningPlayerNumber].words.Remove(lastMove.word);
                if (lastMove.losingWord != null && lastMove.losingPlayerNumber >= 0)
                {
                    players[lastMove.losingPlayerNumber].words.Add(lastMove.losingWord);
                    foreach (char c in lastMove.losingWord.ToCharArray())
                    {
                        leftoverChars.Remove(c);
                    }
                }

                letters.AddRange(leftoverChars);

                lastMove.losingPlayerNumber = -1;
                lastMove.winningPlayerNumber = -1;
                lastMove.losingWord = null;
                lastMove.word = null;
            }
        }

        public class LastMove
        {
            public LastMove()
            {
                word = null;
                losingWord = null;
                losingPlayerNumber = -1;
                winningPlayerNumber = -1;
            }
            public string word { get; set; }
            public int losingPlayerNumber { get; set; }
            public int winningPlayerNumber { get; set; }
            public string losingWord { get; set; }
            public bool isChanged { get; set; }
        }

        private void InitializeWordList()
        {
            if (wordList == null)
            {
                TextReader tr = new StreamReader(HttpContext.Current.Server.MapPath("~/wordlist.txt"));
                wordList = new HashSet<string>();
                String word = null;
                while ((word = tr.ReadLine()) != null)
                {
                    wordList.Add(word);
                }
            }
        }
    }
}