using System.Collections.Generic;
using System.ComponentModel;
using Newtonsoft.Json;

namespace DictaNakdanVsto.Models
{
    public class DictaResponse
    {
        [JsonProperty("data")]
        public List<DictaToken> Data { get; set; }
    }

    public class DictaRequest
    {
        [JsonProperty("task")] public string Task { get; set; } = "nakdan";
        [JsonProperty("data")] public string Data { get; set; }
        [JsonProperty("genre")] public string Genre { get; set; } = "modern";
        [JsonProperty("addmorph")] public bool AddMorph { get; set; } = true;
        [JsonProperty("useTokenization")] public bool UseTokenization { get; set; } = true;
        [JsonProperty("keepqq")] public bool KeepQQ { get; set; } = true;
        [JsonProperty("nodageshdefmem")] public bool NoDageshDefMem { get; set; } = true;
        [JsonProperty("patachma")] public bool PatachMa { get; set; } = true;
        [JsonProperty("keepmetagim")] public bool KeepMetagim { get; set; } = true;
        [JsonProperty("matchpartial")] public bool MatchPartial { get; set; } = true;
        [JsonProperty("fullspelling")] public bool FullSpelling { get; set; } = false;
        [JsonProperty("ignoreoriginal")] public bool IgnoreOriginal { get; set; } = false;
        [JsonProperty("skipnikud")] public bool SkipNikud { get; set; } = true;
        [JsonProperty("skippartial")] public bool SkipPartial { get; set; } = false;
        [JsonProperty("skipsingle")] public bool SkipSingle { get; set; } = false;
        [JsonProperty("skipacronyms")] public bool SkipAcronyms { get; set; } = false;
        [JsonProperty("skipround")] public bool SkipRound { get; set; } = false;
        [JsonProperty("skipsquare")] public bool SkipSquare { get; set; } = false;
        [JsonProperty("skipcurly")] public bool SkipCurly { get; set; } = false;
        [JsonProperty("splitparentheses")] public bool SplitParentheses { get; set; } = false;
        [JsonProperty("shvadagesh")] public bool ShvaDagesh { get; set; } = false;
    }

    public class DictaToken
    {
        [JsonProperty("str")] public string Str { get; set; }
        [JsonProperty("sep")] public bool Sep { get; set; }
        [JsonProperty("nakdan")] public NakdanData Nakdan { get; set; }

        public bool IsPunctuation => Sep || Nakdan == null || Nakdan.Options == null || Nakdan.Options.Count == 0;
        public string Word => Str;
        public List<DictaOption> Options => Nakdan?.Options ?? new List<DictaOption>();
    }

    public class NakdanData
    {
        [JsonProperty("options")] public List<DictaOption> Options { get; set; }
        [JsonProperty("fconfident")] public bool? FConfident { get; set; }
    }

    public class DictaOption
    {
        [JsonProperty("w")] public string W { get; set; }
        [JsonProperty("levelChoice")] public int LevelChoice { get; set; }
        [JsonProperty("lex")] public string Lex { get; set; }
        public string CleanW => W?.Replace("|", "");
    }

    public class LetterNikud : INotifyPropertyChanged
    {
        private string _base, _dagesh, _sinDot, _vowel;
        public string Base { get => _base; set { _base = value; OnPropertyChanged(nameof(Base)); OnPropertyChanged(nameof(FullLetter)); } }
        public string Dagesh { get => _dagesh; set { _dagesh = value; OnPropertyChanged(nameof(Dagesh)); OnPropertyChanged(nameof(FullLetter)); } }
        public string SinDot { get => _sinDot; set { _sinDot = value; OnPropertyChanged(nameof(SinDot)); OnPropertyChanged(nameof(FullLetter)); } }
        public string Vowel { get => _vowel; set { _vowel = value; OnPropertyChanged(nameof(Vowel)); OnPropertyChanged(nameof(FullLetter)); } }
        public string FullLetter => $"{Base}{SinDot}{Dagesh}{Vowel}";
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}