using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Windows.Threading; // ספרייה שנוספה לניהול תהליכים
using DictaNakdanVsto.Models;
using DictaNakdanVsto.Services;
using Newtonsoft.Json;
using Word = Microsoft.Office.Interop.Word;

namespace DictaNakdanVsto.ViewModels
{
    public class RelayCommand : ICommand
    {
        private readonly Action<object> _execute;
        public RelayCommand(Action<object> execute) => _execute = execute;
        public event EventHandler CanExecuteChanged { add { } remove { } }
        public bool CanExecute(object parameter) => true;
        public void Execute(object parameter) => _execute(parameter);
    }

    public class NakdanViewModel : INotifyPropertyChanged
    {
        private readonly DictaApiService _apiService = new DictaApiService();
        private readonly DictaInteropService _interopService = new DictaInteropService();
        private readonly Dispatcher _dispatcher; // שומר את התהליך הראשי של הממשק
        private string _settingsPath;

        public NakdanViewModel()
        {
            // לוכד את התהליך הראשי כשהתוסף עולה
            _dispatcher = Dispatcher.CurrentDispatcher;

            _settingsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "DictaNakdanVsto", "settings.json");
            LoadSettings();

            PunctuateCommand = new RelayCommand(async (o) => await StartPunctuationAsync());
            CancelCommand = new RelayCommand((o) => FinishPunctuation());
            SelectOptionCommand = new RelayCommand(async (o) => await ApplyOptionAsync(o.ToString()));

            ManualModeCommand = new RelayCommand((o) => StartManualMode());
            CancelManualModeCommand = new RelayCommand((o) => { IsManualMode = false; IsOptionsMode = true; });
            ApplyManualModeCommand = new RelayCommand(async (o) => await ApplyManualModeAsync());
            SetNikudCommand = new RelayCommand(ApplyNikudToLetter);
        }

        // פונקציית קסם שדואגת שכל שינוי בממשק יקרה בתהליך המותר ולא יקריס את המערכת
        private void RunOnUI(Action action)
        {
            if (_dispatcher.CheckAccess()) action();
            else _dispatcher.Invoke(action);
        }

        private bool _isProcessing, _isOptionsMode, _isManualMode;

        public bool IsProcessing { get => _isProcessing; set { _isProcessing = value; OnPropertyChanged(nameof(IsProcessing)); UpdateModes(); } }
        public bool IsOptionsMode { get => _isOptionsMode; set { _isOptionsMode = value; OnPropertyChanged(nameof(IsOptionsMode)); UpdateModes(); } }
        public bool IsManualMode { get => _isManualMode; set { _isManualMode = value; OnPropertyChanged(nameof(IsManualMode)); UpdateModes(); } }
        public bool IsMainMode => !IsProcessing && !IsOptionsMode && !IsManualMode;
        private void UpdateModes() => OnPropertyChanged(nameof(IsMainMode));

        public string StatusMessage { get; set; } = "מוכן לניקוד";
        public bool StopOnEveryWord { get; set; } = false;

        private List<DictaToken> _sessionTokens;
        private int _currentIndex = 0;
        private Word.Range _searchBounds;
        private Word.Range _currentWordRange;

        public string CurrentWordStr { get; set; }
        public ObservableCollection<DictaOption> Level1Options { get; set; } = new ObservableCollection<DictaOption>();
        public ObservableCollection<DictaOption> OtherOptions { get; set; } = new ObservableCollection<DictaOption>();

        public ObservableCollection<LetterNikud> ManualLetters { get; set; } = new ObservableCollection<LetterNikud>();

        private LetterNikud _selectedLetter;
        public LetterNikud SelectedLetter
        {
            get => _selectedLetter;
            set
            {
                _selectedLetter = value;
                OnPropertyChanged("");
            }
        }

        public string CurrentBaseLetter => SelectedLetter?.Base ?? " ";
        public string CurrentBaseWithModifiers => CurrentBaseLetter + (SelectedLetter?.SinDot ?? "") + (SelectedLetter?.Dagesh ?? "");

        public string KbDagesh => CurrentBaseLetter + "\u05BC";
        public string KbShin => CurrentBaseLetter + "\u05C1" + (SelectedLetter?.Dagesh ?? "");
        public string KbSin => CurrentBaseLetter + "\u05C2" + (SelectedLetter?.Dagesh ?? "");

        public string KbKamatz => CurrentBaseWithModifiers + "\u05B8";
        public string KbPatach => CurrentBaseWithModifiers + "\u05B7";
        public string KbTzere => CurrentBaseWithModifiers + "\u05B5";
        public string KbSegol => CurrentBaseWithModifiers + "\u05B6";
        public string KbHirik => CurrentBaseWithModifiers + "\u05B4";
        public string KbHolam => CurrentBaseWithModifiers + "\u05B9";
        public string KbKubutz => CurrentBaseWithModifiers + "\u05BB";
        public string KbShva => CurrentBaseWithModifiers + "\u05B0";
        public string KbHatafKamatz => CurrentBaseWithModifiers + "\u05B3";
        public string KbHatafPatach => CurrentBaseWithModifiers + "\u05B2";
        public string KbHatafSegol => CurrentBaseWithModifiers + "\u05B1";

        public bool IsDageshActive => !string.IsNullOrEmpty(SelectedLetter?.Dagesh);
        public bool IsShinActive => SelectedLetter?.SinDot == "\u05C1";
        public bool IsSinActive => SelectedLetter?.SinDot == "\u05C2";
        public bool IsShinLetter => CurrentBaseLetter == "ש";

        public DictaRequest ApiSettings { get; set; } = new DictaRequest();

        public bool IsGenreModern { get => ApiSettings.Genre == "modern"; set { if (value) ApiSettings.Genre = "modern"; } }
        public bool IsGenreRabbinic { get => ApiSettings.Genre == "rabbinic"; set { if (value) ApiSettings.Genre = "rabbinic"; } }
        public bool IsGenrePoetry { get => ApiSettings.Genre == "modernpoetry"; set { if (value) ApiSettings.Genre = "modernpoetry"; } }

        public ICommand PunctuateCommand { get; }
        public ICommand CancelCommand { get; }
        public ICommand SelectOptionCommand { get; }
        public ICommand ManualModeCommand { get; }
        public ICommand CancelManualModeCommand { get; }
        public ICommand ApplyManualModeCommand { get; }
        public ICommand SetNikudCommand { get; }

        private async Task StartPunctuationAsync()
        {
            if (IsProcessing) return;
            SaveSettings();
            IsProcessing = true;
            StatusMessage = "שואב טקסט מוורד...";

            try
            {
                _searchBounds = _interopService.GetInitialSearchRange();
                string text = _searchBounds.Text;
                if (string.IsNullOrWhiteSpace(text))
                {
                    FinishPunctuation();
                    return;
                }

                StatusMessage = "מנקד מול שרתי דיקטה...";
                var sentences = _apiService.SplitToSentences(text);
                var tempTokens = new List<DictaToken>();

                foreach (var sentence in sentences)
                {
                    tempTokens.AddRange(await _apiService.GetNakdanAsync(sentence, ApiSettings));
                }

                // לאחר שסיימנו למשוך מהשרת ברקע, אנו חוזרים לתהליך הראשי (UI) כדי לא להקריס
                RunOnUI(() =>
                {
                    _sessionTokens = tempTokens;

                    if (_sessionTokens.Count == 0)
                    {
                        System.Windows.Forms.MessageBox.Show("השרת לא החזיר תוצאות לניקוד.", "מידע", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                        FinishPunctuation();
                        return;
                    }

                    _interopService.StartUndoRecord();
                    _currentIndex = 0;
                    ProcessNextWord();
                });
            }
            catch (Exception ex)
            {
                RunOnUI(() => {
                    System.Windows.Forms.MessageBox.Show($"אירעה שגיאה:\n{ex.Message}", "שגיאה", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    FinishPunctuation();
                });
            }
        }

        private void ProcessNextWord()
        {
            while (_currentIndex < _sessionTokens.Count)
            {
                var token = _sessionTokens[_currentIndex];
                if (token.IsPunctuation || string.IsNullOrWhiteSpace(token.Word) || token.Options.Count == 0)
                {
                    _currentIndex++;
                    continue;
                }

                var uniqueOptions = token.Options.GroupBy(o => o.CleanW).Select(g => g.First()).ToList();
                bool requiresAttention = uniqueOptions.Count > 1 || StopOnEveryWord;

                _currentWordRange = _interopService.FindAndSelectWord(token.Word, _searchBounds);
                if (_currentWordRange != null)
                {
                    if (requiresAttention)
                    {
                        CurrentWordStr = token.Word;
                        Level1Options.Clear();
                        OtherOptions.Clear();
                        foreach (var opt in uniqueOptions)
                        {
                            if (opt.LevelChoice == 1) Level1Options.Add(opt);
                            else OtherOptions.Add(opt);
                        }

                        OnPropertyChanged(nameof(CurrentWordStr));
                        IsProcessing = false;
                        IsOptionsMode = true;
                        return;
                    }
                    else
                    {
                        _interopService.ReplaceWordAndUpdateBounds(_currentWordRange, uniqueOptions[0].CleanW, ref _searchBounds);
                    }
                }
                _currentIndex++;
            }
            FinishPunctuation();
        }

        private async Task ApplyOptionAsync(string chosenText)
        {
            if (_currentWordRange != null)
            {
                _interopService.ReplaceWordAndUpdateBounds(_currentWordRange, chosenText, ref _searchBounds);
            }
            _currentIndex++;

            IsOptionsMode = false;
            IsProcessing = true;
            StatusMessage = "ממשיך...";

            await Task.Delay(50);

            // חוזרים לתהליך הראשי אחרי ה-Delay
            RunOnUI(() => ProcessNextWord());
        }

        private void FinishPunctuation()
        {
            if (_interopService != null) _interopService.EndUndoRecord();
            IsOptionsMode = false;
            IsManualMode = false;
            IsProcessing = false;
            StatusMessage = "מוכן לניקוד";
            _sessionTokens = null;
        }

        private void StartManualMode()
        {
            IsOptionsMode = false;
            IsManualMode = true;
            ManualLetters.Clear();

            if (CurrentWordStr != null)
            {
                foreach (char c in CurrentWordStr)
                    ManualLetters.Add(new LetterNikud { Base = c.ToString(), Dagesh = "", SinDot = "", Vowel = "" });

                if (ManualLetters.Count > 0)
                    SelectedLetter = ManualLetters[0];
            }
        }

        private void ApplyNikudToLetter(object parameter)
        {
            if (SelectedLetter == null || parameter == null) return;
            string symbol = parameter.ToString();

            if (symbol == "DAGESH") SelectedLetter.Dagesh = string.IsNullOrEmpty(SelectedLetter.Dagesh) ? "\u05BC" : "";
            else if (symbol == "\u05C1" || symbol == "\u05C2") SelectedLetter.SinDot = SelectedLetter.SinDot == symbol ? "" : symbol;
            else if (symbol == "CLEAR") { SelectedLetter.Vowel = ""; SelectedLetter.Dagesh = ""; SelectedLetter.SinDot = ""; }
            else SelectedLetter.Vowel = symbol;

            int index = ManualLetters.IndexOf(SelectedLetter);

            if (index >= 0 && index < ManualLetters.Count - 1 && symbol != "DAGESH" && symbol != "\u05C1" && symbol != "\u05C2" && symbol != "CLEAR")
            {
                SelectedLetter = ManualLetters[index + 1];
            }

            OnPropertyChanged("");
        }

        private async Task ApplyManualModeAsync() => await ApplyOptionAsync(string.Join("", ManualLetters.Select(l => l.FullLetter)));

        private void SaveSettings() { try { Directory.CreateDirectory(Path.GetDirectoryName(_settingsPath)); File.WriteAllText(_settingsPath, JsonConvert.SerializeObject(ApiSettings)); } catch { } }
        private void LoadSettings() { if (File.Exists(_settingsPath)) { try { ApiSettings = JsonConvert.DeserializeObject<DictaRequest>(File.ReadAllText(_settingsPath)); } catch { } } }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    public class BooleanToVisibilityConverter : System.Windows.Data.IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture) => (bool)value ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture) => null;
    }
}