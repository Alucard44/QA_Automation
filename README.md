# PPAP Automotive Quality Hub

Lokalna aplikacja webowa wspierająca przygotowywanie dokumentów PSW (Part Submission Warrant – dokument podsumowujący przedłożenie części klientowi do zatwierdzenia) w procesie PPAP/PPA.

Aplikacja pobiera dane z formularza lub edytowalnej tabeli, uzupełnia nimi szablon Excel, tworzy strukturę folderów według numeru części i wersji raportu, a następnie zapisuje dokumenty w formatach Excel i PDF.

> **Status projektu:** działające MVP obsługujące generowanie dokumentów PSW w formacie VDA.

> **Rezultat testu:** wygenerowanie 10 dokumentów zajęło średnio 2 minuty i 59 sekund. Ręczne przygotowanie takiej liczby dokumentów, przy założeniu około 10 minut na jeden dokument, zajmuje około 1 godziny i 40 minut. Oznacza to szacunkowe skrócenie czasu samego generowania o około 97%. Wynik nie obejmuje czasu przygotowania danych. [Szczegóły pomiaru](#test-wydajności).

## Cel projektu

Ręczne przygotowywanie dokumentów PSW wymaga wielokrotnego wprowadzania danych do arkuszy Excel. Przy większej liczbie części proces staje się czasochłonny, wymaga stałego skupienia i zwiększa ryzyko pomyłek.

Przygotowanie jednego dokumentu PSW zajmuje średnio około 10 minut. Już dla 10 numerów artykułów oznacza to około 100 minut pracy nad samą dokumentacją. W organizacjach zarządzających dużą liczbą części każda zmiana konstrukcyjna, procesowa lub dotycząca łańcucha dostaw (np. zmiana dostawcy komponentu) wymaga udokumentowania. Jeżeli zmiana dotyczy wielu numerów artykułów, odpowiednią dokumentację należy przygotować dla każdego z nich, przez co czas pracy rośnie proporcjonalnie.

Jest to czas netto, zakładający nieprzerwaną pracę w pełnym skupieniu. W rzeczywistych warunkach dochodzą również spotkania, korespondencja, przerwy i inne bieżące obowiązki. Powoduje to rozłożenie pracy na dłuższy okres oraz zwiększa ryzyko pomyłek podczas wielokrotnego wracania do dokumentacji.

Celem projektu jest skrócenie czasu potrzebnego na przygotowanie dokumentacji poprzez automatyczne przenoszenie danych do odpowiednich pól szablonu, generowanie dokumentów pojedynczo lub masowo oraz zapisywanie wyników w uporządkowanej strukturze folderów.

## Najważniejsze funkcje

* tryb pojedynczy z pełnym formularzem dla jednego numeru artykułu,
* tryb masowy, w którym dane wspólne wpisywane są w formularzu, a dane poszczególnych artykułów w edytowalnej tabeli,
* automatyczne mapowanie wprowadzonych danych do odpowiednich komórek szablonu Excel,
* zapisanie uzupełnionego dokumentu Excel oraz odpowiadającego mu pliku PDF w folderze utworzonym na podstawie numeru artykułu i wersji raportu,
* automatyczne tworzenie uporządkowanej struktury folderów,
* zabezpieczenie przed nadpisaniem wcześniej utworzonych plików,
* pomijanie pustych wierszy w trybie masowym,
* blokada przypadkowego zatwierdzenia formularza klawiszem Enter,
* wykorzystanie danych demonstracyjnych zamiast rzeczywistych danych firmowych i osobowych.

## Sposób działania

1. Użytkownik wybiera pozycję `VDA` z menu znajdującego się w panelu bocznym. Jest ona przeznaczona dla klientów korzystających ze standardu VDA, którzy nie posiadają własnej, dedykowanej formatki dokumentu.
2. Użytkownik wybiera pojedynczy lub masowy tryb generowania dokumentów PSW.
3. W trybie pojedynczym wszystkie dane jednego numeru artykułu wprowadzane są w pełnym formularzu.
4. W trybie masowym dane wspólne wprowadzane są w formularzu, a dane poszczególnych artykułów w edytowalnej tabeli.
5. Po uzupełnieniu danych użytkownik klika przycisk `Generuj Dokumentację`.
6. Aplikacja przygotowuje dane i wykorzystuje konfigurację JSON do przypisania ich do odpowiednich komórek szablonu.
7. Biblioteka `openpyxl` uzupełnia szablon i zapisuje dokument Excel.
8. Microsoft Excel, obsługiwany przez `pywin32`, eksportuje dokument do formatu PDF.
9. Uzupełniony dokument Excel oraz odpowiadający mu plik PDF są zapisywane w folderze utworzonym na podstawie numeru artykułu i wersji raportu.

## Technologie

* Python
* Streamlit
* pandas
* openpyxl
* pywin32
* JSON
* Git i GitHub

## Wymagania

Do uruchomienia aplikacji potrzebne są:

* system Windows,
* Python 3.11 lub nowszy,
* zainstalowany Microsoft Excel w wersji desktopowej,
* biblioteki wymienione w pliku `requirements.txt`,
* kompatybilny szablon dokumentu Excel.

Microsoft Excel jest wymagany do eksportowania dokumentów do formatu PDF. Mechanizm eksportu wykorzystuje interfejs COM dostępny w systemie Windows.

## Instalacja

### 1. Sklonowanie repozytorium

```powershell
git clone https://github.com/Alucard44/QA_Automation.git
cd QA_Automation
```

### 2. Utworzenie środowiska wirtualnego

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

### 3. Instalacja bibliotek

```powershell
python -m pip install -r requirements.txt
```

### 4. Dodanie szablonu Excel

Utwórz w katalogu projektu folder:

```text
templates
```

Następnie umieść w nim kompatybilny szablon pod nazwą:

```text
templates/vda_2020.xlsx
```

Szablon Excel nie jest częścią publicznego repozytorium. Został wykluczony ze względu na możliwe ograniczenia dotyczące jego dalszego udostępniania.

### 5. Uruchomienie aplikacji

W głównym katalogu projektu wykonaj:

```powershell
python -m streamlit run ui/app.py
```

Po uruchomieniu Streamlit otworzy aplikację w domyślnej przeglądarce internetowej.

## Obsługa aplikacji

Po uruchomieniu aplikacji wybierz pozycję `VDA` z menu znajdującego się w panelu bocznym.

### Tryb pojedynczy

1. Pozostaw wyłączoną opcję `Tryb masowej produkcji PSW`.
2. Uzupełnij dane organizacji i klienta, informacje o części oraz dane dotyczące dostarczanych próbek, takie jak ich liczba, numer partii i masa.
3. Zaznacz powód utworzenia raportu oraz zdarzenie, które spowodowało uruchomienie procedury PPA.
4. Kliknij przycisk `Generuj Dokumentację`.
5. Aplikacja utworzy dokumenty Excel i PDF.

### Tryb masowy

1. Włącz opcję `Tryb masowej produkcji PSW`.
2. Uzupełnij dane wspólne dla wszystkich dokumentów, między innymi dane organizacji, klienta, dostawy i próbek.
3. Dodaj w tabeli osobny wiersz dla każdego numeru artykułu i uzupełnij jego dane zmienne.
4. Kliknij przycisk `Generuj Dokumentację`.
5. Aplikacja pominie całkowicie puste wiersze i utworzy dokumenty dla wierszy zawierających numer artykułu.

## Zapisywanie dokumentów

Dokumenty są zapisywane poza głównym katalogiem repozytorium, zgodnie ze strukturą:

```text
vda/
└── <numer części>/
    └── <wersja raportu>/
        ├── PSW_<numer części>.xlsx
        └── PSW_<numer części>.pdf
```

Jeżeli plik o takiej samej nazwie już istnieje, aplikacja dodaje do nazwy bieżącą datę oraz kolejny numer, np. `_1`, `_2` lub `_3`. Dzięki temu wcześniej wygenerowane dokumenty nie zostają nadpisane.

## Struktura projektu

```text
QA_Automation/
├── config/
│   └── vda_config.json
├── services/
│   ├── pdf_generator.py
│   └── excel_generator.py
├── templates/
│   └── vda_2020.xlsx        # plik wymagany lokalnie, nieuwzględniony w repozytorium
├── ui/
│   ├── app.py
│   └── vda_view.py
├── .gitignore
├── requirements.txt
└── README.md
```

### Odpowiedzialność modułów

* `ui/app.py` – konfiguracja strony głównej i nawigacja aplikacji,
* `ui/vda_view.py` – formularz VDA, tabela trybu masowego i przygotowanie danych,
* `services/excel_generator.py` – mapowanie danych oraz generowanie dokumentu Excel,
* `services/pdf_generator.py` – eksport pierwszego arkusza dokumentu do PDF,
* `config/vda_config.json` – przypisanie danych aplikacji do komórek szablonu oraz określenie sposobu obsługi poszczególnych typów pól.

## Konfiguracja mapowania

Plik `vda_config.json` przechowuje adres komórki oraz sposób obsługi każdego pola. Podczas generowania aplikacja odczytuje te informacje z konfiguracji. Jeżeli układ szablonu Excel ulegnie zmianie, adresy pól można zaktualizować w pliku JSON bez modyfikowania głównej logiki generowania dokumentu.

Obsługiwane typy pól:

* `standard` – wpisanie standardowej wartości do komórki,
* `checkbox` – wpisanie znaku `X` i wyśrodkowanie go, jeżeli dana opcja została zaznaczona,
* `merged_checkbox` – obsługa pola wyboru znajdującego się w połączonych komórkach,
* `merged_multiline` – wpisanie tekstu wielowierszowego do połączonych komórek z zawijaniem tekstu i wyrównaniem do góry.

## Weryfikacja działania

Działanie MVP zostało sprawdzone manualnie dla następujących scenariuszy:

* wygenerowanie pojedynczego dokumentu Excel i PDF,
* wygenerowanie wielu dokumentów w trybie masowym,
* pomijanie pustych wierszy bez przerywania przetwarzania kolejnych danych,
* brak wartości `None` w wygenerowanych dokumentach,
* utworzenie osobnych folderów dla różnych wersji raportu,
* Jeżeli plik o takiej samej nazwie już istnieje, aplikacja dodaje do nazwy bieżącą datę oraz kolejny numer, np. `_1`, `_2` lub `_3`. Dzięki temu wcześniej wygenerowane dokumenty nie zostają nadpisane.
* zatwierdzenie formularza wyłącznie przez kliknięcie przycisku,
* Dodatkowo sprawdzono poprawność składni plików Python za pomocą `py_compile`.

Projekt nie posiada automatycznych testów.

## Test wydajności

Kontrolowany test trybu masowego obejmował dwukrotne wygenerowanie 10 dokumentów PSW w formatach Excel i PDF. Pierwszy pomiar wykonano bezpośrednio po ponownym uruchomieniu aplikacji, a drugi bez jej zatrzymywania.

| Test | Czas dla 10 dokumentów | Średni czas jednego dokumentu |
|---|---:|---:|
| Pierwszy przebieg po restarcie | 3 min 28 s | 20,8 s |
| Drugi przebieg bez restartu | 2 min 30 s | 15,0 s |
| **Średnia** | **2 min 59 s** | **17,9 s** |

Na podstawie średniego wyniku wygenerowanie 10 dokumentów zajęło 2 minuty i 59 sekund. Ręczne przygotowanie takiej liczby dokumentów, przy założeniu około 10 minut na jeden dokument, zajmuje około 1 godziny i 40 minut.

Oznacza to szacunkowe skrócenie czasu samego generowania dokumentów o około 97%. Porównanie nie obejmuje czasu potrzebnego na przygotowanie i wprowadzenie danych do tabeli.

Czas mierzono od kliknięcia przycisku generowania do wyświetlenia komunikatu o zakończeniu. Pliki były często dostępne wcześniej, ponieważ końcowy wynik obejmuje również eksport PDF, zamykanie programu Excel oraz zwalnianie zasobów. Rezultaty mogą zależeć od parametrów komputera, wersji programu Excel i aktualnego obciążenia systemu.

## Ograniczenia obecnej wersji

* aplikacja generuje obecnie wyłącznie dokument PSW,
* pełne generowanie wymaga lokalnego szablonu Excel,
* eksport PDF działa tylko w systemie Windows z zainstalowanym programem Microsoft Excel,
* obsługa klientów CNH, MAN i DAF jest planowana, a ich nazwy są już widoczne w nawigacji,
* aplikacja nie posiada jeszcze automatycznych testów ani pełnej walidacji wszystkich pól.

## Planowany rozwój

* przygotowanie autorskiego szablonu demonstracyjnego PSW,
* dodanie walidacji wymaganych pól i czytelnych komunikatów o błędach,
* przygotowanie testów automatycznych dla logiki aplikacji,
* dodanie dokumentu DIM (`Dimensional`), zawierającego raport wyników pomiarowych,
* dodanie dokumentu MAT (`Material`), zawierającego zestawienie materiałów i komponentów,
* połączenie modułu MAT z lokalną bazą materiałową w celu automatycznego uzupełniania informacji o komponentach i wskazywania brakujących danych,
* dodanie dokumentu `Self Assessment Production Process`, dotyczącego samooceny procesu produkcyjnego,
* dodanie dokumentu `Self Assessment Part`, dotyczącego samooceny części,
* dodanie dokumentu `Part History`, zawierającego historię zmian dotyczących części,
* dodanie dokumentu `Supply Chain`, dotyczącego informacji o łańcuchu dostaw,
* dodanie kolejnych konfiguracji klientów na podstawie ich dedykowanych formatek i wymagań.
* przeniesienie przygotowania danych i sterowania procesem generowania z pliku `vda_view.py` do osobnych modułów, aby ułatwić testowanie, utrzymanie kodu oraz dodawanie kolejnych dokumentów.
* przygotowanie osobnej dokumentacji testów manualnych, zawierającej przypadki testowe, wyniki oraz znalezione błędy,

## Wykorzystanie AI

Projekt rozwijałem z wykorzystaniem narzędzi AI jako wsparcia przy tworzeniu i analizowaniu kodu. Sam określiłem wymagania aplikacji i sposób działania procesu, wykonywałem testy, diagnozowałem znalezione błędy oraz weryfikowałem wprowadzane zmiany.

Projekt był dla mnie również praktyczną nauką pracy z Pythonem, Streamlit, Git oraz GitHubem.

## Informacja o danych

W publicznej wersji projektu wykorzystano wyłącznie dane demonstracyjne. Projekt nie zawiera rzeczywistych danych osobowych, danych projektowych ani poufnych danych przedsiębiorstwa.

Nazwy VDA, MAN, CNH i DAF zostały użyte wyłącznie jako odniesienia do kontekstu branżowego oraz planowanych konfiguracji klientów. Projekt nie jest oficjalnym narzędziem ani produktem wymienionych organizacji.
