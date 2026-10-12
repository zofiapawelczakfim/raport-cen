# Automat raportu cen mieszkań → dane.gov.pl

Darmowy zamiennik płatnego raportowania (art. 19b ustawy deweloperskiej, „jawność cen”).
Ty aktualizujesz jeden arkusz Google, a cała reszta dzieje się sama, codziennie, także w weekendy i święta.

```
Arkusz Google (Ty zmieniasz ceny)
        │  codziennie ~01:17–02:47 w nocy
        ▼
GitHub Actions (darmowy automat)
  • sprawdza arkusz
  • zapisuje datę każdej zmiany ceny
  • tworzy CSV w formacie ministerstwa (58 kolumn, puste pola = „X”)
  • aktualizuje plik XML + MD5
        │
        ▼
GitHub Pages (darmowy hosting): https://TWOJ-LOGIN.github.io/raport-cen/<id>.xml
        │  dane.gov.pl sam pobiera ten plik raz na dobę (Importer XML)
        ▼
dane.gov.pl: codziennie nowy plik z cenami w Twoim zbiorze danych
```

Koszt: **0 zł**. Repozytorium musi być publiczne, ale publikowane są wyłącznie dane, które i tak są jawne z mocy ustawy.

---

## Wdrożenie krok po kroku (ok. 1–2 h + czas oczekiwania na dane.gov.pl)

### Krok 1. Arkusz z cennikiem
1. Wgraj `szablon/Cennik_szablon.xlsx` na Google Drive, otwórz go i wybierz **Plik → Zapisz jako Arkusze Google**.
2. Wypełnij zakładki **Firma**, **Inwestycje** i **Oferta**. Zielone wiersze to przykład: usuń je. Szczegóły są w zakładce *Instrukcja*.
   - W polu **Identyfikator techniczny** wpisz krótką nazwę bez spacji i polskich znaków, np. `nazwa-firmy`. **Po starcie tego nie zmieniaj.**
   - Przenosisz ofertę od Raportuje Ceny? W kolumnie *Ceny obowiązują od* wpisz datę, od kiedy obowiązuje obecna cena. Weź ją z ostatniego pliku na dane.gov.pl. Wypełniasz tylko raz.
3. **Udostępnij → Ogólny dostęp: „Każda osoba mająca link” → Przeglądający.**
   Trzymaj w tym pliku tylko cennik, bo link daje dostęp do całego pliku.
4. Skopiuj **ID arkusza**, czyli fragment adresu między `/d/` a `/edit`:
   `https://docs.google.com/spreadsheets/d/`**`1AbC...xyz`**`/edit`

### Krok 2. Konto i repozytorium GitHub
1. Załóż darmowe konto na github.com, najlepiej na firmowy e-mail. Na ten adres będą przychodzić alerty.
2. **New repository** → nazwa `raport-cen` → **Public** → *Create*.
3. **Add file → Upload files** → przeciągnij foldery `generator`, `szablon`, `testy` oraz pliki `README.md`, `MAILE.md`, `.gitignore` → *Commit changes*.
4. Plik automatu dodaj osobno, bo folder `.github` bywa ukryty na Macu i w Windows:
   **Add file → Create new file** → w nazwie wpisz dokładnie `.github/workflows/codzienny-raport.yml` → wklej zawartość tego pliku z paczki → *Commit changes*.

### Krok 3. Ustawienia repozytorium
1. **Settings → Pages → Build and deployment → Source: „GitHub Actions”.**
2. **Settings → Secrets and variables → Actions → zakładka Variables → New repository variable**
   - Name: `GOOGLE_SHEET_ID`, Value: ID z kroku 1.4.
   - *(opcjonalnie)* `PAGES_URL`: tylko jeśli podepniesz własną domenę, np. `https://dane.twojafirma.pl`.
3. **Settings → Actions → General → Workflow permissions: „Read and write permissions”** → *Save*.

### Krok 4. Pierwsze uruchomienie
1. Zakładka **Actions** → *Codzienny raport cen (dane.gov.pl)* → **Run workflow**.
2. Po ok. 1 minucie powinna się pojawić zielona ikonka. Jeśli jest czerwona, kliknij w nią, a w logu będzie po polsku napisane, co poprawić w arkuszu.
3. Otwórz `https://TWOJ-LOGIN.github.io/raport-cen/`. Zobaczysz link do pliku XML i dzisiejszego CSV. **Skopiuj adres XML.**

### Krok 5. Wniosek do dane.gov.pl (jednorazowo)
Wyślij z konta, które ma uprawnienia edytora Twojego profilu dostawcy (wzór maila jest w `MAILE.md`).
Portal sprawdzi plik i ustawi codzienne pobieranie. Może poprosić o wypełnienie formularza
„Dane niezbędne do uruchomienia automatyzacji procesu udostępniania danych”.

### Krok 6. Odejście od Raportuje Ceny (bez dziury w raportowaniu)
1. **Nie wypowiadaj umowy, dopóki dane.gov.pl nie potwierdzi uruchomienia automatu** i nie zobaczysz w swoim profilu 2–3 kolejnych dni z nowego źródła.
2. Przez te kilka dni obie drogi działają równolegle. To nie problem, a lepiej mieć dane podwójnie niż mieć dzień przerwy.
3. Potem: wypowiedzenie umowy (sprawdź okres wypowiedzenia) + mail do dane.gov.pl o **odebraniu uprawnień edytora** Raportuje Ceny (wzór w `MAILE.md`).
4. Historia z poprzednich miesięcy zostaje na portalu. W mailu prosimy, żeby administratorzy jej nie usuwali.

---

## Codzienna obsługa
- **Zmiana ceny:** zmieniasz liczbę w arkuszu. Datę zmiany automat zapisze sam.
- **Sprzedaż lub rezerwacja:** usuwasz wiersz z zakładki *Oferta*.
- **Nowa inwestycja:** dodajesz wiersz w *Inwestycje* i lokale w *Oferta*.
- **Coś nie działa:** GitHub wyśle maila „Run failed”. Najczęściej chodzi o literówkę w arkuszu, a log mówi, w którym wierszu.
  Jeśli arkusz Google był chwilowo niedostępny, automat i tak opublikuje kopię ostatniego cennika i wyśle alert.
- Raz na jakiś czas zajrzyj na swój profil na dane.gov.pl i sprawdź, czy przybywają dzienne pliki.

## Ważne zasady
- **Nie zmieniaj** identyfikatora technicznego ani nazwy repozytorium. Zmiana adresu = nowy zbiór na portalu. Automat blokuje zmianę identyfikatora.
- **Nie kasuj** plików z `docs/csv/` ani `stan/`. Importer usuwa z portalu wszystko, czego nie ma w XML.
- Nazwy zakładek i nagłówki w 1. wierszu arkusza muszą zostać takie jak w szablonie.

## Czego ten automat NIE robi
Ustawa ma dwa osobne obowiązki:
1. **art. 19b**: codzienne przekazywanie danych do dane.gov.pl. **To robi ten automat.**
2. **art. 19a**: ceny na **Twojej stronie WWW** z historią zmian i datami. Na stronie musi być też m.in. prospekt informacyjny.

Sprawdź, czy Raportuje Ceny nie robi Ci przypadkiem czegoś na stronie WWW (widżet, historia cen). Jeśli tak, trzeba to zastąpić osobno.

## Pliki
| Ścieżka | Co to |
|---|---|
| `generator/generuj.py` | cała logika (Python, tylko biblioteka openpyxl) |
| `generator/naglowki_ministerstwo.json` | 58 nagłówków z „Wzorcowego zakresu danych” MC |
| `.github/workflows/codzienny-raport.yml` | harmonogram (2× w nocy), publikacja na Pages |
| `szablon/Cennik_szablon.xlsx` | arkusz do uzupełnienia |
| `testy/test_generator.py` | test: 3 dni, zmiana ceny, sprzedaż, przełom roku, poprawność XML/MD5 |
| `docs/` | (tworzy się sam) to, co widzi dane.gov.pl |
| `stan/` | (tworzy się sam) historia cen i dat zmian |

Test lokalny: `pip install openpyxl && python testy/test_generator.py`

## Zgodność z wytycznymi
- Struktura XML: przestrzeń nazw `urn:otwarte-dane:harvester:1.13`, wartości zgodne z „Instrukcją przygotowania pliku XML” (COI, wersje 1.0.3 i 1.0.6):
  `ECON`, `daily`, `availability=local`, znak umowny `X`, flagi `hasHighValueData=true` itd.
- MD5: małe litery, w pliku o tej samej nazwie co XML (`.md5` zamiast `.xml`).
- GitHub Pages serwuje `.xml` jako `application/xml` na porcie 443, czyli tak, jak wymaga walidator portalu.
- CSV: UTF-8, przecinek, jeden wiersz nagłówka, bez scalonych komórek, daty `RRRR-MM-DD`, NIP/REGON/KRS jako same cyfry, kwoty brutto.
