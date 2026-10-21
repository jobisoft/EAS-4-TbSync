# IMAP i kalendarz Microsoft 365 — kandydat 5.3.11.1583

Poczta zostaje na IMAP. Thunderbird obsługuje odpowiedzi i aktualizacje iTIP.
Microsoft 365 przechowuje powiązany termin. Nie przekazujemy maili do Exchange.

## Zmiany

- Jawna opcja konta „Zaproszenia z poczty IMAP”, domyślnie wyłączona.
  Działa w domyślnym kalendarzu Exchange (Type 8).
- Pełny iCalendar zachowuje źródłowy UID, organizatora, uczestników,
  SEQUENCE, DTSTAMP oraz RECURRENCE-ID. Trwały zapis źródła jest niezależny
  od pamięci podręcznej kalendarza i wiąże UID z ServerId.
- Tylko dane wysyłane do EAS mają postać terminu bez uczestników.
  Zapobiega to wysłaniu nowych zaproszeń w imieniu odbiorcy.
- Przyjęcie nowej wersji zaproszenia w Thunderbirdzie zmienia ten sam wpis
  Microsoft 365. Nie tworzy nowego UID ani drugiego terminu.
- REPLY obsługuje natywny transport SMTP, z tożsamością zgodną z uczestnikiem.
  Synchronizacja nie wysyła drugiego REPLY ani MeetingResponse dla tego wpisu.
- Włączenie opcji obejmuje już przyjęte, oczekujące zaproszenia bez ServerId,
  z potwierdzonym śladem iTIP. Nie wysyła zaległej odpowiedzi pocztowej.
- Zmiana i usunięcie z Microsoft 365 wracają do tego samego lokalnego wpisu.
  SoftDelete związany z oknem synchronizacji zachowuje źródło i kolejkę.
- Zwykła zmiana bez ServerId pozostaje w kolejce z ostrzeżeniem. Poprzedni
  kod upstream usuwał ją z kolejki. Nie tworzymy zastępczego duplikatu.
- Nieobsługiwany cykl nie jest dzielony ani skracany w zaproszeniu.

## Walidacja

- 534 testy jednostkowe: 533 PASS, 1 FAIL.
- Jedyny FAIL: reminder-ack.test.mjs, „a server change to something else does
  not ring the reminder again”. Ten sam błąd potwierdzono na czystej kopii
  bazowego commita 04c8f75e. Codec i ten test nie są zmienione.
- Testy obejmują zwykły Add, odrzucenie Add, nowy i istniejący mail REQUEST,
  zmianę tego samego ServerId, odpowiedź z właściwej tożsamości, cache reset,
  wystąpienia cyklu, SoftDelete, usunięcie z Teams i równoległy zapis z ACK.
- Dwa przeglądy Sol high: zgłoszone błędy poprawiono i sprawdzono ponownie.
- Thunderbird 158.0, osobny profil bez kont, offline: oba dodatki aktywne,
  bez błędów startowych; rejestracja dostawcy, zapis i czyszczenie terminu,
  natywny wybór transportu REPLY oraz interfejs transportu — PASS.
- Pozostaje znane ostrzeżenie manifestu o calendar_provider.
- Autor nie uruchamiał synchronizacji ani wysyłki na rzeczywistym koncie.
  Użytkownik zainstalował kandydata i wykonał własną próbę po włączeniu opcji.
  Dostarczony log potwierdza pobranie wpisu z Microsoft 365 z ClientUid
  źródłowego zaproszenia i ServerId. Zawiera też Change/Status 7 (konflikt),
  po którym serwer zwrócił swój wpis. Nie jest to dowód powodzenia aktualizacji.
  Nie potwierdzono jeszcze SMTP REPLY, wyższego SEQUENCE ani usunięcia z Teams.

## Granice

Źródłowy UID i dane zaproszenia są przechowywane w profilu Thunderbirda.
Zachowaj kopię profilu wraz z danymi rozszerzenia. Nowy, niezależny profil
nie odzyska tych danych z samego terminu Microsoft 365. W Outlooku/Teams
widać termin; przyjmowanie maili i odpowiedzi prowadzi Thunderbird.

Nie potwierdzono przyczyny konkretnego zwykłego wydarzenia zgłoszonego przez
użytkownika. Testy nie odtwarzają jego stanu konta ani odpowiedzi serwera.

## Baza i wycofanie

Gałąź fix/imap-invitation-lifecycle, baza lokalna
04c8f75e13b80fb3fd5e7a68ae38721aab924a2c; upstream
0d7638cfe286f0774cb9156de0fb1e3b0a94694e.
Lokalny XPI zawiera też wcześniejsze poprawki zgodności Thunderbirda 158.
TbSync pozostaje w wersji 5.3.8.158.

Przed instalacją wykonaj kopię profilu przy normalnie zamkniętym Thunderbirdzie.
Wycofanie: wyłącz nową opcję i przywróć poprzedni XPI. Nie usuwa to źródeł
zaproszeń. Powrót do starej wersji nie zapewnia obsługi nowych oznaczonych
zaproszeń; zachowaj pełną kopię profilu. Nie resetuj kalendarza ani konta.

PR #369 pozostaje zamknięty bez scalania. Nowy PR jest roboczy. Pełny cykl
odpowiedzi, aktualizacji i usunięcia w Microsoft 365 wymaga dalszego testu.
