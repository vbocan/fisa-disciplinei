---
name: fisa-disciplinei
description: Actualizează și verifică fișele de disciplină UPT (DPPD) în format .docx — secțiunea 3 ("Timp total estimat"). Folosește această skill ori de câte ori utilizatorul deschide, modifică sau cere verificarea unui fișier numit "Fișa disciplinei", "FISA_DISCIPLINEI", sau orice fișier .docx care urmează șablonul UPT pentru fișa unei discipline, inclusiv când utilizatorul cere doar să schimbe credite, ore curs, ore seminar/laborator/proiect, ore neasistate, sau să verifice consistența acestor câmpuri. Triggere și pentru "fișă disciplină", "ore pe semestru", "calculator ore fișă", "fișa Bocan", sau când utilizatorul indică un fișier docx ce conține antetul "FIŞA DISCIPLINEI" cu rubrici 3.1-3.9.
---

# Fișa disciplinei — actualizare și verificare

## Când să folosești

- Utilizatorul cere modificarea unui câmp din secțiunea 3 a fișei (ore/săptămână, ore/semestru, credite, sub-componente neasistate).
- Utilizatorul cere o verificare a consistenței fișei (chiar fără modificări).
- Există ca referință un "calculator ore" în Excel — formulele de mai jos îl reproduc; **nu** este nevoie să citești Excelul de fiecare dată, formulele sunt deja codificate aici.

## Modelul de calcul (din `Calculator ore pentru fișa de disciplină v4.xlsx`)

Secțiunea 3 are 9 grupuri de câmpuri. Câmpurile cu `*` sunt valori semestriale; celelalte sunt săptămânale. Doar câteva sunt **intrări**; restul se **derivă**.

### Intrări (cele pe care utilizatorul le decide)
- `3.2` — ore curs / săptămână
- `3.3` — ore seminar/laborator/proiect / săptămână
- `3.5` — ore practică / săptămână (de obicei 0)
- `3.6` — ore elaborare proiect de diplomă / săptămână (de obicei 0)
- `3.9` — număr de credite
- Sub-componentele lui `3.7` și `3.7*` (cum se distribuie orele neasistate)

### Derivări (formulele canonice)

```
3.1   = 3.2 + 3.3
3.4   = 3.5 + 3.6
3.2*  = 3.2 × 14
3.3*  = 3.3 × 14
3.5*  = 3.5 × 14
3.6*  = 3.6 × 14
3.1*  = 3.2* + 3.3*
3.4*  = 3.5* + 3.6*
3.8*  = 3.9 × 25                     # 1 credit = 25 ore
3.8   = ROUNDUP(3.8* / 14, 1)        # rotunjire în sus la 1 zecimală
3.7   = 3.8 - 3.1 - 3.4
3.7*  = 3.8* - 3.1* - 3.4*           # preferă această formulă semestrială
                                      # (Excelul folosește 3.7 × 14 dar dă 0.4h în plus
                                      #  din cauza ROUNDUP-ului)
sum(sub-componente 3.7)  = 3.7
sum(sub-componente 3.7*) = 3.7*

# Consistență între secțiunea 3 și secțiunea 8 (planul de conținut):
sum(orele din tabelul 8.1 Curs)                = 3.2*
sum(orele din tabelul 8.2 Activităţi aplicative) = 3.3*
```

**Observație importantă despre rotunjire**: `3.7* = 3.7 × 14` și `3.7* = 3.8* − 3.1* − 3.4*` pot diferi cu până la ~0.4 ore din cauza `ROUNDUP` în `3.8`. Versiunea semestrială (`3.8* − 3.1* − 3.4*`) este "curată" pentru că provine din valori întregi. Folosește această versiune când stabilești `3.7*` și sub-componentele sale; folosește `3.7 × 14` doar ca avertisment dacă există discrepanță.

## Workflow de modificare

1. **Despachetează** docx-ul: `python skills/docx/scripts/office/unpack.py file.docx unpacked/`
2. **Localizează** câmpurile în `unpacked/word/document.xml`. Etichetele "3.1", "3.2*" etc. apar literal în text — caută-le și valoarea apare în următorul `<w:t>` (sau două).
3. **Modifică doar intrările** identificate cu utilizatorul (vezi tabelul de mai sus). Folosește unealta Edit, nu scripturi.
4. **Recalculează** toate câmpurile derivate aplicând formulele și actualizează-le în docx.
5. **Verifică** cu `scripts/verify.py` înainte de re-împachetare.
6. **Re-împachetează**: `python skills/docx/scripts/office/pack.py unpacked/ output.docx --original file.docx`.
7. **Confirmă cu utilizatorul** ce s-a schimbat — afișează un tabel before/after pentru toate câmpurile afectate.

## Workflow de verificare (fără modificări)

1. Despachetează docx-ul.
2. Extrage toate valorile din secțiunea 3.
3. Rulează `scripts/verify.py /cale/spre/fișa.docx`.
4. Raportează în tabel: pentru fiecare formulă, valoarea din document vs. valoarea calculată, cu OK / MISMATCH.
5. **Nu emite avertismente pentru diferența `3.7* = 3.7 × 14` vs. `3.7* = 3.8* − 3.1* − 3.4*`** dacă diferența este sub 0.5 ore — este o consecință cunoscută a `ROUNDUP`. Avertizează numai dacă suma sub-componentelor `3.7*` nu coincide cu valoarea declarată a lui `3.7*`.

## Stiluri și conservare a formatului

- Documentul are formatare bogată (tabele, fonturi, paragrafe italice). Nu rescrie blocuri `<w:r>` întregi — modifică doar `<w:t>` care conține valoarea numerică. Asta păstrează `<w:rPr>` și orice formatare existentă.
- Dacă o valoare apare în mai multe câmpuri (de ex. "28" la 3.1* și 3.3* într-un curs fără ore de curs), folosește contextul XML din jur ca să țintești câmpul corect. Caută eticheta "3.X*" și apoi prima valoare numerică care urmează.

## Erori frecvente de evitat

- **Confundarea săptămânal cu semestrial**: utilizatorul spune "28 de ore" — întreabă dacă este 3.1\*, 3.3\* sau 3.8\*. Aceste valori coexistă în același document.
- **Modificarea derivatelor fără să schimbi intrările**: dacă schimbi 3.3\* la 42, trebuie să schimbi 3.3 la 3 (42/14). Niciodată nu seta o valoare derivată direct; setează intrarea și recalculează.
- **Sub-componentele 3.7\* nu se actualizează automat**: utilizatorul le distribuie după preferință. La modificare, întreabă explicit cum vrea distribuția.
- **Credite vs. ore**: `3.8* = 3.9 × 25`. Dacă utilizatorul cere 50 ore/semestru, credite = 2. Dacă cere un număr de ore care nu împarte exact la 25, semnalează inconsistența — Carta UPT cere credite întregi.
- **Secțiunea 8 trebuie să totalizeze secțiunea 3**: când modifici 3.2 sau 3.3 (deci implicit 3.2\* / 3.3\*), planul de curs (8.1) și planul de activități aplicative (8.2) trebuie redistribuite să totalizeze noile valori. Skill-ul nu poate ghici cum să redistribuie orele între teme — întreabă utilizatorul ce activități adaugă/scoate pentru a închide diferența.

## Scripturi

- `scripts/verify.py <file.docx>` — verifică o singură fișă; afișează tabel detaliat cu OK/MISMATCH pentru toate formulele și sumele 8.1/8.2.
- `scripts/verify.py <folder>` — **mod batch**: scanează recursiv folderul și verifică toate `.docx`-urile găsite; afișează un sumar de tipul `Fișier | Verificate | Status`.
- `scripts/verify.py <folder> --verbose` — batch + raport detaliat pentru fiecare fișă.

Scriptul tratează corect notele de subsol marcate ca `vertAlign="superscript"` (UPT folosește acest pattern în loc de stilul standard `FootnoteReference`) și etichetele de secțiune concatenate fără spațiu cu valorile.

## Exemplu — verificare reușită

```
$ python scripts/verify.py "INFO_ZI_25.S3.7-02_Proiect sincretic.docx"

Câmpuri intrare:
  3.2  (curs/săpt):       0
  3.3  (sem/săpt):        2
  3.9  (credite):         2

Verificare:
  3.1  = 3.2 + 3.3        2    = 2    OK
  3.3* = 3.3 × 14         28   = 28   OK
  3.1* = 3.2* + 3.3*      28   = 28   OK
  3.8* = 3.9 × 25         50   = 50   OK
  3.8  = ROUNDUP(3.8*/14) 3.6  = 3.6  OK
  3.7* = 3.8* − 3.1*      22   = 22   OK
  Σ sub-3.7*              22   = 22   OK

Fișa este consistentă.
```
