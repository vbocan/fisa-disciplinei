# Validarea automată a fișelor de disciplină

[![Claude Code Skill](https://img.shields.io/badge/Claude%20Code-Skill-D97757?logo=anthropic&logoColor=white)](https://docs.claude.com/en/docs/claude-code/skills)
[![Python 3](https://img.shields.io/badge/Python-3.x-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![Format: .docx](https://img.shields.io/badge/format-.docx-2B579A?logo=microsoftword&logoColor=white)](#)
[![UPT / DPPD](https://img.shields.io/badge/șablon-UPT%20%2F%20DPPD-005BAA)](https://www.upt.ro/)
[![Limba: română](https://img.shields.io/badge/limba-română-CE1126)](#)
[![PRs welcome](https://img.shields.io/badge/PRs-welcome-brightgreen)](https://github.com/vbocan/fisa-disciplinei/pulls)

Un **skill pentru Claude Code** care verifică și actualizează **Fișa Disciplinei** în format Word (`.docx`), folosind regulile standard din șablonul UPT/DPPD. Conceput pentru cadrele didactice care vor să se asigure că aritmetica orelor și conținutul fișei sunt consistente cu materialele de curs și laborator.

## Ce face

- **Verifică aritmetica** secțiunii 3 ("Timp total estimat") — toate cele 14 formule canonice (3.1 = 3.2 + 3.3, 3.8\* = 3.9 × 25, etc.) sunt evaluate automat și raportate cu OK / MISMATCH.
- **Verifică consistența între secțiuni** — totalul orelor de curs din tabelul 8.1 trebuie să fie egal cu 3.2\*, iar totalul activităților aplicative din 8.2 trebuie să fie egal cu 3.3\*.
- **Modifică valorile** la cerere — schimbă ore săptămânale, credite, sau redistribuie orele pe capitole, și recalculează automat toate câmpurile derivate.
- **Funcționează pe loturi** — verifică toate fișele dintr-un folder cu o singură comandă; util la sfârșit de semestru când ai 5-10 fișe de pregătit.

## Cui se adresează

Cadrelor didactice de la UPT (sau orice universitate care folosește un șablon similar de fișă a disciplinei) care:

- Pregătesc/actualizează fișe de disciplină în fiecare an academic.
- Vor să se asigure că orele declarate (3.2\*, 3.3\*) corespund cu planul de curs (8.1) și laborator (8.2).
- Au făcut modificări la materialul de curs și vor să verifice că fișa reflectă noul conținut.

**Nu ai nevoie să fii dezvoltator software pentru a-l folosi.** Te ajuți de Claude care înțelege contextul și execută operațiile.

## Cum funcționează (la nivel înalt)

Skill-ul **nu este un program autonom** — este un **set de instrucțiuni pentru Claude Code** (un instrument AI care lucrează în terminalul tău). Atunci când îi spui lui Claude *"verifică fișa de disciplină X"*, el va detecta automat că trebuie să folosească acest skill și va:

1. Citi fișierul `.docx` și extrage valorile din secțiunea 3 și tabelele 8.1, 8.2.
2. Aplica cele ~14 formule de consistență.
3. Raporta într-un tabel ce e OK și ce nu.
4. Dacă vrei să modifici ceva, va edita direct fișierul `.docx` păstrând formatarea (font, culori, tabele).

Avantajul față de calculatorul Excel oficial (vezi [Originea regulilor](#originea-regulilor) mai jos) este că skill-ul **lucrează direct pe fișa ta în Word** — nu trebuie să copiezi/lipiți valori dintr-un fișier în altul, iar verificarea include și consistența cu tabelele 8.1 și 8.2, ceea ce calculatorul nu face.

## Originea regulilor

Regulile de calcul implementate în acest skill provin din [`Calculator ore pentru fișa de disciplină v4.xlsx`](Calculator%20ore%20pentru%20fi%C8%99a%20de%20disciplin%C4%83%20v4.xlsx) realizat pe baza normelor DPPD și a Cartei UPT. Fișierul este inclus în acest repository pentru:

- Formulele de derivare a câmpurilor (3.1 până la 3.9), cu detalii pe rotunjire (`ROUNDUP` pentru 3.8) și pe coeficienții de corecție pentru calculul creditelor.
- Algoritmul de calcul al numărului de credite pe baza timpului total de muncă (Tt), valorii unității de credit (Tt/Cs) și a coeficientului de corecție pe specializare (Ingineri: 0,61 | Arhitecți: 0,44–0,55).

Dacă găsești o discrepanță între ce face skill-ul și ce calculează Excel-ul, **Excel-ul este referința** — deschide un Issue, iar formulele din `scripts/verify.py` și `SKILL.md` vor fi corectate să se alinieze.

Folosește Excel-ul când:
- Vrei să **explorezi** ce s-ar întâmpla cu un curs nou înainte de a-l materializa într-o fișă (joacă-te cu numărul de credite, vezi cum se reflectă în orele neasistate).
- Vrei să **verifici manual** un calcul izolat pe care skill-ul nu îl face (de exemplu, coeficientul de corecție pentru o specializare diferită de inginerie).

Folosește skill-ul când:
- Ai deja fișa scrisă și vrei să **validezi automat** că tot ce e în ea respectă regulile.
- Ai mai multe fișe de verificat simultan (batch mode).
- Vrei să **modifici** o fișă existentă păstrând formatarea.

## Instalare

### Pasul 1 — Claude Code

Dacă nu ai deja Claude Code instalat:

```bash
# macOS (cu Homebrew)
brew install claude-code

# Sau descarcă de la
# https://claude.com/download
```

După instalare, deschide un terminal și rulează `claude` o dată pentru a-l autentifica cu contul tău Anthropic.

### Pasul 2 — Skill-ul

Clonează acest repository oriunde îți convine, apoi instalează doar subfolderul `skills/fisa-disciplinei` în folderul de skill-uri al Claude Code:

```bash
# Clonare repo (oriunde)
git clone <URL-ul-acestui-repo> ~/Repositories/fisa-disciplinei

# Instalare skill prin symlink (recomandat — primești actualizări cu `git pull`)
ln -s ~/Repositories/fisa-disciplinei/skills/fisa-disciplinei ~/.claude/skills/fisa-disciplinei

# Sau, dacă preferi copie (fără update automat)
# cp -r ~/Repositories/fisa-disciplinei/skills/fisa-disciplinei ~/.claude/skills/
```

Structura repo-ului:

```
fisa-disciplinei/                                     # repo clonat
├── README.md                                         # acest fișier (docs pentru oameni)
├── Calculator ore pentru fișa de disciplină v4.xlsx  # sursa de autoritate (Valer Bocan)
└── skills/
    └── fisa-disciplinei/                             # ← skill-ul propriu-zis (instalat în ~/.claude/skills/)
        ├── SKILL.md                                  # manualul pe care îl citește Claude
        └── scripts/
            └── verify.py                             # validatorul automat
```

### Pasul 3 — Dependențe Python

Scriptul de verificare are nevoie de Python 3 (deja instalat pe macOS). Nu necesită librării externe — doar biblioteca standard.

## Utilizare

Există două moduri: **prin Claude** (recomandat) sau **direct din terminal**.

### Modul recomandat — prin Claude

Pornește Claude într-un folder care conține fișele tale:

```bash
cd ~/Downloads
claude
```

Apoi spune-i pur și simplu, în limba română, ce vrei să facă. Exemple reale:

> *"Verifică fișa din `Programare .NET.docx`."*

> *"Schimbă numărul de credite la 4 în fișa MPS și recalculează toate câmpurile derivate."*

> *"Compară curricula din fișa SSC cu cursul predat de la `~/CURSURI/SSC/Curs.docx` și spune-mi unde nu se potrivesc."*

> *"Verifică toate fișele din folderul curent și fă-mi un sumar."*

Claude va folosi skill-ul automat — nu trebuie să specifici explicit nimic.

### Modul direct — terminal

Dacă preferi să rulezi scriptul fără Claude:

```bash
# Verificare o singură fișă (raport detaliat)
python3 ~/.claude/skills/fisa-disciplinei/scripts/verify.py "Fisa.docx"

# Verificare batch — toate fișele dintr-un folder
python3 ~/.claude/skills/fisa-disciplinei/scripts/verify.py ~/Downloads/

# Batch + detalii pentru fiecare
python3 ~/.claude/skills/fisa-disciplinei/scripts/verify.py ~/Downloads/ --verbose
```

Exemplu de output pe modul batch:

```
Verificare batch: 4 fișe în /Users/voi/Downloads

======================================================================
Sumar — 4 fișe verificate
======================================================================
Fișier                                                    Verif.  Status
---------------------------------------------------------------------------
CRO_25.S8.2-01_Managementul proiectelor software.docx      14/14  CONSISTENT
INFO_ZI_25.S3.7-02_Proiect sincretic I B.docx              14/14  CONSISTENT
TI_25.S6.2_Securitatea sistemelor de calcul.docx           13/14  1 discrepanțe
TI_25.S6.3-02_Programare _NET.docx                         13/14  1 discrepanțe

Consistente: 2/4
```

## Ce verifică, concret

Skill-ul aplică următoarele reguli (toate trebuie să fie respectate pentru ca fișa să fie consistentă):

| Regulă | Semnificație |
|---|---|
| `3.1 = 3.2 + 3.3` | Orele asistate/săptămână = curs + seminar/lab/proiect |
| `3.4 = 3.5 + 3.6` | Orele asistate parțial/săptămână = practică + proiect diplomă |
| `3.X* = 3.X × 14` | Toate valorile semestriale = valoarea săptămânală × 14 săpt |
| `3.1* = 3.2* + 3.3*` | Totalul asistat integral pe semestru |
| `3.4* = 3.5* + 3.6*` | Totalul asistat parțial pe semestru |
| `3.8* = 3.9 × 25` | Total ore/semestru = credite × 25 ore/credit (Carta UPT) |
| `3.8 = ROUNDUP(3.8*/14; 1)` | Total ore/săpt = total/sem ÷ 14, rotunjit în sus la 1 zecimală |
| `3.7* = 3.8* − 3.1* − 3.4*` | Orele neasistate = totalul minus partea asistată |
| `Σ ore în tabelul 8.1 = 3.2*` | Suma orelor pe capitole de curs = total curs/semestru |
| `Σ ore în tabelul 8.2 = 3.3*` | Suma orelor pe activități aplicative = total seminar/lab/semestru |

## Ce **nu** verifică (limitări)

- **Conținutul efectiv al cursului** — skill-ul nu poate ști dacă "Capitolul 3: Criptografie" predă efectiv ce spune fișa. Pentru asta poți cere lui Claude să compare fișa cu materialul de curs (vezi exemplul de mai sus); el va citi ambele și va raporta divergențe.
- **Bibliografia** — nu verifică dacă sursele citate există sau sunt actuale.
- **Coerența pedagogică** — nu evaluează dacă distribuția orelor pe capitole are sens didactic.
- **Alte secțiuni decât 3, 8.1 și 8.2** — secțiunile 1, 2, 4-7, 9 (date administrative, precondiții, obiective, evaluare) nu sunt validate automat.

În toate aceste cazuri, Claude poate face validarea calitativă dacă îi dai materialul de comparație — dar nu există formulă matematică universală pentru aceste verificări, deci nu sunt în scriptul automat.

## Cazuri limită cunoscute

- **3.7 = 3.8 − 3.1 − 3.4** poate da MISMATCH din cauza rotunjirii lui 3.8 (`ROUNDUP` introduce o eroare de până la 0.1 ore/săpt × 14 = 1.4 ore/sem). Verificarea semestrială echivalentă, `3.7* = 3.8* − 3.1* − 3.4*`, este "curată" și se bazează pe numere întregi. Dacă semestrialul e OK, ignoră săptămânalul.
- În unele fișe, câmpurile 3.7 și 3.8 (săptămânal) nu sunt detectate automat — scriptul afișează `—` pentru ele. Aritmetica este verificată oricum prin echivalentele semestriale.
- Notele de subsol din șablon (cifrele superscript de lângă etichete, ex. "Categoria formativă⁴") sunt filtrate automat. Dacă apare totuși o eroare bizară, deschide fișa și asigură-te că valoarea numerică din celulă este într-un run de text normal, nu superscript.

## Cum să modifici skill-ul

`SKILL.md` este "manualul" pe care Claude îl citește când activează skill-ul. Conține:

- Lista formulelor de consistență.
- Workflow-ul pe care îl urmează când e rugat să modifice o fișă.
- Capcanele pe care le evită.

Dacă șablonul UPT se schimbă (de exemplu Carta schimbă "1 credit = 25 ore" în "1 credit = 26 ore"), modifică `SKILL.md` și `scripts/verify.py` în consecință.

## Cum să contribui

Acest skill a fost creat colaborativ și e gândit să crească. Dacă găsești:

- O regulă pe care fișa ar trebui să o respecte și nu e încă în script,
- Un șablon de fișă pe care scriptul nu îl parsează corect,
- O nouă variantă a fișei (de exemplu pentru ciclul Master),

deschide un **Issue** sau trimite un **Pull Request**. Cele mai utile contribuții includ exemple de fișiere `.docx` care au declanșat problema, ca să poată fi reproduse.

## Întrebări frecvente

**Î: Trebuie să trimit fișa la Anthropic ca să o verifice?**
R: Nu. Tot procesul rulează local pe calculatorul tău. Claude Code citește fișa și o procesează local; doar întrebarea în limba română (ex. "verifică această fișă") este trimisă către serverele Anthropic. Conținutul fișierului `.docx` poate fi citit de Claude prin tool-uri de filesystem, dar acel conținut nu se păstrează după ce conversația se termină.

**Î: Pot folosi același skill pentru fișe în engleză sau pentru alte universități?**
R: Scriptul caută etichete specifice în limba română (ex. "3.1 Număr de ore asistate integral/săptămână"). Pentru alte universități/limbi, ar trebui adaptate șirurile căutate în `verify.py` și formulele din `SKILL.md` (de exemplu, valoarea credit-ului în ore poate diferi).

**Î: Funcționează și pentru fișe în format `.doc` (vechi) sau `.pdf`?**
R: Nu direct. Skill-ul lucrează doar cu `.docx`. Pentru `.doc` vechi, deschide în Word și salvează ca `.docx`. Pentru PDF, ar trebui regenerat din docx-ul original (PDF nu poate fi editat preservând formatarea).

**Î: Ce facem dacă Claude propune o modificare cu care nu sunt de acord?**
R: Refuzi pur și simplu — "nu, păstrează valoarea inițială" sau "altă variantă: …". Claude lucrează interactiv, nu autonom, și se va ajusta. De asemenea, înainte de fiecare modificare a fișei tale, există un backup automat în `/tmp/` la care poți reveni.

## Licență

MIT. Folosește, copiază, modifică, redistribuie fără restricții.

## Autor

**[Valer Bocan, PhD, CSSLP](https://www.bocan.ro)** — Universitatea Politehnica Timișoara, Facultatea de Automatică și Calculatoare. Autor al calculatorului Excel original (`Calculator ore pentru fișa de disciplină v4.xlsx`) și al acestui skill.

Codul este open source — îmbunătățiri și adaptări pentru alte universități sunt binevenite.
