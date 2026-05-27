# Exemple de fișe de disciplină

Două fișe reale pe care le poți folosi ca să încerci skill-ul fără să-ți pregătești propriile fișe în avans.

| Fișier | Disciplina | Structură | Credite |
|---|---|---|---|
| [`TI_25.S6.3-02_Programare _NET.docx`](TI_25.S6.3-02_Programare%20_NET.docx) | Programare .NET | Curs (2h) + Laborator (2h) | 3 |
| [`TI_25.S6.2_Securitatea sistemelor de calcul.docx`](TI_25.S6.2_Securitatea%20sistemelor%20de%20calcul.docx) | Securitatea sistemelor de calcul | Curs (2h) + Laborator (2h) | 4 |

Ambele provin din programul de studii **Tehnologia Informației**, semestrul 6, Universitatea Politehnica Timișoara.

## Cum le folosești

```bash
# Verificare batch a ambelor fișe
python3 ~/.claude/skills/fisa-disciplinei/scripts/verify.py examples/

# Verificare detaliată a uneia
python3 ~/.claude/skills/fisa-disciplinei/scripts/verify.py "examples/TI_25.S6.3-02_Programare _NET.docx"
```

Sau, prin Claude:

> *"Verifică fișele din folderul examples/."*

## La ce te poți aștepta

- **Programare .NET** — fișă complet consistentă (14/14 OK). Util ca să vezi cum arată un raport "totul în regulă".
- **Securitatea sistemelor de calcul** — conține intenționat **o eroare reală de rotunjire** la câmpul 3.8 (`7.1` în loc de `7.2`, fiindcă cineva a folosit ROUND obișnuit în loc de ROUNDUP cum cere norma). Util ca să vezi cum raportează skill-ul un MISMATCH și cum poți depista astfel de erori în fișele tale.

