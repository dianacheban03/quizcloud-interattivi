"""
Validatore per file DOCX di flashcard.

Controlla:
  1. Ogni card ha esattamente un separatore Q/A (- o -- o ---)
  2. Ogni card ha almeno una riga di domanda e una di risposta
  3. Le card sono separate da righe vuote
  4. Le domande numerate seguono una sequenza coerente
  5. Nessuna card è completamente vuota

Uso:
  python valida_flashcard.py percorso/del/file.docx
"""

import sys
import io
import re
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

try:
    from docx import Document
    from docx.oxml.ns import qn
except ImportError:
    print("ERRORE: installa python-docx con:  pip install python-docx")
    sys.exit(1)


SEP_RE = re.compile(r"^\s*[-=—–]{1,}\s*$")
Q_RE   = re.compile(r"^\s*(\d+)\.\s+")

VERDE  = "\033[92m"
GIALLO = "\033[93m"
ROSSO  = "\033[91m"
RESET  = "\033[0m"
GRASSETTO = "\033[1m"


def ha_immagine(p):
    return bool(p._p.findall(".//" + qn("a:blip")))


def solo_blank(p):
    return not p.text.strip() and not ha_immagine(p)


def colore(testo, c):
    return f"{c}{testo}{RESET}"


# ──────────────────────────────────────────────────────────────────────────────
def valida(path: Path):
    print(f"\n{GRASSETTO}Validazione: {path.name}{RESET}")
    print("─" * 60)

    try:
        doc = Document(str(path))
    except Exception as e:
        print(colore(f"ERRORE apertura file: {e}", ROSSO))
        return

    paras = doc.paragraphs

    # ── Raccolta card ──────────────────────────────────────────────────────
    # Una card = sequenza di paragrafi non-blank terminata da una riga vuota
    cards_raw = []   # lista di liste di paragrafi (una lista per card)
    current = []

    for p in paras:
        if solo_blank(p):
            if current:
                cards_raw.append(current)
                current = []
        else:
            current.append(p)
    if current:
        cards_raw.append(current)

    n_cards = len(cards_raw)
    print(f"Card trovate: {GRASSETTO}{n_cards}{RESET}")
    print()

    problemi = []

    # ── Analisi di ogni card ───────────────────────────────────────────────
    for idx, card_paras in enumerate(cards_raw):
        card_num = idx + 1
        errori_card = []

        # Testo completo e righe logiche (espandi soft-return)
        righe = []
        for p in card_paras:
            for sub in p.text.split("\n"):
                s = sub.strip()
                righe.append((s, ha_immagine(p)))

        # Conta separatori Q/A
        sep_indices = [i for i, (r, _) in enumerate(righe) if SEP_RE.match(r)]
        n_sep = len(sep_indices)

        if n_sep == 0:
            errori_card.append("nessun separatore Q/A (manca il trattino -)")
        elif n_sep > 1:
            errori_card.append(f"{n_sep} separatori Q/A (dovrebbe essercene 1)")

        # Righe domanda e risposta
        if sep_indices:
            primo_sep = sep_indices[0]
            q_righe = [(r, img) for r, img in righe[:primo_sep] if r or img]
            a_righe = [(r, img) for r, img in righe[primo_sep + 1:] if r or img]
        else:
            q_righe = [(r, img) for r, img in righe if r or img]
            a_righe = []

        if not q_righe:
            errori_card.append("domanda vuota (nessun testo prima del separatore)")
        if not a_righe:
            errori_card.append("risposta vuota (nessun testo/immagine dopo il separatore)")

        # Prima riga della domanda: è numerata?
        prima_riga = q_righe[0][0] if q_righe else ""
        m = Q_RE.match(prima_riga)
        num_domanda = int(m.group(1)) if m else None

        if errori_card:
            problemi.append((card_num, prima_riga[:60], errori_card))

    # ── Controllo numerazione ──────────────────────────────────────────────
    # Raggruppa le card per "sezione" (la numerazione riparte da 1 dopo
    # ogni interruzione discendente)
    numeri_per_sezione = []
    sezione_corrente = []
    ultimo_num = 0

    for idx, card_paras in enumerate(cards_raw):
        righe = []
        for p in card_paras:
            for sub in p.text.split("\n"):
                righe.append(sub.strip())

        prima = next((r for r in righe if r and not SEP_RE.match(r)), "")
        m = Q_RE.match(prima)
        if m:
            num = int(m.group(1))
            if num < ultimo_num:
                # nuova sezione
                numeri_per_sezione.append(sezione_corrente)
                sezione_corrente = []
            sezione_corrente.append((idx + 1, num, prima[:60]))
            ultimo_num = num
        else:
            sezione_corrente.append((idx + 1, None, prima[:60]))

    if sezione_corrente:
        numeri_per_sezione.append(sezione_corrente)

    problemi_num = []
    for sezione in numeri_per_sezione:
        numerati = [(ci, n, t) for ci, n, t in sezione if n is not None]
        for i in range(1, len(numerati)):
            ci_prev, n_prev, _ = numerati[i - 1]
            ci_curr, n_curr, t_curr = numerati[i]
            if n_curr != n_prev + 1:
                problemi_num.append(
                    (ci_curr, t_curr,
                     f"numerazione anomala: dopo domanda {n_prev} arriva domanda {n_curr}")
                )

    # ── Stampa risultati ───────────────────────────────────────────────────
    if not problemi and not problemi_num:
        print(colore("✓  Nessun problema trovato — il file è formattato correttamente.", VERDE))
    else:
        if problemi:
            print(colore(f"PROBLEMI STRUTTURALI ({len(problemi)} card):", ROSSO))
            for card_num, prima_riga, errori in problemi:
                print(f"\n  Card {card_num}: {colore(prima_riga or '(senza testo)', GIALLO)}")
                for e in errori:
                    print(f"    • {e}")

        if problemi_num:
            print()
            print(colore(f"PROBLEMI DI NUMERAZIONE ({len(problemi_num)}):", GIALLO))
            for card_num, prima_riga, msg in problemi_num:
                print(f"\n  Card {card_num}: {colore(prima_riga, GIALLO)}")
                print(f"    • {msg}")

    print()
    print("─" * 60)
    ok = n_cards - len(problemi)
    print(f"Riepilogo: {colore(str(ok), VERDE)} card OK  |  "
          f"{colore(str(len(problemi)), ROSSO if problemi else VERDE)} struttura  |  "
          f"{colore(str(len(problemi_num)), GIALLO if problemi_num else VERDE)} numerazione")
    print()


# ──────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python valida_flashcard.py file.docx [file2.docx ...]")
        print()
        print("Trascina uno o più file .docx sulla riga di comando.")
        sys.exit(0)

    for arg in sys.argv[1:]:
        p = Path(arg)
        if not p.exists():
            print(colore(f"File non trovato: {arg}", ROSSO))
            continue
        if p.suffix.lower() != ".docx":
            print(colore(f"Non è un file .docx: {arg}", GIALLO))
            continue
        valida(p)
