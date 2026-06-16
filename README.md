# Quiz Cloud Interattivi

Versione Streamlit con archivio cloud persistente.

## Struttura del cloud

```text
Materia/
├── quiz_originale.docx
└── quiz_originale_sessione/
    ├── corrette.docx
    ├── sbagliate.docx
    ├── non_fatte.docx
    └── sessione.json
```

## Funzioni incluse

- creazione delle cartelle per materia;
- caricamento dei quiz originali DOCX/PDF;
- apertura diretta di un quiz originale;
- apertura di una cartella sessione;
- scelta fra:
  - non fatte;
  - sbagliate;
  - corrette;
  - sbagliate + non fatte;
  - tutte;
- immagini collegate alle domande;
- feedback immediato;
- revisione opzionale del quiz;
- salvataggio cumulativo e senza duplicati;
- un unico pulsante per aggiornare i tre DOCX nel cloud;
- un unico backup ZIP scaricabile;
- soluzioni finali salvate solo come lettere (`1. A`);
- importazione locale di emergenza.

## 1. Installa i pacchetti

```powershell
py -3.12 -m venv .venv
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.venv\Scripts\activate
pip install -r requirements.txt
```

## 2. Configura Supabase

1. Crea un progetto Supabase.
2. Recupera:
   - Project URL;
   - Service Role Key.
3. Copia:

```text
.streamlit/secrets.example.toml
```

e rinominalo:

```text
.streamlit/secrets.toml
```

4. Inserisci URL e chiave.

La chiave Service Role deve restare segreta: il file `secrets.toml` è escluso da Git.

L'app prova a creare automaticamente un bucket privato chiamato `quiz-cloud`.
In alternativa, crealo manualmente da Supabase Storage.

## 3. Avvia

```powershell
streamlit run app.py
```

## Uso

1. Crea una materia.
2. Apri la materia.
3. Carica il quiz originale.
4. Premi `Avvia` sul file.
5. Svolgi le domande.
6. Premi `Salva/aggiorna tutto nel cloud`.

Alla sessione successiva apri la cartella `nomequiz_sessione`,
scegli quali domande svolgere e continua.

## Pubblicazione su Streamlit Community Cloud

Carica il progetto su GitHub senza `secrets.toml`.
Nelle impostazioni dell'app Streamlit inserisci gli stessi secrets:

```toml
[cloud]
supabase_url = "..."
supabase_key = "..."
bucket = "quiz-cloud"
root = "diana"
```
