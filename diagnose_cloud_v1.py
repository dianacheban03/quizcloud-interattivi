#!/usr/bin/env python3
"""
SCRIPT DI DIAGNOSTICA CLOUD
Controlla se Supabase è configurato correttamente
"""

import os
import sys
import json
from pathlib import Path

print("=" * 60)
print("🔍 DIAGNOSITICA CLOUD - QUIZ STORAGE")
print("=" * 60)

# ============================================================
# CHECK 1: File di configurazione
# ============================================================

print("\n1️⃣ CHECK: File di configurazione (.streamlit/secrets.toml)")
print("-" * 60)

secrets_path = Path(".streamlit/secrets.toml")
if not secrets_path.exists():
    print("❌ ERRORE: .streamlit/secrets.toml NON TROVATO")
    print("   Dove metterlo: Nella stessa cartella di app.py")
    print("   Nome esatto: .streamlit/secrets.toml (con il punto)")
    sys.exit(1)
else:
    print("✅ File trovato:", secrets_path.absolute())

# ============================================================
# CHECK 2: Contenuto del file
# ============================================================

print("\n2️⃣ CHECK: Contenuto di secrets.toml")
print("-" * 60)

try:
    import toml
    secrets = toml.load(str(secrets_path))
except ImportError:
    print("❌ ERRORE: La libreria 'toml' non è installata")
    print("   Installa con: pip install toml --break-system-packages")
    sys.exit(1)

required_keys = [
    "SUPABASE_URL",
    "SUPABASE_KEY",
    "CLOUD_BUCKET",
    "CLOUD_ROOT"
]

missing_keys = []
for key in required_keys:
    if key in secrets:
        value = secrets[key]
        # Nascondi i valori sensibili
        if len(value) > 20:
            masked = value[:10] + "***" + value[-5:]
        else:
            masked = "***"
        print(f"✅ {key}: {masked}")
    else:
        print(f"❌ MANCANTE: {key}")
        missing_keys.append(key)

if missing_keys:
    print(f"\n❌ ERRORE: Mancano le chiavi: {', '.join(missing_keys)}")
    print("\nEsempio di secrets.toml valido:")
    print("""
[default]
SUPABASE_URL = "https://tuoproggetto.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9..."
CLOUD_BUCKET = "quiz-storage"
CLOUD_ROOT = "quiz"
    """)
    sys.exit(1)

# ============================================================
# CHECK 3: Connessione a Supabase
# ============================================================

print("\n3️⃣ CHECK: Connessione a Supabase")
print("-" * 60)

try:
    from supabase import create_client, Client
    print("✅ Libreria supabase importata correttamente")
except ImportError:
    print("❌ ERRORE: La libreria 'supabase' non è installata")
    print("   Installa con: pip install supabase --break-system-packages")
    sys.exit(1)

try:
    client = create_client(secrets["SUPABASE_URL"], secrets["SUPABASE_KEY"])
    print("✅ Client Supabase creato")
    print(f"   URL: {secrets['SUPABASE_URL']}")
except Exception as e:
    print(f"❌ ERRORE nella creazione del client: {e}")
    sys.exit(1)

# ============================================================
# CHECK 4: Accesso al bucket
# ============================================================

print("\n4️⃣ CHECK: Accesso al bucket storage")
print("-" * 60)

bucket_name = secrets["CLOUD_BUCKET"]

try:
    # Prova a listare i file nel bucket
    files = client.storage.from_(bucket_name).list()
    print(f"✅ Bucket '{bucket_name}' accessibile")
    print(f"   File nel bucket: {len(files) if files else 0}")
    
    if files:
        print("\n   Struttura del bucket:")
        for item in files[:10]:  # Mostra primi 10
            print(f"   - {item['name']}")
        if len(files) > 10:
            print(f"   ... e altri {len(files) - 10} file")
    else:
        print("   (Bucket vuoto)")
        
except Exception as e:
    print(f"❌ ERRORE nell'accesso al bucket: {e}")
    print("\nPossibili cause:")
    print("1. Nome del bucket non corretto")
    print("2. Credenziali Supabase non valide")
    print("3. Bucket non esiste in Supabase")
    print("4. Mancano i permessi di accesso")
    sys.exit(1)

# ============================================================
# CHECK 5: Cartelle root
# ============================================================

print("\n5️⃣ CHECK: Cartelle nel cloud")
print("-" * 60)

cloud_root = secrets["CLOUD_ROOT"]
print(f"Cloud root: {cloud_root}")

try:
    # Prova a creare una cartella di test
    test_path = f"{cloud_root}/.health_check"
    test_data = b"health_check"
    
    client.storage.from_(bucket_name).upload(
        test_path,
        test_data,
        {
            "content-type": "text/plain"
        }
    )
    print(f"✅ Permessi di SCRITTURA: OK")
    
    # Leggi il file di test
    file_content = client.storage.from_(bucket_name).download(test_path)
    if file_content == test_data:
        print(f"✅ Permessi di LETTURA: OK")
    
    # Elimina il file di test
    client.storage.from_(bucket_name).remove([test_path])
    print(f"✅ Permessi di ELIMINAZIONE: OK")
    
    print("\n✅ Tutti i permessi sono OK!")
    
except Exception as e:
    print(f"❌ ERRORE nei permessi: {e}")
    print("\nDevi configurare i permessi nel bucket Supabase:")
    print("1. Vai su: Supabase Dashboard → Storage")
    print("2. Seleziona il bucket")
    print("3. Vai su 'Policies'")
    print("4. Aggiungi le policy per: SELECT, INSERT, UPDATE, DELETE")
    sys.exit(1)

# ============================================================
# CHECK 6: Struttura delle materie
# ============================================================

print("\n6️⃣ CHECK: Struttura delle materie nel cloud")
print("-" * 60)

try:
    files = client.storage.from_(bucket_name).list(cloud_root)
    
    if not files:
        print(f"⚠️  ATTENZIONE: Nessuna materia trovata in /{cloud_root}")
        print("\nPer creare una materia:")
        print("1. Apri l'app")
        print("2. Vai su 'Le mie materie'")
        print("3. Scrivi il nome (es: 'Bionanotecnologie')")
        print("4. Clicca '＋ Crea materia'")
    else:
        print(f"✅ Trovate {len(files)} materie:")
        for subject in files:
            print(f"   📁 {subject['name']}")
            
except Exception as e:
    print(f"❌ ERRORE: {e}")

# ============================================================
# CHECK 7: Test di salvataggio
# ============================================================

print("\n7️⃣ CHECK: Test salvataggio sessione")
print("-" * 60)

try:
    # Crea una cartella di test
    test_subject = "TEST_Diagnostica"
    test_quiz = "test_quiz"
    
    # Prova a salvare un file di test
    test_session_path = f"{cloud_root}/{test_subject}/sessions/{test_quiz}_sessione/test.json"
    test_data = json.dumps({
        "test": "success",
        "timestamp": str(__import__('datetime').datetime.now())
    }).encode('utf-8')
    
    client.storage.from_(bucket_name).upload(
        test_session_path,
        test_data,
        {
            "content-type": "application/json",
            "upsert": "true"
        }
    )
    print(f"✅ Test salvataggio: OK")
    print(f"   File salvato in: {test_session_path}")
    
    # Leggi il file
    content = client.storage.from_(bucket_name).download(test_session_path)
    print(f"✅ Test lettura: OK")
    
    # Elimina il file di test
    client.storage.from_(bucket_name).remove([test_session_path])
    print(f"✅ Test eliminazione: OK")
    
    print("\n✅ Salvataggio nel cloud funziona correttamente!")
    
except Exception as e:
    print(f"❌ ERRORE nel salvataggio: {e}")
    print("\nPossibili cause:")
    print("1. Permessi insufficienti nel bucket")
    print("2. Path non valido")
    print("3. Spazio insufficiente")
    print("4. Errore di rete")

# ============================================================
# RIEPILOGO
# ============================================================

print("\n" + "=" * 60)
print("✅ DIAGNOSTICA COMPLETATA")
print("=" * 60)

print("""
Se tutti i check sono ✅:
→ Il cloud è configurato correttamente
→ Prova a salvare una sessione dall'app

Se alcuni check sono ❌:
→ Leggi gli errori qui sopra
→ Correggi la configurazione
→ Riavvia lo script
""")

print("\nNote:")
print("- Usa questo script per capire perché il salvataggio non funziona")
print("- Salva l'output per ricordare i dettagli della tua configurazione")
print("- Se tutto OK ma l'app non salva lo stesso, leggi PATCH_BUG_FIX_SALVATAGGIO.md")
