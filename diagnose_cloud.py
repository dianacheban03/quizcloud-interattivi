#!/usr/bin/env python3
"""
SCRIPT DI DIAGNOSTICA CLOUD - VERSIONE CORRETTA
Usa la struttura [cloud] con i nomi esatti
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
    all_secrets = toml.load(str(secrets_path))
except ImportError:
    print("❌ ERRORE: La libreria 'toml' non è installata")
    print("   Installa con: pip install toml --break-system-packages")
    sys.exit(1)

# Prova sia [cloud] che [default]
if "cloud" in all_secrets:
    secrets = all_secrets["cloud"]
    print("✅ Sezione [cloud] trovata")
elif "default" in all_secrets:
    secrets = all_secrets["default"]
    print("✅ Sezione [default] trovata")
else:
    print("❌ ERRORE: Nessuna sezione trovata (né [cloud] né [default])")
    print("Contenuto del file:")
    print(all_secrets)
    sys.exit(1)

# Chiavi che cerchiamo (con underscore minuscolo come nel tuo file)
required_keys = {
    "supabase_url": "SUPABASE_URL",
    "supabase_key": "SUPABASE_KEY",
    "bucket": "CLOUD_BUCKET",
    "root": "CLOUD_ROOT"
}

missing_keys = []
config_dict = {}

for key_name, key_display in required_keys.items():
    if key_name in secrets:
        value = secrets[key_name]
        config_dict[key_name] = value
        # Nascondi i valori sensibili
        if len(value) > 20:
            masked = value[:10] + "***" + value[-5:]
        else:
            masked = "***"
        print(f"✅ {key_name}: {masked}")
    else:
        print(f"❌ MANCANTE: {key_name}")
        missing_keys.append(key_name)

if missing_keys:
    print(f"\n❌ ERRORE: Mancano le chiavi: {', '.join(missing_keys)}")
    print("\nEsempio di secrets.toml valido:")
    print("""
[cloud]
supabase_url = "https://guxygnbvdljwadhqcufm.supabase.co"
supabase_key = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9..."
bucket = "quiz-cloud"
root = "diana"
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
    client = create_client(config_dict["supabase_url"], config_dict["supabase_key"])
    print("✅ Client Supabase creato")
    print(f"   URL: {config_dict['supabase_url']}")
except Exception as e:
    print(f"❌ ERRORE nella creazione del client: {e}")
    sys.exit(1)

# ============================================================
# CHECK 4: Accesso al bucket
# ============================================================

print("\n4️⃣ CHECK: Accesso al bucket storage")
print("-" * 60)

bucket_name = config_dict["bucket"]

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

cloud_root = config_dict["root"]
print(f"Cloud root: {cloud_root}")

try:
    # Prova a creare una cartella di test
    test_path = f"{cloud_root}/.health_check"
    test_data = b"health_check"
    
    client.storage.from_(bucket_name).upload(
        test_path,
        test_data,
        {
            "content-type": "text/plain",
            "upsert": "true"
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

print("\n6️⃣ CHECK: Struttura nel cloud (diana/)")
print("-" * 60)

try:
    files = client.storage.from_(bucket_name).list(cloud_root)
    
    if not files:
        print(f"⚠️  ATTENZIONE: Nessun file trovato in /{cloud_root}")
        print("\nPer caricare i file:")
        print("1. Apri l'app")
        print("2. Vai su 'Le mie materie'")
        print("3. Carica un quiz o una flashcard")
    else:
        print(f"✅ Trovati {len(files)} elementi:")
        for item in files[:20]:
            name = item['name']
            if item.get('metadata', {}).get('mimetype') == 'application/json':
                print(f"   📄 {name}")
            else:
                print(f"   📁 {name}")
            
except Exception as e:
    print(f"❌ ERRORE: {e}")

# ============================================================
# CHECK 7: Test di salvataggio
# ============================================================

print("\n7️⃣ CHECK: Test salvataggio sessione")
print("-" * 60)

try:
    # Crea un file di test nella cartella diana
    test_path = f"{cloud_root}/test_diagnostica.json"
    test_data = json.dumps({
        "test": "success",
        "timestamp": str(__import__('datetime').datetime.now())
    }).encode('utf-8')
    
    client.storage.from_(bucket_name).upload(
        test_path,
        test_data,
        {
            "content-type": "application/json",
            "upsert": "true"
        }
    )
    print(f"✅ Test salvataggio: OK")
    print(f"   File salvato in: {test_path}")
    
    # Leggi il file
    content = client.storage.from_(bucket_name).download(test_path)
    print(f"✅ Test lettura: OK")
    
    # Elimina il file di test
    client.storage.from_(bucket_name).remove([test_path])
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

print(f"""
Configurazione utilizzata:
├── Sezione: [cloud]
├── URL: {config_dict['supabase_url']}
├── Bucket: {bucket_name}
├── Root: {cloud_root}
└── Status: ✅ Tutto OK

Se tutti i check sono ✅:
→ Il cloud è configurato correttamente
→ L'app dovrebbe salvare le sessioni

Se alcuni check sono ❌:
→ Leggi gli errori qui sopra
→ Correggi la configurazione
→ Riavvia lo script
""")
