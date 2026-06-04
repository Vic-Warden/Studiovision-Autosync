import win32com.client
import pythoncom
import sys

def fill_refraction_form():
    # Initialisation de COM
    pythoncom.CoInitialize()
    
    print("Tentative de connexion à Studio Vision (Access)...")
    try:
        # Se connecter à l'instance Access déjà ouverte
        access = win32com.client.GetActiveObject("Access.Application")
    except Exception as e:
        print(f"❌ Erreur de connexion. Assurez-vous que Studio Vision est ouvert. Détails: {e}")
        sys.exit(1)

    try:
        # Cibler spécifiquement le formulaire de réfraction
        form = access.Forms("REFRACTION")
        print(f"✅ Formulaire '{form.Name}' trouvé !")
    except Exception as e:
        print("❌ Le formulaire 'REFRACTION' n'est pas ouvert ou n'est pas accessible.")
        sys.exit(1)

    # Dictionnaire des valeurs à remplir. 
    # Remplace les chaînes de caractères par les vraies valeurs cliniques à tester.
    valeurs_a_injecter = {
        "TOD": "14",           # Tension OD
        "Champ31": "15",       # Tension OG (Nommé Champ31 dans ton Access)
        "SPHERE OD": "+1.25",  
        "SPHERE OG": "+1.50",
        "CYLINDRE OD": "-0.25",
        "CYLINDRE OG": "-0.50",
        "AXE OD": "90",
        "AXE OG": "85",
        "AVL OD": "10",        # Acuité visuelle de loin
        "AVL OG": "10",
        "ADD OD": "2.50",      # Addition
        "ADD OG": "2.50",
        "AVP OD": "P2",        # Parinaud
        "AVP OG": "P2",
        "Binoc": "10"
    }

    print("\n--- Début du remplissage ---")
    # Boucle d'injection
    for nom_champ, nouvelle_valeur in valeurs_a_injecter.items():
        try:
            # On assigne la nouvelle valeur au contrôle correspondant
            form.Controls(nom_champ).Value = nouvelle_valeur
            print(f"  [OK] {nom_champ:<12} -> {nouvelle_valeur}")
        except Exception as e:
            print(f"  [Erreur] Impossible de remplir '{nom_champ}' : {e}")

    print("--- Remplissage terminé ---")

if __name__ == "__main__":
    fill_refraction_form()