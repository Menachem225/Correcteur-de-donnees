import pandas as pd
from rapidfuzz import process, fuzz

# Charger le fichier Excel
df = pd.read_excel('correcteur_excel.xlsx')  # Remplacez par le chemin réel de votre fichier

# Nettoyer les données : convertir en chaînes et remplacer les NaN par des chaînes vides
df['Nom_Incomplet'] = df['Nom_Incomplet'].astype(str).fillna('')
df['Nom_Complet'] = df['Nom_Complet'].astype(str).fillna('')

# Créer une liste de noms bien formés
noms_complets = df['Nom_Complet'].tolist()

# Fonction pour trouver le nom complet le plus proche
def trouver_nom_proche(nom_incomplet):
    if not nom_incomplet.strip():  # Vérifier si la chaîne est vide
        return ''
    match = process.extractOne(nom_incomplet, noms_complets, scorer=fuzz.ratio)
    return match[0] if match else nom_incomplet

# Appliquer la fonction pour créer une nouvelle colonne avec les noms corrigés
df['Nom_Corrige'] = df['Nom_Incomplet'].apply(trouver_nom_proche)

# Sauvegarder le fichier Excel avec les résultats
df.to_excel('fichier_corrige.xlsx', index=False)

print("Les noms ont été corrigés et le fichier a été sauvegardé sous 'fichier_corrige.xlsx'")
