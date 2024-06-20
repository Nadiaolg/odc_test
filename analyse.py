import pandas as pd
import openpyxl
import matplotlib.pyplot as plt
import seaborn as sns


# 1 LECTURE DU FICHIER #

# Chargeons le fichier Excel de type xlsx en DataFrame
df = pd.read_excel('Online Retail.xlsx', engine='openpyxl')

#Nous devons avoir les informations sur les données du df
df.info()


# Afficheons les premières lignes du DataFrame pour vérifier ce qu'il contient et que nous l'avons effectivement chargé
print(df.head())


# Vérification les valeurs manquantes dans chaque colonne
missing_values = df.isnull().sum()

#Afficher les colonnes avec leurs valeurs manquantes
print("Nombre de valeurs manquantes par colonne :")
print(missing_values)

# 2 ANALYSE / TRAITEMENT ET NETTOYAGE DES DONNEES #

# Supprimer les lignes avec des valeurs manquantes dans la colonne 'Description'
df.dropna(subset=['Description'], inplace=True)

# Remplac les valeurs manquantes dans la colonne 'CustomerID' par 'Inconnu'
df['CustomerID'].fillna('Inconnu', inplace=True)

# Afficheons les premières lignes du dataFrame pour vérification
print(df.head())

# Compter les doublons dans la colonne 'InvoiceNo'
invoice_duplicates = df[df.duplicated(subset=['InvoiceNo'], keep=False)]

# Afficher les lignes doublons basées sur 'InvoiceNo'
print("Lignes doublons basées sur 'InvoiceNo' :")
print(invoice_duplicates)

# Afficher le nombre total de lignes doublons basées sur 'InvoiceNo'
print("\nNombre total de lignes doublons basées sur 'InvoiceNo' :", len(invoice_duplicates))

# 3 VISUALISATION #
# Statistiques descriptives pour les variables numériques
print(df[['Quantity', 'UnitPrice']].describe())

# Histogramme de la variable 'Quantity'
plt.figure(figsize=(10, 6))
sns.histplot(df['Quantity'], bins=50, kde=True)
plt.title('Distribution de la Quantité')
plt.xlabel('Quantité')
plt.ylabel('Fréquence')
plt.show()

# Boxplot de la variable 'UnitPrice'
plt.figure(figsize=(8, 6))
sns.boxplot(x=df['UnitPrice'])
plt.title('Boxplot du Prix Unitaire')
plt.xlabel('Prix Unitaire')
plt.show()

# Diagramme à barres des ventes par pays
plt.figure(figsize=(12, 6))
sns.countplot(x='Country', data=df, order=df['Country'].value_counts().index)
plt.title('Nombre de Transactions par Pays')
plt.xlabel('Pays')
plt.ylabel('Nombre de Transactions')
plt.xticks(rotation=90)
plt.show()

# Matrice de corrélation
corr_matrix = df[['Quantity', 'UnitPrice']].corr()
print("Matrice de corrélation :")
print(corr_matrix)

# Heatmap de la matrice de corrélation
plt.figure(figsize=(6, 4))
sns.heatmap(corr_matrix, annot=True, cmap='coolwarm', vmin=-1, vmax=1)
plt.title('Matrice de Corrélation')
plt.show()

# Détection  dans 'Quantity' avec la méthode IQR
Q1 = df['Quantity'].quantile(0.25)
Q3 = df['Quantity'].quantile(0.75)
IQR = Q3 - Q1
outliers_quantity = df[(df['Quantity'] < Q1 - 1.5 * IQR) | (df['Quantity'] > Q3 + 1.5 * IQR)]
print("Nombre d'outliers dans 'Quantity' :", len(outliers_quantity))



# clustering K-means sur 'Quantity' et 'UnitPrice'
X = df[['Quantity', 'UnitPrice']]


# Visualisation des clusters
plt.figure(figsize=(8, 6))
sns.scatterplot(x='Quantity', y='UnitPrice', hue='Cluster', data=df, palette='Set1', legend='full')
plt.title('Clustering des Clients')
plt.xlabel('Quantité')
plt.ylabel('Prix Unitaire')
plt.show()

