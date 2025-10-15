# Module Demande de Facturation Projet - Odoo 18

## Description
Ce module ajoute un workflow de validation pour les demandes de facturation sur les projets Odoo.

## Fonctionnalités
- ✅ Bouton "Demande de Facturation" sur le formulaire du projet
- ✅ Wizard interactif pour sélectionner les lignes et quantités à facturer
- ✅ Workflow de validation : Brouillon → Soumise → Approuvée → Facturée
- ✅ Gestion des facturations partielles multiples
- ✅ Vérification automatique des quantités disponibles
- ✅ Création automatique de la facture depuis la demande approuvée

## Installation
1. Copiez le dossier `project_invoice_request` dans votre répertoire addons
2. Mettez à jour la liste des modules : Apps > Update Apps List
3. Recherchez "Demande de Facturation Projet"
4. Cliquez sur "Installer"

## Utilisation
1. Ouvrez un projet ayant une commande client associée
2. Cliquez sur le bouton "Demande de Facturation"
3. Sélectionnez les lignes et quantités à facturer
4. Soumettez la demande
5. Un manager peut approuver/rejeter la demande
6. Une fois approuvée, créez la facture en un clic

## Dépendances
- project
- sale_project
- account

## Auteur
Votre Société

## License
LGPL-3
