# Outil de relance client

Mandataire : Arbre Conseil, structure rattachée à l'ONF"
Mandaté : IA School, école rattachée au groupe GEMA

Auteurs: Quoc Hiep DAO, Sven LOTHE

Ce programme a pour but la relance client par mail pour le suivi de santé de leurs arbres.
Pour cela le programme est conçu afin:

- d'extraire des informations depuis un fichier excel de référence
- d'effectuer un envoi de mail type via outlook figurant les informations collectées
- de permettre une utilisation simple au moyen d'une interface graphique

Le programme a été codé en Python 3.8
Consulter le fichier "Requirements.md" pour la liste des prérequis système

***

**ATTENTION :** 

La solution a été codée dans une logique temporaire.
Ce programme ne saurait constituer une solution durable et sera **__inefficace__** si:

- la structure du fichier excel de référence est modifiée
- la boite mail d'envoi n'utilise pas les services d'outlook
- les prérequis figurant dans "Requirements.md" ne sont pas satisfaits

Les colonnes de références du fichier excel sont:

- Nom du client
- Coordonnées mail
- Site
- Numéro de l'arbre
- Date du relevé
- Nom français de l'arbre
- Type de contrôle/suivi
- Date de relance pour contrôle/suivi
- Type d'intervention
- Date de relance pour intervention 1
- Type d'intervention 2
- Date de relance pour intervention 2
- x
- y

Modifier l'intitulé de ces colonnes ou les supprimer rendra le programme inefficace.
Le programme est en mesure d'analyser n'importe quel fichier .xlsx présentant ces colonnes.

***

**NOTA BENE :** 

Ce programme stocke des dates de relance dans un fichier texte par soucis d'ergonomie.
Aucune information autre que ces dates n'est stockée.
