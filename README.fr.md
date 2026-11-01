[ 🌐 عربي ](README.ar.md) | [ 🇩🇪 Deutsch ](README.de.md) | [ 🇫🇷 Français ](README.fr.md) | [ 🇳🇱 Nederlands ](README.nl.md) | [ 🇪🇸 Español ](README.sp.md) | [ 🇬🇧 English ](README.md)

# Calculateur de prix de carte de restaurant, configuration des modificateurs & kit de préparation POS (modèle Excel & Web)
<img width="1920" height="1032" alt="image" src="https://github.com/user-attachments/assets/45a7c5cb-ec7e-4744-a38b-e705561fe0d6" />


[![License: Apache 2.0](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](LICENSE)
![Platform](https://img.shields.io/badge/Platform-Web%20App%20%2B%20Excel%20Template-2251FF)
![Tool Type](https://img.shields.io/badge/Tool-Restaurant%20Menu%20Management%20%26%20POS%20Pricing-051C2C)
![Architecture](https://img.shields.io/badge/Architecture-Formula--Driven%20No--VBA-success)

**Un calculateur de prix de carte de restaurant en ligne gratuit et un modèle Excel hors ligne de configuration POS conçus pour les restaurants indépendants, les pizzerias et les exploitants QSR. Centralisez les données maîtres du menu, calculez les suppléments de modificateurs complexes, configurez les règles de garnitures demi-pizza/entière, prévenez la facturation double des ingrédients de recette inclus et exportez des feuilles de menu validées prêtes pour le POS sans macros VBA.**

> **Testez immédiatement vos règles de prix de carte dans votre navigateur** — Aucune inscription de compte, installation de logiciel ni formule de tableur avancée requise.
>
> - 🌐 **[Lancez l'application web gratuite et interactive de calcul de prix de carte de restaurant](https://hyvoid.github.io/restaurant-menu-pricing-template/)** *(Testez en ligne des configurations de carte en session unique, la logique des modificateurs et les prix des commandes clients)*
> - 📥 **[Téléchargez le modèle Excel hors ligne de configuration de menu POS pour restaurant (.xlsx)](https://theseusworkshop.com/l/dznbyd?utm_source=github&utm_medium=GitHub%20README&utm_campaign=readme%20new%20launch&utm_content=restaurant-menu-pricing-template)** *(Conçu pour des enregistrements permanents de données maîtres du menu, des mises à jour de prix multi-saisons et des pistes d'audit POS opérationnelles — accompagné d'une garantie de remboursement de 30 jours sans questions)*

---

## Ce que ce modèle de gestion de carte de restaurant vous aide à suivre (des points sensibles aux solutions)

| Métrique de données & de prix de carte clé | Point sensible opérationnel courant des restaurants | Comment ce kit de configuration de carte le résout |
|---|---|---|
| **Hiérarchie des données maîtres du menu & disponibilité des tailles** | Les catégories de carte, les tailles d'articles (Petite/Moyenne/Grande) et la disponibilité saisonnière sont dispersées dans des PDF obsolètes et des exports POS. | Suit les catégories actives, sous-catégories, SKU d'articles et disponibilités par taille dans une seule base de données relationnelle de données maîtres du menu. |
| **Prix de base de la carte & suppléments de modificateurs à paliers** | Le personnel devine les suppléments de garnitures ou facture des forfaits fixes qui ignorent les différences de tailles, érodant les marges alimentaires. | Calcule les prix de base exacts plus les surcharges de modificateurs standard et premium proportionnées à la taille du produit. |
| **Ingrédients de recette par défaut (pas de facturation double)** | Les clients commandant des articles spéciaux (ex. pizza Hawaiian ou Supreme) se voient facturés en double lorsqu'ils conservent les garnitures standard incluses. | Mappe les ingrédients inclus par défaut pour chaque article afin que les quantités standard ne coûtent pas un centime de plus ($0) tandis que les portions « Extra » déclenchent automatiquement un supplément. |
| **Tarification fractionnaire & demi-pizza des garnitures** | Le calcul manuel des garnitures moitié gauche, moitié droite et pizza entière entraîne des retards au passage en caisse et des factures clients incohérentes. | Applique des multiplicateurs de prix fractionnaires globaux pour calculer automatiquement les suppléments de garnitures exacts pour demi-portion et portion entière. |
| **Limites de sélection min/max des groupes de modificateurs** | Les clients sélectionnent des options de combo invalides ou omettent des choix obligatoires (comme le type de pâte ou le degré de cuisson) à cause de règles POS défectueuses. | Applique des limites strictes de minimum, de maximum et de quantité par modificateur au niveau du groupe de modificateurs avant la mise en production. |
| **Préparation de la synchronisation de carte & statut QA avant POS** | Charger des enregistrements d'articles incomplets ou des prix non vérifiés directement dans un POS en production perturbe les tickets de cuisine et les commandes en ligne. | Signale les prix manquants et les liens de modificateurs rompus sur un tableau de bord de validation, isolant les lignes d'export propres et prêtes pour le POS. |

---

## Tutoriel pas à pas : comment configurer les prix de carte de restaurant & préparer les imports POS

Suivez ce flux de travail d'ingénierie de carte en 4 étapes pour standardiser votre base d'articles, tester des combinaisons de commandes complexes et générer une feuille d'import POS propre :

1. **Étape 1 : Configurez les paramètres globaux de taxe, de devise & de tarification fractionnaire du restaurant**
   **Action :** Ouvrez l'onglet `Parameters` et saisissez votre taux de taxe locale, votre symbole monétaire, le multiplicateur de prix des garnitures en demi-portion et le facteur de supplément des modificateurs premium.
   *Pourquoi c'est important :* Définir une seule fois les variables de prix globales établit une base de calcul unifiée pour toutes les catégories de carte et toutes les tailles d'articles.

2. **Étape 2 : Importez ou saisissez votre liste de données maîtres du menu, les tailles & les groupes de modificateurs**
   **Action :** Remplissez les tableaux d'entrée structurés avec vos catégories de carte, noms de plats, variantes de tailles, groupes de modificateurs et prix bruts d'ingrédients — ou collez des lignes existantes depuis votre export CSV POS actuel.
   *Pourquoi c'est important :* Relier les articles et les modificateurs via des identifiants relationnels propres élimine la double saisie de données lorsque les prix des ingrédients changent.

3. **Étape 3 : Mappez les inclusions de recette par défaut & simulez de vraies commandes clients**
   **Action :** Attribuez des groupes de modificateurs à des articles spécifiques du menu, marquez les ingrédients spéciaux intégrés et exécutez des commandes de test dans le **Simulateur de prix de commande** interactif (en testant les changements de taille, les garnitures moitié gauche/droite et les portions supplémentaires).
   *Pourquoi c'est important :* La simulation en temps réel des commandes avant et après taxe révèle les options sous-tarifées et les bugs de facturation double avant qu'un seul client ne passe commande.

4. **Étape 4 : Auditez la santé de la configuration & pérennisez la maintenance du menu avec Excel**
   **Action :** Consultez le tableau de bord QA automatisé pour lever les signaux d'articles incomplets, exportez votre table de préparation POS validée et **passez de la démo navigateur en session unique au classeur Excel réutilisable**.
   *Pourquoi c'est important :* Bien que le [calculateur de carte gratuit dans le navigateur](https://hyvoid.github.io/restaurant-menu-pricing-template/) soit idéal pour des tests ponctuels rapides, les sessions navigateur ne conservent pas durablement l'historique des modifications. **[Téléchargez le modèle Excel complet hors ligne de données maîtres du menu (.xlsx)](https://theseusworkshop.com/l/dznbyd?utm_source=github&utm_medium=GitHub%20README&utm_campaign=readme%20new%20launch&utm_content=restaurant-menu-pricing-template)** pour enregistrer la base de données permanente du menu de votre restaurant, effectuer des mises à jour trimestrielles récurrentes des prix, tenir des journaux d'audit historiques et réutiliser votre configuration sur plusieurs établissements sans repartir de zéro.

**Configurez une seule fois les règles de prix globales. Maintenez des données maîtres de menu propres et uniques. Simulez de vrais montants de passage en caisse client. N'exportez vers votre POS qu'après une validation automatisée.**

---

## Pourquoi les tableurs de carte de restaurant standard échouent sur les prix des modificateurs & des garnitures

La gestion de la carte d'un restaurant devient trompeusement complexe dès qu'un plat n'est plus défini par un simple prix fixe.

Prenons une pizza spéciale ou un repas combo personnalisable avec tarification par taille, garnitures par défaut incluses, options premium facultatives et quantités choisies par le client. Une liste de menu Excel ou Google Sheets classique peut enregistrer les prix de base des articles, mais elle échoue à séparer **les ingrédients déjà inclus dans la recette de base** des **modificateurs supplémentaires nécessitant un supplément**.

Cette zone d'ombre structurelle crée deux risques commerciaux coûteux pour les entreprises de restauration :

1. **Le piège de la facturation double du client :** Des clients se voient facturer par erreur en supplément des ingrédients standards déjà intégrés au prix de base d'un plat spécial, provoquant des litiges au passage en caisse et des abandons de commandes en ligne.
2. **L'érosion silencieuse des marges :** Caissiers et administrateurs POS appliquent des prix incohérents aux garnitures supplémentaires, aux passages à une taille supérieure ou aux choix demi-demi de pizza, car les règles de prix sont calculées à la main.

Par exemple, une **Pizza Hawaiian** inclut des portions normales de jambon et d'ananas dans son prix de base. Conserver les quantités standard de jambon et d'ananas doit ajouter **$0.00** à l'addition. Demander de l'ananas *en supplément* doit cependant déclencher une surcharge ajustée selon la taille — et poser une garniture premium sur la **moitié gauche** seulement doit suivre une règle de tarification fractionnaire définie plutôt qu'une estimation approximative.

Ce ne sont pas de simples erreurs de frappe. Elles surviennent parce que des tableurs à plat ne peuvent pas modéliser les règles commerciales derrière une carte de restaurant moderne.

Ce kit transforme ces relations de menu en un système de configuration structuré et piloté par formules. Au lieu de maintenir des listes de garnitures dupliquées pour chaque article et de vérifier manuellement les prix, vous définissez les ingrédients, les groupes de modificateurs et les règles de prix une seule fois — et vous les appliquez de manière cohérente sur toute votre carte.

---

## Tableurs de carte manuels vs kit structuré de configuration de menu POS

| Défi de configuration de carte de restaurant | Tableurs manuels & modifications POS directes (sans cet outil) | Données maîtres du menu structurées & kit de tarification (avec cet outil) |
|---|---|---|
| **Fragmentation des données maîtres du menu** | Catégories, tailles d'articles et prix de modificateurs vivent dans des onglets déconnectés ou des exports CSV désorganisés, provoquant un chaos de gestion de versions lors des mises à jour de carte. | Des tables maîtres relationnelles séparent catégories, articles de base et matrices de prix par taille tout en préservant les recherches automatiques d'identifiants. |
| **Maintenance dupliquée des modificateurs & garnitures** | Le même fromage, la même sauce ou la même garniture est saisi de façon répétée sur des dizaines d'articles, créant des noms divergents et des prix de suppléments obsolètes. | Chaque modificateur d'ingrédient est créé une fois dans une bibliothèque centrale et relié à plusieurs plats via des groupes de modificateurs réutilisables. |
| **Bugs de tarification des articles spéciaux & recettes par défaut** | Les garnitures standard incluses sur des pizzas ou burgers spéciaux sont facturées par erreur comme options payantes lorsque les clients personnalisent leur commande. | Le mappage des recettes par défaut distingue automatiquement les quantités standard intégrées (facturation $0) des upgrades payants « Extra ». |
| **Tarification incohérente des demi-pizzas & portions fractionnaires** | Le personnel calcule manuellement les suppléments de garnitures moitié gauche, moitié droite et pizza entière, provoquant des erreurs de facturation et des approximations de caissiers. | Des facteurs de tarification fractionnaire globaux calculent automatiquement les prix des garnitures demi-portion à partir du tarif de référence de la pizza entière. |
| **Règles de sélection de modificateurs POS rompues** | Les choix minimum obligatoires et les plafonds maximum de garnitures sont oubliés lors de la configuration, laissant des commandes impossibles atteindre l'écran de cuisine. | Les règles de contraintes au niveau du groupe de modificateurs appliquent des limites explicites de sélection min/max et des plafonds de quantité par article pendant la préparation POS. |
| **Erreurs de déploiement & de synchronisation POS en production** | Des prix non vérifiés et des lignes d'articles incomplètes sont saisis directement dans les terminaux POS en production, provoquant des plantages de caisse aux heures de pointe. | Une validation QA intégrée signale les données manquantes, simule les tickets de caisse clients et génère une table d'export POS propre et pré-vérifiée. |

---

## À qui s'adresse ce modèle de carte & de tarification de restaurant (rôles cibles & cas d'usage)

Conçu pour les restaurants indépendants, les pizzerias, les restaurants rapides (QSR), les concepts fast-casual et les exploitations multi-unités de restauration qui gèrent des cartes multi-tailles, des modificateurs d'ingrédients partagés, des suppléments conditionnels et des règles strictes de commande client.

### Scénarios de recherche par rôle & cas d'usage opérationnels

- **Propriétaires de pizzerias & exploitants de restaurants indépendants (`Pizza Menu Pricing Calculator & Topping Spreadsheet`)**
  Cessez de perdre de l'argent sur des garnitures supplémentaires non facturées ou de surfacturer des clients sur des pizzas spéciales. Modélisez les multiplicateurs de garnitures pizza entière vs demi-pizza, les surcharges par taille et les ingrédients de recette inclus avant de mettre à jour vos ardoises physiques ou votre carte de commande en ligne.
- **Responsables d'exploitation QSR & fast-casual (`Multi-Size Restaurant Menu Management Template`)**
  Standardisez les hiérarchies de carte, les groupes de modificateurs de combos et les ajustements de prix saisonniers sur plusieurs catégories de produits dans un seul fichier maître hors ligne centralisé, sans dépendre de scripts VBA fragiles.
- **Spécialistes de l'onboarding POS & administrateurs de carte (`Restaurant POS Menu Staging & Import Checklist`)**
  Nettoyez, organisez et validez les données brutes de carte avant de construire ou de mettre à jour des bases de données dans Toast, Square, Clover, Lightspeed ou des plateformes POS personnalisées. Détectez les prix de base manquants et les règles min/max de modificateurs rompues avant la mise en production.
- **Consultants en restauration & ingénieurs de carte (`Menu Engineering & Modifier Architecture Toolkit`)**
  Livrez aux clients de l'hôtellerie-restauration un classeur de configuration de carte répétable et prêt pour l'audit. Restructurez des cartes clients désordonnées en groupes de modificateurs logiques et démontrez le prix exact au passage en caisse lors de vos missions de conseil.

**Pour qui ce kit n'est PAS conçu :**
Les entreprises de restauration cherchant un terminal de point de vente en production, un processeur de paiements par carte, un système d'affichage cuisine (KDS), un suivi de déstockage de matières premières ou un calculateur automatique de nutrition des recettes. Ce kit est un **système pré-POS de données maîtres du menu, de simulation de tarification et de QA de configuration** — il prépare et valide la logique commerciale de votre carte afin que votre plateforme POS fonctionne sans erreurs de prix.

---

## La logique métier & la méthodologie : résoudre les angles morts cachés de la tarification en restaurant

*(Rédigé pour les propriétaires et exploitants de restaurants — aucune compétence technique ni avancée en tableur requise)*

Lorsque les clients commandent, la tarification du menu semble simple en surface : choisir un article, ajouter une garniture, payer le total. En coulisses, cependant, la plupart des restaurants indépendants perdent de l'argent ou frustrent leurs clients à cause de **trois angles morts commerciaux cachés** que les tableurs standards ne détectent pas.

Voici, en clair, comment la logique métier de ce kit protège vos marges et l'expérience client :

### 1. Éliminer le problème « une mise à jour, dix erreurs » (source unique de vérité)
* **L'angle mort :** Supposons que le coût de gros de la mozzarella ou du bacon bondisse de 20 %. Dans un tableur de restaurant typique, « Bacon supplément » figure séparément sous Burgers, Sandwichs, Salades et Pizzas. Si vous mettez le prix à jour à trois endroits mais oubliez le quatrième, vous perdez de la marge ou vous perturbez les clients avec des prix incohérents.
* **Comment le kit le corrige :** Vous saisissez chaque ingrédient et modificateur **une seule fois** dans une liste centrale. Que cet ingrédient soit utilisé sur 2 articles ou 50, la mise à jour de son prix dans une seule cellule met automatiquement à jour chaque article de menu connecté et chaque calcul de commande dans tout le classeur.

### 2. Prévenir le piège de la « facturation double des articles spéciaux » (tarification conditionnelle de recette)
* **L'angle mort :** Quand un client commande une « Meat Lover's Pizza » ou un « Deluxe Burger », il paie déjà un prix de base plus élevé parce que le pepperoni, la saucisse ou le fromage est inclus dans la recette. Si votre configuration de menu traite toutes les garnitures pareil, votre système peut facturer par erreur le client *une seconde fois* juste pour confirmer les ingrédients standards — ou, si vous rendez toutes les garnitures gratuites pour éviter les plaintes, vous perdez de l'argent dès que quelqu'un demande de la viande *double*.
* **Comment le kit le corrige :** Le modèle sépare **« Inclus par défaut »** de **« Payant en supplément »**. Il sait que 1x Pepperoni sur une Pepperoni Pizza ne coûte pas $0.00 de plus, tandis que porter la quantité à 2x (Pepperoni supplément) facture automatiquement le prix exact de l'upgrade. Vous préservez la confiance des clients sans distribuer de la nourriture gratuite.

### 3. Mettre fin au calcul à l'aveugle des caissiers sur les demi-garnitures & les tailles (logique de tarification proportionnelle)
* **L'angle mort :** Les clients demandent fréquemment des garnitures sur la seule moitié gauche ou droite d'une pizza, ou passent des tailles Petite, Moyenne et Grande. Sans formule claire, le personnel facture soit le prix plein de la garniture pour pizza entière (ce qui irrite le client), soit divise les prix en deux à la main sans tenir compte du travail de préparation (ce qui rogne vos marges).
* **Comment le kit le corrige :** Vous définissez une règle unique — par exemple, « une garniture demi-pizza coûte 60 % d'une garniture pizza entière ». Le kit calcule automatiquement le montant exact, au centime près, pour les placements moitié gauche, moitié droite et entière, sur chaque taille de pizza.

### 4. Bloquer les tickets de cuisine impossibles (garde-fous de sélection min/max)
* **L'angle mort :** Si un client commande un combo en ligne et oublie de choisir un accompagnement obligatoire — ou sélectionne 12 sauces gratuites alors que 2 seulement sont autorisées — votre personnel de cuisine doit arrêter la cuisson, appeler le client ou absorber le coût.
* **Comment le kit le corrige :** Vous définissez des règles de bornes simples pour chaque groupe d'options (ex. *« Il faut choisir au moins 1 pâte ; maximum 1 »* ou *« Jusqu'à 4 garnitures standards »*). Le simulateur de commande intégré teste immédiatement votre carte contre ces règles et vous montre si un client pourrait passer par erreur une commande invalide.

### 5. Tester avant la mise en production (assurance qualité pré-POS)
* **L'angle mort :** La plupart des restaurants découvrent les erreurs de tarification *après* avoir mis à jour leur système POS — généralement lorsqu'un client à l'heure de pointe se plaint à la caisse ou qu'une application de livraison en ligne facture un mauvais total.
* **Comment le kit le corrige :** Le kit joue le rôle d'un mannequin de crash-test sans risque pour votre carte. Vous pouvez simuler de vraies commandes clients, voir le ticket exact avant et après taxe, et consulter un tableau de santé automatique qui surligne en rouge les prix manquants. Vous n'exportez votre menu vers votre POS qu'une fois chaque article inspecté et validé.

---

## À propos du créateur

Je construis des outils métier pratiques et pilotés par formules qui transforment des données opérationnelles désordonnées en flux d'aide à la décision structurés et répétables.

La plupart des restaurants indépendants et petites entreprises disposent déjà de leurs données de prix, mais elles sont enfermées dans des tableurs déconnectés, des exports POS brouillons et des habitudes de personnel non documentées. Le vrai défi opérationnel consiste à relier ces enregistrements, à faire respecter des règles de prix cohérentes et à détecter les erreurs de configuration avant qu'elles n'atteignent les clients payants.

Le **Restaurant Menu Pricing Calculator & POS Configuration Toolkit** offre aux exploitants de restauration une validation de menu de niveau entreprise au sein d'un flux de travail navigateur et Excel transparent et sans superflu.

- 🌐 **[Essayez dès maintenant le calculateur de prix de carte gratuit dans le navigateur](https://hyvoid.github.io/restaurant-menu-pricing-template/)**
- 📥 **[Obtenez le classeur Excel complet et réutilisable de données maîtres du menu (.xlsx)](https://theseusworkshop.com/l/dznbyd?utm_source=github&utm_medium=GitHub%20README&utm_campaign=readme%20new%20launch&utm_content=restaurant-menu-pricing-template)**


## Détails techniques

<details>
<summary>Pour les réviseurs techniques, praticiens Excel et développeurs — dépliez pour examiner l'architecture du classeur, la logique de tarification, les formules et les règles de validation.</summary>

### Architecture du classeur

Le kit est conçu comme une base de données relationnelle de configuration de menu et un moteur de validation de prix, plutôt que comme une collection de feuilles de menu indépendantes.

Son architecture repose sur trois principes :

1. **Source unique de vérité :** chaque catégorie, produit, taille, modificateur et règle de prix possède un emplacement désigné dans les données maîtres.
2. **Relations réutilisables :** les articles de menu héritent des groupes de modificateurs via des affectations au niveau de la catégorie ou de l'article, au lieu de dupliquer des enregistrements de configuration.
3. **Validation et sortie automatisées :** des formules dynamiques relient les données maîtres, calculent les prix clients, détectent les erreurs de configuration et génèrent des enregistrements d'export POS standardisés.

Le classeur contient cinq modules fonctionnels et 13 feuilles de calcul.

#### 1. Paramètres globaux & règles de tarification

| Feuille de calcul | Responsabilité |
|---|---|
| `00_Param_Config` | Centralise la devise, le taux de taxe, les facteurs de prix entière/demi, les surcharges de modificateurs premium, les seuils opérationnels, les indicateurs de statut et les règles de tarification conditionnelle. |

Cette feuille est la couche de contrôle globale. Les hypothèses de tarification et les identifiants de statut métier sont gérés de manière centrale au lieu d'être figés comme valeurs fixes dans des formules de calcul individuelles.

#### 2. Données maîtres des produits du menu

| Feuille de calcul | Responsabilité |
|---|---|
| `01_Menu_Categories` | Gère les catégories de menu, sous-catégories, ordres d'affichage et disponibilités de catégorie. |
| `02_Menu_Items` | Stocke les identités de produits individuels, descriptions, relations de catégorie, types de produits et disponibilités. |
| `03_Item_Sizes` | Stocke les prix de produits spécifiques par taille et relie chaque taille à son article de menu parent. |

La séparation entre l'identité du produit et la tarification par taille empêche les données maîtres produits de s'élargir horizontalement à chaque introduction d'une nouvelle taille ou option de prix.

#### 3. Modificateurs réutilisables & tarification multidimensionnelle

| Feuille de calcul | Responsabilité |
|---|---|
| `04_Modifier_Groups` | Définit les groupes de modificateurs visibles par les clients et leur périmètre opérationnel. |
| `05_Modifiers` | Gère les ingrédients et identités de modificateurs réutilisables à l'échelle globale. |
| `06_Modifier_Pricing` | Stocke les prix des modificateurs par taille de produit et niveau de portion, avec les calculs de placement entière/gauche/droite. |

Le même ingrédient peut être réutilisé sur plusieurs produits sans créer de copies séparées de son enregistrement maître.

#### 4. Cartographie des relations & contraintes de sélection

| Feuille de calcul | Responsabilité |
|---|---|
| `07_Default_Modifiers` | Mappe les produits spéciaux vers leurs ingrédients inclus et quantités par défaut. |
| `08_Group_Assignment` | Attribue les groupes de modificateurs au niveau de la catégorie ou du produit individuel. |
| `09_Selection_Rules` | Contrôle les sélections minimales, les sélections maximales, les groupes obligatoires et les autorisations de sélection multiple. |

Ce module détermine quels groupes de modificateurs s'appliquent à chaque produit et si la configuration de sélection client qui en résulte est logiquement valide.

#### 5. Simulation de commande, export POS & surveillance

| Feuille de calcul | Responsabilité |
|---|---|
| `10_Order_Test` | Simule les sélections clients et calcule les prix finaux tout en vérifiant la conformité aux règles de sélection. |
| `11_POS_Export` | Génère une liste standardisée des produits disponibles, tailles, prix, résumés de modificateurs et statuts de synchronisation. |
| `12_Dashboard` | Résume la santé du menu, les anomalies de configuration, la disponibilité des produits et la préparation POS. |

#### Flux de données

Le flux de données complet suit cette séquence :

```text
GLOBAL PARAMETERS
00_Param_Config
       |
       v
MENU MASTER DATA                 MODIFIER MASTER DATA
01_Menu_Categories              04_Modifier_Groups
       |                        05_Modifiers
02_Menu_Items                   06_Modifier_Pricing
       |                               |
03_Item_Sizes                           |
       |                                |
       +---------------+----------------+
                       |
                       v
             RELATIONSHIP & RULES
             07_Default_Modifiers
             08_Group_Assignment
             09_Selection_Rules
                       |
          +------------+-------------+
          |            |             |
          v            v             v
       10_Order     11_POS        12_Dashboard
         Test       Export
          |            |             |
          v            v             v
       Price QA    POS Update    Configuration
                                  Monitoring
````

### Cadre de décision

Le kit est construit autour de quatre décisions opérationnelles.

| Décision | Informations requises | Résultat de la décision |
| ------------------------------------------ | ---------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------- |
| Que peut-on vendre ? | Enregistrements de catégorie, de produit, de taille et de disponibilité. | Identifie la structure de menu valide et les configurations de produits disponibles. |
| Que doivent payer les clients ? | Prix de base, prix des modificateurs, ingrédients inclus, quantités de portion, placement et taxe. | Détermine le prix client calculé pour une configuration de commande spécifique. |
| Que peuvent sélectionner les clients ? | Affectations de modificateurs, appartenance aux groupes, statut obligatoire et limites de sélection min/max. | Identifie les configurations de commande valides et invalides. |
| Le menu est-il prêt pour la synchronisation POS ? | Intégrité des données maîtres, complétude de la tarification, relations de modificateurs et statut de validation. | Sépare les enregistrements prêts pour l'export de ceux nécessitant une correction. |

Le cadre distingue trois types d'informations :

- Faits produits : ce que le restaurant vend et quelles tailles sont disponibles.
- Règles commerciales : comment les ingrédients, les quantités et les placements affectent le prix de vente.
- Contraintes opérationnelles : quelles sélections sont permises et si la configuration est prête pour le déploiement.

Cette séparation empêche qu'une description de produit, une hypothèse de tarification et une restriction de commande soient traitées comme des données interchangeables.

### Trois pièges courants de tarification & de configuration

#### Piège 1 — Facturer deux fois des ingrédients déjà inclus

La décision : fixer le prix client d'une pizza spéciale d'après ses ingrédients listés.

Le défaut non remarqué : un calcul de modificateurs conventionnel ajoute le prix standard de la garniture, que celle-ci soit ou non déjà incluse dans le produit spécial.

Prenons une Pizza Hawaiian avec jambon et ananas inclus dans son prix de base.

| Sélection client | Traitement incorrect | Traitement correct |
| ------------------ | ----------------------------------------------------------- | -------------------------------------------------- |
| Jambon normal | Ajoute un supplément de garniture | Aucun supplément |
| Ananas en portion légère (Light) | Ajoute un supplément de garniture | Aucun supplément |
| Jambon supplémentaire (Extra) | Ajoute un supplément de garniture sans tenir compte de l'inclusion | Applique la règle de tarification configurée pour quantité supplémentaire |
| Ananas supplémentaire (Extra) | Utilise le prix standard du modificateur sans vérifier l'inclusion | Applique la règle de tarification configurée pour quantité supplémentaire |

L'approche incorrecte traite chaque ingrédient sélectionné comme un nouvel achat. L'approche corrigée reconnaît qu'un ingrédient peut déjà faire partie de la définition commerciale du produit.

Le mappage `07_Default_Modifiers` identifie les ingrédients inclus. Les règles de tarification conditionnelle déterminent ensuite le multiplicateur applicable selon le statut d'inclusion et le type de portion sélectionné.

#### Piège 2 — Traiter les garnitures demi-pizza comme des prix fixes indépendants

La décision : facturer le même prix de modificateur, que la garniture couvre toute la pizza ou une seule moitié.

Le défaut non remarqué : la quantité de produit recevant la garniture change, mais le calcul de prix ne reflète pas le placement sélectionné.

Par exemple, supposons qu'une garniture pizza entière coûte $2.00.

| Placement | Prix incorrect | Prix correct |
| --------- | --------------- | ------------- |
| Entière | $2.00 | $2.00 |
| Gauche | $2.00 | $1.00 |
| Droite | $2.00 | $1.00 |

L'approche incorrecte surfacture les choix demi-pizza lorsque la politique du restaurant est une tarification proportionnelle.

L'approche corrigée déduit les prix des moitiés gauche et droite du prix de référence de la pizza entière à l'aide de facteurs de tarification centralisés.

#### Piège 3 — Traiter la préparation du menu POS comme une simple tâche de saisie

La décision : envoyer la dernière feuille de calcul du menu au fournisseur POS après avoir mis à jour les prix des produits.

Le défaut non remarqué : une saisie de prix complète n'établit pas que l'ensemble de la configuration du produit est valide.

Un produit peut avoir un prix de base mais contenir encore :

- Un enregistrement de taille manquant.
- Une relation de catégorie invalide.
- Un groupe de modificateurs indisponible.
- Des mappages d'ingrédients par défaut manquants.
- Des règles de sélection minimale et maximale incohérentes.

| Approche de revue | Risque | Approche corrigée |
| ----------------------------------------- | ------------------------------------------------ | ---------------------------------------------------------- |
| Vérifier si le nom du produit existe. | Des produits incomplets semblent prêts. | Valider l'identité du produit et les relations de catégorie. |
| Vérifier si un prix de base existe. | Certaines tailles disponibles restent sans prix. | Valider chaque combinaison produit-taille configurée. |
| Vérifier si les groupes de modificateurs sont listés. | Des groupes invalides ou vides peuvent atteindre le POS. | Valider l'appartenance aux groupes et les contraintes de sélection. |
| Exporter tout le menu. | Les erreurs sont transférées à la configuration POS. | Exporter les enregistrements éligibles et passer en revue le statut de synchronisation. |

L'approche corrigée distingue l'achèvement de la saisie de données de la disponibilité opérationnelle.

### Exemple de scénario

Scénario : tester une Pizza Hawaiian Large avec Ananas supplémentaire sur la moitié gauche

Un restaurant veut confirmer que la tarification de sa pizza spéciale reste correcte lorsque les clients changent la quantité et le placement des ingrédients inclus.

Supposons la configuration illustrative suivante :

| Entrée | Valeur |
| -------------------- | -------------- |
| Produit | Pizza Hawaiian |
| Taille | Large |
| Prix de base | $18.00 |
| Ingrédients inclus | Jambon, Ananas |
| Modificateur sélectionné | Ananas |
| Quantité sélectionnée | Extra |
| Placement | Gauche |
| Prix du modificateur entier | $2.00 |
| Facteur demi gauche | 0.50 |
| Taxe de vente | 8.00% |

Étape 1 — Identifier l'ingrédient inclus.

L'ananas est mappé comme ingrédient inclus pour la Pizza Hawaiian. Le client n'est donc pas facturé à nouveau pour la quantité standard incluse.

Étape 2 — Déterminer le multiplicateur de quantité applicable.

Le client sélectionne Extra plutôt que Normal. Selon la matrice de règles illustrative, un ingrédient inclus sélectionné en Extra reçoit un multiplicateur de 1.00.

Étape 3 — Calculer le prix du modificateur ajusté selon le placement.

```
Whole Modifier Price = $2.00

Left Half Price = $2.00 × 0.50

Raw Portion Price = $1.00

Applicable Multiplier = 1.00

Additional Charge = $1.00 × 1.00

Additional Charge = $1.00
```

Étape 4 — Calculer le prix client.

| Calcul | Montant |
| --------------------------- | ------ |
| Pizza Hawaiian Large | $18.00 |
| Ananas supplémentaire — moitié gauche | $1.00 |
| Total avant taxe | $19.00 |
| Taxe de vente (8%) | $1.52 |
| Prix client final | $20.52 |

Interprétation analytique

Un calculateur de garnitures conventionnel pourrait facturer à nouveau le prix standard de l'ananas, même lorsque l'ingrédient est déjà inclus. À l'inverse, un calculateur ignorant le placement pourrait facturer les $2.00 complets pour un choix demi-pizza.

Le résultat correct reflète les deux conditions commerciales : le statut d'ingrédient inclus et le placement sélectionné.

Implication de décision : le restaurant peut valider le prix vu par le client avant de modifier son menu en production, réduisant le risque de facturations incohérentes et de corrections POS évitables.

Ce scénario illustre la logique de tarification configurée. Les résultats réels dépendent de la matrice de prix propre au restaurant, des prix des modificateurs et des réglages de taxe.


### Référence des formules

Le classeur utilise des fonctions natives de tableaux dynamiques d'Excel pour maintenir les relations entre les enregistrements de menu, calculer les prix et générer les sorties de validation.

Les fonctions principales incluent :

| Fonction | Objectif |
|---|---|
| `LET` | Nomme les calculs intermédiaires et garde les formules complexes lisibles. |
| `FILTER` | Extrait les enregistrements valides et génère dynamiquement des listes de sortie. |
| `XLOOKUP` | Récupère les informations liées de produits, catégories, modificateurs et prix. |
| `COUNTIF` / `COUNTIFS` | Compte les enregistrements et vérifie les conditions de sélection ou de configuration. |
| `UNIQUE` | Génère des identifiants distincts de catégories et de produits. |
| `SUM` / `SUMIFS` | Agrège les suppléments de modificateurs et les valeurs métier. |
| `IF` / `AND` / `OR` | Applique les conditions de prix et les règles de validation. |
| `ROUND` | Contrôle la précision des calculs monétaires. |
| `AVERAGEIFS` | Calcule les prix de départ moyens au niveau des catégories. |

#### 1. Paramètres globaux — `00_Param_Config`

La feuille de paramètres gère deux types de contrôles :

**Paramètres de configuration statiques**

| Paramètre | Référence | Objectif |
|---|---|---|
| Symbole monétaire | `B3` | Contrôle la présentation monétaire. |
| Taux de taxe de vente | `B4` | Calcule les prix clients finaux. |
| Facteur demi gauche | `B5` | Calcule les prix des modificateurs moitié gauche. |
| Facteur demi droite | `B6` | Calcule les prix des modificateurs moitié droite. |
| Facteur entier | `B7` | Contrôle la tarification de la portion entière. |
| Surcharge de modificateur premium | `B8` | Fournit l'ajustement de prix premium configuré. |
| Seuil d'alerte des groupes de modificateurs | `B9` | Signale les affectations excessives de groupes de modificateurs. |
| Seuil d'alerte des prix élevés | `B10` | Identifie les prix de produits anormalement élevés. |
| Référentiel de disponibilité pour l'export POS | `B11` | Définit le statut utilisé pour le filtrage de l'export POS. |
| Statut actif par défaut | `B12` | Centralise l'identifiant de statut actif. |
| Statut inclus | `B13` | Identifie les ingrédients inclus par défaut. |
| Palier de modificateur standard | `B14` | Identifie le palier de modificateur standard. |

**Matrice de tarification conditionnelle**

La zone `E3:J10` maintient les règles de tarification conditionnelle.

Sa logique combine :

- Condition de statut inclus
- Condition de type de sélection
- Résultat du multiplicateur de prix

Une clé de recherche composite est générée pour relier les conditions au multiplicateur de prix.

Cela empêche les mêmes hypothèses de tarification d'être intégrées de manière indépendante dans plusieurs feuilles de calcul.

#### 2. Données maîtres du menu — `01_Menu_Categories`, `02_Menu_Items`, `03_Item_Sizes`

**Hiérarchie des catégories**

`01_Menu_Categories!H3` génère les chemins de catégories :

```excel
=LET(
    cat_name, FILTER(B3:B100, E3:E100<>""),
    sub_name, FILTER(F3:F100, E3:E100<>""),
    cat_name & " > " & sub_name
)
````

Les chemins obtenus fournissent des relations de catégories lisibles, telles que :

```
Pizza > Create Your Own
Pizza > Gourmet Pizza
Subs > Hot Subs
```

Nombre de produits actifs par sous-catégorie

`01_Menu_Categories!I3` :

```
=LET(
    sub_ids, FILTER(E3:E100, E3:E100<>""),
    COUNTIFS(
        '02_Menu_Items'!D3:D500, sub_ids,
        '02_Menu_Items'!G3:G500, '00_Param_Config'!$B$12
    )
)
```

Statut de configuration des catégories

`01_Menu_Categories!J3` :

```
=LET(
    cat_active, FILTER(D3:D100, E3:E100<>""),
    item_cnt, I3#,
    IF(cat_active <> '00_Param_Config'!$B$12, "⏸️ 已停用",
        IF(item_cnt = 0, "⚠️ 空分类预警", "✅ 正常")
    )
)
```

Clé composite produit-taille

`03_Item_Sizes` génère une clé composite à l'aide de :

```
=Item_ID & "|" & Size_ID
```

Par exemple :

```
ITM001|SZ_SM
ITM001|SZ_LG
```

Cette clé permet au classeur de distinguer les multiples tailles appartenant à un même produit.

Validation des prix produit-taille

La couche de validation des prix vérifie si un prix de taille configuré est manquant, invalide ou dépasse le seuil d'alerte centralisé.

Le statut obtenu est utilisé à la fois par les données maîtres produits et par le processus d'export POS.

#### 3. Données maîtres des modificateurs & tarification — `04_Modifier_Groups`, `05_Modifiers`, `06_Modifier_Pricing`

Nombre de modificateurs actifs

`04_Modifier_Groups!F3` :

```
=LET(
    grp_ids, FILTER(A3:A100, A3:A100<>""),
    COUNTIFS(
        '05_Modifiers'!B3:B500, grp_ids,
        '05_Modifiers'!E3:E500, '00_Param_Config'!$B$12
    )
)
```

Recherche des règles de sélection de groupe de modificateurs

`04_Modifier_Groups!G3` :

```
=LET(
    grp_ids, FILTER(A3:A100, A3:A100<>""),
    XLOOKUP(
        grp_ids,
        '09_Selection_Rules'!A3:A100,
        '09_Selection_Rules'!B3:B100,
        -1
    )
)
```

La recherche du maximum de sélections correspondant utilise la colonne C de `09_Selection_Rules`.

Une valeur de retour de `-1` identifie une règle manquante.

Dimensions de tarification des modificateurs

`06_Modifier_Pricing` maintient la relation entre :

```
Modifier ID
    +
Size ID
    +
Selection Type
```

La clé de tarification composite permet de récupérer le prix correct du modificateur pour une taille de produit et un niveau de quantité sélectionnés.

Les prix entier, gauche et droite sont dérivés des facteurs de tarification centralisés.

L'architecture de tarification sépare le prix de référence du calcul de placement, permettant d'ajuster des politiques de tarification spécifiques au restaurant sans réécrire le simulateur de commande.

#### 4. Ingrédients par défaut & cartographie des relations — `07_Default_Modifiers`, `08_Group_Assignment`, `09_Selection_Rules`

Clé de recherche des modificateurs par défaut

`07_Default_Modifiers!E3` :

```
=LET(
    itm_ids, FILTER(A3:A1000, A3:A1000<>""),
    mod_ids, FILTER(B3:B1000, A3:A1000<>""),
    itm_ids & "|" & mod_ids
)
```

Référence du nom de produit

`07_Default_Modifiers!F3` :

```
=LET(
    itm_ids, FILTER(A3:A1000, A3:A1000<>""),
    XLOOKUP(
        itm_ids,
        '02_Menu_Items'!A3:A500,
        '02_Menu_Items'!B3:B500,
        "❌ 商品ID无效"
    )
)
```

Référence du nom de modificateur

`07_Default_Modifiers!G3` :

```
=LET(
    mod_ids, FILTER(B3:B1000, A3:A1000<>""),
    XLOOKUP(
        mod_ids,
        '05_Modifiers'!A3:A500,
        '05_Modifiers'!C3:C500,
        "❌ 配料ID无效"
    )
)
```

Validation de la configuration par défaut

`07_Default_Modifiers!H3` vérifie :

- Si l'identifiant du produit existe.
- Si l'identifiant du modificateur existe.
- Si la même combinaison produit-modificateur est dupliquée.
- Si le statut inclus est correctement configuré.

Affectations au niveau catégorie et au niveau article

`08_Group_Assignment` utilise le champ `Level` pour distinguer le périmètre d'affectation :

```
Level = Category
    Category ID = valid category
    Item ID = -

Level = Item
    Item ID = valid product
    Category ID = -
```

La couche de validation des affectations vérifie le périmètre effectif et identifie les affectations invalides ou excessives.

Règles de sélection

`09_Selection_Rules` utilise des règles au niveau du groupe pour déterminer :

- Les sélections minimales.
- Les sélections maximales.
- Le statut obligatoire.
- Si les sélections multiples sont permises.

Ces valeurs sont référencées par le simulateur de commande et les formules de validation des groupes de modificateurs.

#### 5. Simulation du prix client — `10_Order_Test`

Cette feuille de calcul combine le prix de base du produit avec les suppléments de modificateurs réels générés par la configuration de commande sélectionnée.

Nom du produit sélectionné

`E3` :

```
=XLOOKUP(
    B3,
    '02_Menu_Items'!A3:A500,
    '02_Menu_Items'!B3:B500 & " (" &
    '02_Menu_Items'!F3:F500 & ")",
    "⚠️ 请选择有效商品ID"
)
```

Prix de base de la taille

`E4` :

```
=XLOOKUP(
    B3 & "|" & B4,
    '03_Item_Sizes'!F3#,
    '03_Item_Sizes'!D3:D1000,
    0
)
```

Supplément total des modificateurs

`E5` :

```
=SUM(J10#)
```

Prix client avant taxe

`E6` :

```
=E4+E5
```

Prix client final incluant la taxe

`E7` :

```
=ROUND(
    E6 * (1 + '00_Param_Config'!$B$4),
    2
)
```

Les calculs de détail des modificateurs suivent cette séquence :

```
Selected Modifier ID
        |
        v
Modifier Master Lookup
        |
        v
Default Ingredient Recognition
        |
        v
Conditional Pricing Rule Lookup
        |
        v
Size & Portion Price Retrieval
        |
        v
Placement Adjustment
        |
        v
Calculated Line Price
        |
        v
Group Selection Validation
```

Le prix final est calculé à partir de la sélection réellement simulée plutôt que d'un devis assemblé manuellement.

#### 6. Export POS — `11_POS_Export`

La feuille d'export POS filtre dynamiquement les produits selon leur disponibilité configurée.

Filtrage des produits disponibles

`A3` :

```
=LET(
    sz_itm_ids,
    FILTER(
        '03_Item_Sizes'!A3:A1000,
        '03_Item_Sizes'!A3:A1000<>""
    ),
    itm_avail,
    XLOOKUP(
        sz_itm_ids,
        '02_Menu_Items'!A3:A500,
        '02_Menu_Items'!G3:G500,
        "No"
    ),
    FILTER(
        sz_itm_ids,
        itm_avail = '00_Param_Config'!$B$11
    )
)
```

Nom du produit exporté

`B3` :

```
=XLOOKUP(
    A3#,
    '02_Menu_Items'!A3:A500,
    '02_Menu_Items'!B3:B500,
    ""
)
```

Nom de la taille exportée

`D3` :

```
=XLOOKUP(
    A3# & "|" & C3#,
    '03_Item_Sizes'!F3#,
    FILTER(
        '03_Item_Sizes'!C3:C1000,
        '03_Item_Sizes'!A3:A1000<>""
    ),
    ""
)
```

Prix de base exporté

`E3` :

```
=XLOOKUP(
    A3# & "|" & C3#,
    '03_Item_Sizes'!F3#,
    FILTER(
        '03_Item_Sizes'!D3:D1000,
        '03_Item_Sizes'!A3:A1000<>""
    ),
    0
)
```

Statut de synchronisation POS

La validation de l'export combine l'intégrité du produit et la validation des prix par taille.

Le résultat attendu est :

```
READY_FOR_POS
HOLD_CHECK_ERROR
```

L'export est conçu pour faciliter la revue et le transfert des modifications de configuration produit. Il ne s'agit pas d'une intégration directe à l'API POS ni d'un service de synchronisation automatique.

#### 7. Tableau de bord & surveillance de la configuration — `12_Dashboard`

Le tableau de bord combine six indicateurs opérationnels :

| KPI | Objectif du calcul |
| ------------------------- | --------------------------------------------- |
| Catégories actives totales | Compte les catégories de menu activées. |
| Articles disponibles totaux | Compte les produits disponibles. |
| Modificateurs réutilisables totaux | Compte les modificateurs activés. |
| Règles de tarification totales | Compte les règles de tarification configurées. |
| Anomalies de configuration totales | Agrège les exceptions de configuration détectées. |
| Articles prêts pour POS | Compte les enregistrements marqués prêts pour l'export. |

Articles disponibles totaux

`C4` :

```
=COUNTIF(
    '02_Menu_Items'!G3:G500,
    '00_Param_Config'!$B$12
)
```

Anomalies de configuration totales

`F4` :

```
=COUNTIF('02_Menu_Items'!M3#, "❌*")
+ COUNTIF('03_Item_Sizes'!I3#, "❌*")
+ COUNTIF('04_Modifier_Groups'!I3#, "❌*")
+ COUNTIF('08_Group_Assignment'!I3#, "⚠️*")
+ COUNTIF('09_Selection_Rules'!H3#, "❌*")
```

Nombre d'enregistrements prêts pour POS

`G4` :

```
=COUNTIF(
    '11_POS_Export'!I3#,
    "READY_FOR_POS"
)
```

Le tableau de bord offre une vue consolidée de la santé de la configuration du menu, tandis que les feuilles sous-jacentes conservent les enregistrements détaillés nécessaires pour examiner et corriger les exceptions individuelles.

### Règles de validation

Le cadre de validation est conçu pour identifier les erreurs avant qu'elles n'affectent les prix vus par les clients ou la configuration POS.

| Champ / Zone | Règle de validation | Comportement en cas d'erreur |
| --------------------------- | ---------------------------------------------------------------------------- | ------------------------------------------------------------------- |
| ID de catégorie | Doit être unique et référencer une catégorie valide. | Les relations de catégories invalides sont signalées. |
| ID de sous-catégorie | Doit correspondre à une sous-catégorie existante. | La classification de produits invalide est identifiée. |
| Catégorie active | Doit utiliser les valeurs de statut Oui/Non configurées. | Les valeurs de statut invalides affectent la validation des catégories. |
| ID d'article | Doit être un identifiant de produit unique. | Les identités de produits invalides ou dupliquées nécessitent une correction. |
| Catégorie de l'article | Doit référencer une catégorie et une sous-catégorie valides. | La validation d'intégrité des produits identifie les relations incohérentes. |
| Disponibilité de l'article | Doit utiliser l'identifiant de statut actif configuré. | Une disponibilité invalide peut empêcher un filtrage correct. |
| Taille de l'article | Doit référencer une combinaison produit-taille valide. | Les enregistrements produit-taille manquants ou invalides sont signalés. |
| Prix de base | Doit être présent et satisfaire les contrôles de tarification configurés. | Les prix manquants ou anormaux empêchent la préparation POS. |
| ID de groupe de modificateurs | Doit être un identifiant de groupe unique. | Les relations invalides sont signalées. |
| ID de modificateur | Doit être unique et référencer un groupe de modificateurs valide. | Les erreurs de recherche de modificateurs sont identifiées. |
| Modificateur actif | Doit utiliser l'identifiant de statut actif configuré. | Les modificateurs inactifs sont exclus des décomptes de modificateurs actifs. |
| Tarification des modificateurs | Doit contenir les enregistrements de prix requis par taille et par type de sélection. | Les matrices de prix incomplètes nécessitent une correction. |
| Palier de modificateur | Doit correspondre aux identifiants de palier de modificateur configurés. | Une classification de palier invalide affecte la validation des prix. |
| Mappage des modificateurs par défaut | Les identifiants de produit et de modificateur doivent exister. | Les identifiants invalides sont signalés. |
| Unicité des modificateurs par défaut | La même combinaison produit-modificateur ne doit pas être dupliquée. | Les mappages dupliqués sont signalés. |
| Statut inclus | Doit correspondre à l'identifiant de statut inclus configuré le cas échéant. | Une configuration d'inclusion par défaut incorrecte est signalée. |
| Niveau d'affectation | Doit être soit `Category`, soit `Item`. | Un périmètre d'affectation invalide est signalé. |
| Affectation au niveau catégorie | Requiert un ID de catégorie valide et le placeholder d'ID d'article prescrit. | Une configuration de périmètre invalide est signalée. |
| Affectation au niveau article | Requiert un ID d'article valide et le placeholder d'ID de catégorie prescrit. | Une configuration de périmètre invalide est signalée. |
| Affectation de groupe de modificateurs | Doit référencer un groupe de modificateurs existant. | Les relations de groupes invalides sont signalées. |
| Capacité du groupe de modificateurs | Le nombre maximum de sélections ne doit pas dépasser le nombre de modificateurs actifs disponibles. | La configuration de sélection nécessite une correction. |
| Groupe de modificateurs obligatoire | Les groupes obligatoires doivent avoir une sélection minimale supérieure à zéro. | Le conflit logique est signalé. |
| Groupe de modificateurs facultatif | Les groupes facultatifs doivent avoir une sélection minimale de zéro. | Les règles de sélection incohérentes sont signalées. |
| Sélection multiple | Doit respecter l'autorisation de sélection multiple configurée. | Les configurations de sélection invalides nécessitent une revue. |
| Produit de test client | Doit référencer un produit valide. | Une sélection de produit invalide affiche un avertissement. |
| Taille de test client | Doit référencer une taille valide pour le produit sélectionné. | Une recherche de prix de taille manquante renvoie un résultat invalide ou à prix zéro. |
| Sélection de modificateur client | Doit référencer un modificateur disponible. | Les références de modificateurs invalides nécessitent une correction. |
| Type de portion client | Doit utiliser `Light`, `Normal`, `Extra` ou `Xxtra`. | Les valeurs de sélection invalides ne sont pas des entrées de tarification valides. |
| Placement client | Doit utiliser `Whole`, `Left` ou `Right`. | Les valeurs de placement invalides ne sont pas des entrées de tarification valides. |
| Tarification conditionnelle | Doit résoudre vers la règle d'inclusion et de quantité configurée. | Une configuration de prix incorrecte ou manquante nécessite une investigation. |
| Prix client | Doit refléter le prix de base, les suppléments de modificateurs et le taux de taxe configuré. | Le simulateur de commande expose le calcul résultant pour revue. |
| Export POS | La validation du produit et du prix par taille doit réussir. | `HOLD_CHECK_ERROR` remplace `READY_FOR_POS`. |

Distinction opérationnelle importante : la couche de validation vérifie les relations et règles métier représentées dans le classeur. Elle ne vérifie pas de manière indépendante si un fournisseur POS accepte chaque champ exporté ni s'il prend en charge chaque comportement de tarification.

</details>


## Autres outils de cette série

D'autres outils légers d'aide à la décision et classeurs opérationnels :

\- [Theseus Workshop](https://theseusworkshop.com) — Une collection d'outils pratiques pour les problèmes opérationnels et financiers récurrents des petites entreprises.

## Licence

Ce projet est sous licence Apache License 2.0.

Consultez le fichier LICENSE pour les conditions de licence complètes.
