# Option Pricing and Hedging – Excel & VBA

Ce projet implémente des modèles de valorisation et de couverture d’options en utilisant **Excel et VBA**, dans le cadre d’un projet de finance de marché centré sur les produits dérivés et la gestion des risques.

L’objectif est de construire un cadre quantitatif permettant de :

- Valoriser des options européennes via le modèle de Black-Scholes
- Calculer les sensibilités (Delta, Gamma, Theta, Vega)
- Implémenter des stratégies de couverture
- Analyser le P&L quotidien et son exposition aux risques

---

## Contexte du projet

Ce projet a été réalisé dans le cadre d’un cours de **Finance de Marché** portant sur la valorisation des dérivés et la gestion des risques.

L’objectif était d’appliquer le modèle théorique de **Black-Scholes-Merton** dans un environnement pratique sous Excel/VBA et d’évaluer des stratégies de couverture sur données historiques réelles.

Sous-jacent étudié : **Indice Euro Stoxx 50**

---

## Méthodologie

### Implémentation du modèle Black-Scholes (VBA)

Le modèle inclut :

- Pricing des Calls et Puts européens
- Calcul du Delta
- Calcul du Gamma
- Calcul du Theta
- Calcul du Vega

Les formules sont codées en VBA afin de permettre un recalcul dynamique et une mise à jour automatique des valorisations.

---

### Stratégies étudiées

#### Protective Put (90%)

- Position longue sur l’indice
- Achat d’un Put avec strike à 90%

Objectif :
- Protection contre les baisses
- Limitation du drawdown
- Conservation du potentiel haussier

---

#### Put Spread (90% / 80%)

- Position longue sur l’indice
- Achat d’un Put (90%)
- Vente d’un Put (80%)

Objectif :
- Réduction du coût de couverture
- Protection partielle
- Optimisation du ratio coût / protection

---

### Décomposition du P&L

Le P&L quotidien est analysé via une approximation basée sur les Greeks :

ΔP&L ≈  
Delta × ΔS  
+ ½ Gamma × (ΔS)²  
+ Theta × Δt  
+ Vega × Δσ  

Cette décomposition permet d’identifier :

- Le risque directionnel (Delta)
- L’effet de convexité (Gamma)
- L’érosion temporelle (Theta)
- L’exposition à la volatilité (Vega)

Une comparaison est effectuée entre :
- P&L exact (revalorisation Black-Scholes)
- Approximation par les sensibilités

---

## Résultats principaux

- Performance de l’indice sur la période : −10,85%
- Volatilité réalisée : ~22,60%

Protective Put :
- Protection efficace en cas de baisse
- Coût significatif lié au Theta
- Forte sensibilité au Vega

Put Spread :
- Coût réduit
- Protection limitée sous 80%
- Meilleur compromis coût / couverture

---

## Structure du repository

Option-Pricing-and-Hedging-Excel-VBA  
│  
├── excel_model/      # Fichier Excel (.xlsm) avec implémentation VBA  
├── report/           # Rapport académique détaillé (PDF)  
├── README.md  

---

## Outils utilisés

- Microsoft Excel
- VBA (Visual Basic for Applications)
- Modèle Black-Scholes
- Analyse quantitative du P&L

---

## Limites du modèle

- Hypothèse de volatilité constante
- Absence de coûts de transaction
- Rebalancement discret
- Approximation des Greeks valide pour petits mouvements
