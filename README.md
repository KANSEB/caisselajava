# Caisse Bar — La Java

Application web simple pour gérer les encaissements au bar pendant La Java.

## Stack

- HTML/CSS/JS pur, un seul fichier `index.html`
- Sync des ventes vers une Google Sheet via Apps Script webhook
- Déployé sur Vercel (déploiement auto à chaque push)

## Configuration

Deux constantes en haut du `<script>` dans `index.html` :

- `SHEETS_WEBHOOK_URL` — URL du Web App Google Apps Script qui reçoit les ventes
- PIN admin par défaut : `1234` (stocké en `localStorage`, changeable depuis le panneau admin)

## Workflow

Modifier `index.html` → `git commit` → `git push` → Vercel redéploie automatiquement.

---

# Caisse Merch — La Java

Deuxième app, **totalement séparée** du bar : `merch.html` (URL `/merch.html`).
Aucune donnée partagée avec le bar (catalogue, panier, file de sync et Google
Sheet indépendants).

## Spécificités

- **Écran d'accès** au stand (PIN équipe, séparé du bar) + PIN admin pour stats/inventaire.
- Encaissement : **Espèces · Carte · Prix préférentiel** (applique le PU Préf du
  doc, modifiable) · **Offert** (0 €) · **Don libre** (somme sans article).
- **Stock par taille** (XS→3XL + TU) avec sélecteur ; « reste X » en temps réel.
- Catalogue et prix **pilotés par la Google Sheet merch** (onglets *Inventaire*
  + *Rentabilité*). Fallback local `SEED` si la Sheet ne répond pas.

## Configuration (`merch.html`, en haut du `<script>`)

- `SHEETS_WEBHOOK_URL` — Web App Apps Script de la Sheet merch (voir `apps-script-merch.gs`)
- `ACCESS_PIN` — code d'accès au stand (bénévoles)
- `ADMIN_PIN` — code admin (stats, inventaire, prix)

## Branchement à la Sheet

Voir l'en-tête de [`apps-script-merch.gs`](apps-script-merch.gs) : coller le
script dans la Sheet, lancer `verifierStructure()`, déployer en Application Web,
puis reporter l'URL `/exec` dans `SHEETS_WEBHOOK_URL`. Le script lit/écrit dans
les onglets *Inventaire* (décrément par taille) et *Rentabilité* (compteurs
Normal/Préf/Don) sans toucher aux colonnes en formule, et journalise chaque
vente dans un onglet *Ventes* créé automatiquement.
