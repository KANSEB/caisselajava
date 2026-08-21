# Caisses — La Java

| App | URL bénévoles | Fichier |
|---|---|---|
| Bar | **https://lajava-festival.fr/appbar** | `index.html` |
| Merch | **https://lajava-festival.fr/appmerch** | `merch.html` |

Les deux URLs sont des rewrites dans `next.config.ts` du site officiel
(repo `kanseb/lajava`) qui proxyfient vers ce projet Vercel
(`caisselajava.vercel.app`). Un push ici suffit à mettre à jour les caisses,
sans rebuild du site.

# Caisse Bar — La Java

Application web de caisse pour le bar pendant La Java. Un seul fichier
`index.html`, servi à la racine du site (URL `/`).

## Spécificités

- **Aucun login** — n'importe quel bénévole ouvre l'URL et encaisse direct.
- **Grille produits** codée par couleur (bières / vins / softs / consigne).
  Chaque touche = ajout au ticket ; « +1 » flottant + vibration tactile.
- **Ticket** affiché en haut, scroll interne jusqu'à 38 vh pour rester
  utilisable sur mobile même avec beaucoup d'articles.
- **Encaissement** : deux modes — **🎟️ Tickets** · **💳 Carte** (pas
  d'espèces au bar).
- **Panneau admin** protégé par PIN 4 chiffres (validation auto au 4e chiffre).
  Édition des prix (par device, en `localStorage`) et **stats globales**
  agrégées depuis la Google Sheet (totaux tickets/carte + %, nb
  d'encaissements, total caisse), rafraîchissables à la demande.
- **File d'attente offline** : si la Sheet est injoignable, la vente est
  gardée en `localStorage` et re-tentée toutes les 30 s / dès retour online.
- **Pastille sync** dans le header (vert / rouge / pulse jaune).

## Configuration (`index.html`, en haut du `<script>`)

- `SHEETS_WEBHOOK_URL` — Web App Apps Script de la Sheet bar
- `ADMIN_PIN` — code admin **fixé dans le source** (partagé sur tous les
  appareils ; pour le changer, éditer la constante puis `git push`)
- `PRODUCTS_STORAGE_KEY` — bumper cette clé quand `DEFAULTS` change pour
  forcer le rechargement du catalogue sur tous les appareils en cache

## Google Sheet

L'Apps Script attaché à la Sheet expose deux endpoints :

- `POST /exec` — ajoute une ligne : *Date · Méthode · Montant · Articles · Device*
- `GET /exec?action=stats` — renvoie `{tickets, card, countTickets, countCard}`
  sommés depuis la Sheet

Déploiement à faire en **Application Web · Exécuter en tant que Moi · Qui a
accès : Tout le monde** (sinon les POST des bénévoles sont bloqués).

## Workflow

Modifier `index.html` → `git commit` → `git push` → Vercel redéploie
automatiquement en ~30 s.

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
