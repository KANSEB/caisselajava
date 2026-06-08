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
