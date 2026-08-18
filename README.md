# Notes de suivi - Add-in Outlook

Add-in Outlook qui ajoute un volet permettant d'attacher une note libre a
n'importe quel mail recu, pour faciliter le suivi (relances, statut,
rappels...).

## Fonctionnement

- La note est stockee directement sur le mail via les "custom properties"
  Office.js (`item.loadCustomPropertiesAsync` / `customProperties.saveAsync`).
- Aucune base de donnees ni serveur externe : la note voyage avec le mail sur
  la boite Exchange/Outlook (bureau, web, mobile).
- Seul cet add-in peut lire ces proprietes ; elles ne sont pas visibles par
  les destinataires ni par d'autres personnes consultant la boite.
- Un second onglet **Suivi** recapitule les mails annotes. Cette liste est un
  index local (stocke dans le navigateur du volet, sur ce poste) qui pointe
  vers les mails ; elle se remplit au fur et a mesure que vous enregistrez des
  notes depuis cet ordinateur. Cliquer sur une ligne rouvre le mail
  correspondant. Contrairement a la note elle-meme, cette liste **ne se
  synchronise pas** entre plusieurs postes.

## Prerequis

- Node.js 18+
- Outlook desktop (Windows/Mac) connecte a un compte Microsoft 365/Exchange,
  ou Outlook sur le web.

## Installation

```bash
cd outlook-notes-addin
npm install
```

## Lancer en local

1. Demarrer le serveur de dev (sert les fichiers en HTTPS sur le port 3000) :

```bash
npm run dev-server
```

   Au premier lancement, `office-addin-dev-certs` installe un certificat de
   developpement local (une confirmation Windows peut apparaitre).

2. Dans un autre terminal, side-charger l'add-in dans Outlook desktop :

```bash
npm start
```

   Cela ouvre Outlook et installe automatiquement l'add-in a partir de
   `manifest.xml`. Ouvrez un mail recu : un bouton **Note de suivi** apparait
   dans le ruban, il ouvre le volet lateral.

   Pour arreter/desinstaller : `npm run stop`.

### Alternative : side-chargement manuel (Outlook sur le web)

Si `npm start` ne fonctionne pas pour votre configuration :

1. Sur [outlook.office.com](https://outlook.office.com), ouvrir un mail.
2. Menu `...` > **Obtenir des compléments** > **Mes compléments** >
   **Ajouter un complement personnalise** > **Ajouter à partir d'un fichier**.
3. Selectionner `manifest.xml` (le serveur `npm run dev-server` doit tourner).

## Utilisation

1. Ouvrir un mail recu.
2. Cliquer sur **Note de suivi** dans le ruban : le volet s'ouvre toujours sur
   l'onglet **Note** du mail actuellement ouvert.
3. Saisir une note (ex: "Relancer le 25/08, en attente de devis"), cliquer sur
   **Enregistrer**.
4. La note reste attachee au mail : elle reapparait a chaque reouverture, y
   compris sur un autre appareil.

### Garder le volet ouvert en changeant de mail (epingler)

Outlook ne permet pas d'ouvrir automatiquement le volet des l'ouverture
d'Outlook (ce n'est pas propose par la plateforme Office Add-ins). En
revanche, une fois le volet ouvert, une icone **epingle** apparait en haut du
volet (a cote du bouton fermer) :

- **Epingle active** : le volet reste ouvert quand vous passez d'un mail a
  l'autre, sans avoir a recliquer sur le bouton du ruban a chaque fois.
- **Epingle desactivee** : le volet se ferme des que vous changez de mail
  (comportement par defaut).

Cette fonctionnalite est native a Outlook (disponible sur Outlook nouveau/web
et Outlook 2016+ recent) : voir la
[documentation Microsoft sur les pinnable task panes](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/pinnable-taskpane).

## Structure du projet

```
outlook-notes-addin/
├── manifest.xml              # Declaration de l'add-in pour Outlook
├── app/server.js             # Serveur HTTPS de dev (sert les fichiers statiques)
├── src/taskpane/
│   ├── taskpane.html
│   ├── taskpane.css
│   └── taskpane.js           # Logique de lecture/ecriture de la note
└── assets/                   # Icones du ruban
```

## Limites connues

- Les notes ne sont pas recherchables ni listables globalement (pas de vue
  "tous les mails annotes"). Si ce besoin apparait, il faudra migrer vers un
  stockage externe (API + base de donnees).
- Fonctionne sur Outlook desktop et Outlook web ; a valider specifiquement
  sur Outlook mobile si besoin.
- Pour un usage en production (hors poste de dev), il faudra heberger les
  fichiers statiques sur un vrai serveur HTTPS et deployer le manifest via le
  Centre d'administration Microsoft 365 (deploiement centralise) plutot que
  le side-chargement.
