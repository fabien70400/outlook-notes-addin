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

## Hebergement (production)

Les fichiers du volet (`src/taskpane/*`, `assets/*`) sont publies publiquement
via GitHub Pages :

**https://fabien70400.github.io/outlook-notes-addin/**

Le `manifest.xml` du depot pointe deja vers cette URL. Toute modification du
code source (`src/taskpane/`) doit etre commitee et poussee sur la branche
`main` du depot [fabien70400/outlook-notes-addin](https://github.com/fabien70400/outlook-notes-addin)
pour etre publiee (GitHub Pages se redeploie automatiquement en ~1 minute).

## Partager l'add-in avec d'autres utilisateurs

Comme le volet est deja heberge publiquement, **aucune installation Node.js
n'est necessaire pour les autres utilisateurs**. Il suffit de leur transmettre
le fichier [`manifest.xml`](manifest.xml) (par mail, Teams, fichier partage...)
puis, pour chacun :

1. Ouvrir [outlook.office.com](https://outlook.office.com) (ou Outlook
   desktop), ouvrir un mail recu.
2. Menu `...` > **Obtenir des complements** > **Mes complements** >
   **Ajouter un complement personnalise** > **Ajouter a partir d'un fichier**.
3. Selectionner le fichier `manifest.xml` recu.
4. Le bouton **Note de suivi** apparait desormais dans le ruban de lecture des
   mails.

Chaque utilisateur ne voit et ne modifie que les notes sur les mails de **sa
propre boite** (la note est stockee sur le mail via les proprietes Exchange,
propres a chaque mailbox).

### Deploiement centralise (si vous obtenez les droits admin M365)

Si un administrateur Microsoft 365 est disponible, il est preferable de
deployer via le **Centre d'administration Microsoft 365** > *Parametres* >
*Applications integrees* > *Charger un complement personnalise*, en fournissant
la meme URL de manifest. Cela evite a chaque utilisateur de le side-charger
manuellement et permet de cibler des groupes/utilisateurs precis.

## Developper / modifier le code en local

Ces etapes ne sont necessaires que pour modifier le code du volet, pas pour
l'utiliser une fois publie sur GitHub Pages.

### Prerequis

- Node.js 18+
- Outlook desktop (Windows/Mac) connecte a un compte Microsoft 365/Exchange,
  ou Outlook sur le web.
- Un compte GitHub avec acces en ecriture au depot (pour publier les
  changements).

### Installation

```bash
cd outlook-notes-addin
npm install
```

### Tester ses modifications avant de les publier

1. Modifier temporairement `manifest.xml` pour repointer les URLs vers
   `https://localhost:3000` (voir l'historique git pour un exemple), ou tester
   directement en local avant de commit/push :

```bash
npm run dev-server
```

   Au premier lancement, `office-addin-dev-certs` installe un certificat de
   developpement local (une confirmation Windows peut apparaitre).

2. Dans un autre terminal, side-charger l'add-in dans Outlook desktop :

```bash
npm start
```

   Pour arreter/desinstaller : `npm run stop`.

3. Une fois valide, remettre les URLs de `manifest.xml` sur
   `https://fabien70400.github.io/outlook-notes-addin`, puis :

```bash
git add -A
git commit -m "Description du changement"
git push
```

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
- Le depot GitHub est **public** (requis pour GitHub Pages gratuit) : le code
  source du volet est visible par tous, mais aucune note ni donnee
  utilisateur n'y transite (tout reste dans Exchange).
