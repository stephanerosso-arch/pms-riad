# Données d'exemple (dossier data)

Ce dossier est destiné à contenir des exports d'exemples (sandbox) pour tester les scripts Google Apps Script du projet sans toucher aux données de production.

Important :
- Ne placez ici que des copies anonymisées des feuilles (emails, noms et données sensibles supprimés ou remplacés par des valeurs factices).
- Si vous avez besoin que je déplace `recupMail.xlsx` depuis la racine vers `data/`, dites‑le — je peux le faire, mais je ne vais pas supprimer la version racine sans votre confirmation.

Comment déplacer le fichier localement (si vous préférez le faire vous‑même) :

```bash
# clone
git clone git@github.com:stephanerosso-arch/pms-riad.git
cd pms-riad
# déplacer le fichier vers le dossier data/
mkdir -p data
git mv recupMail.xlsx data/recupMail.xlsx
git commit -m "Move sample recupMail.xlsx into data/ folder"
git push
```

Si vous voulez que je fasse le déplacement dans le dépôt, confirmez ici et je l'effectuerai pour vous (je créerai `data/recupMail.xlsx`).
