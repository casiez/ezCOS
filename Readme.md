# ezCOS

[![PyPI Version](https://img.shields.io/pypi/v/ezCOS)](https://pypi.org/project/ezCOS/)
[![Downloads](https://static.pepy.tech/badge/ezCOS)](https://pepy.tech/project/ezCOS)

Outil pour faciliter la gestion des comités de sélection (COS) dans les établissements d'enseignement supérieur. Il permet de télécharger les dossiers des candidats depuis Odyssée, de les décompresser et de remplir automatiquement un fichier Excel avec les informations pertinentes sur les candidats... (voir ci-dessous)

ezCOS s'appuie sur [OdysseeAPI](https://github.com/casiez/OdysseeAPI/) pour interagir avec la plateforme Odyssée.

## Installation

[Python](https://www.python.org/) et [pip](https://pip.pypa.io/en/stable/installation/) doivent être installés sur votre machine. Ensuite, installer ezCOS avec la commande suivante :

```
pip install ezCOS -U
```

ezCOS est installé dans le répertoire bin de votre environnement Python, et est accessible depuis la ligne de commande avec la commande `ezCOS`. Assurez-vous que le répertoire correspondant de votre environnement Python est dans votre variable d'environnement PATH pour pouvoir utiliser la commande `ezCOS` depuis n'importe quel répertoire.

Si vous ne trouvez pas où est installé ezCOS, vous pouvez exécuter :

```
pip show ezCOS
```

A partir du chemin indiqué dans la section "Location", vous pouvez remonter l'arborescence jusqu'à trouver le répertoire `bin` dans un sous-répertoire.

Créer un répertoire de travail pour le comité de sélection, puis créer le fichier `config.ini` avec les informations ci-dessous, copier les modèles de fichiers Excel et Word depuis le répertoire `templates` dans un répertoire `templates`. Copier également le contenu du répertoire `carbone`dans un répertoire du même nom.

```
[ODYSSEE]
numposte = 123456
login = xxx@yyyy
password = xxxxx

[FILES]
xlsx-candidats = monCOS.xlsx
convocations-candidats = chemin absolu vers le répertoire convocationsCandidats

[SOFTWARE]
soffice = /Applications/LibreOffice.app/Contents/MacOS/soffice
thunderbird-bin = /Applications/Thunderbird.app/Contents/MacOS/thunderbird
```

La génération de documents s'appuie sur la librairie [Carbone](https://carbone.io/). Il faut installer [node.js](https://nodejs.org/) et [npm](https://www.npmjs.com/), puis installer Carbone avec la commande ```npm install carbone```. Si cette librairie n'est pas installée, les autres commandes (téléchargement dossiers...) restent fonctionnelles.

OdysseeAPI s'appuie sur [playwright](https://playwright.dev/python/). Il faut installer les navigateurs supportés par Playwright avec la commande suivante (à faire une seule fois après l'installation de ezCOS) :
```
playwright install
```

## Etapes

A tout moment, vous pouvez exécuter la commande `ezCOS -h` pour afficher l'aide et les différentes options disponibles. L'option `-updateCache` permet de forcer la mise à jour du cache des données d'Odyssée, ce qui peut être utile si les données ont changé depuis la dernière exécution.

1. Télécharger les dossiers des candidats depuis Odyssee :
   ```
   ezCOS -auth -downloadApplications
   ```
   L'option `-auth` permet de s'authentifier auprès d'Odyssee pour accéder aux dossiers des candidats. L'option `-downloadApplications` télécharge les archives zip des dossiers dans le répertoire `dossiers_candidats`.

1. Décompresser les dossiers des candidats :
   ```
   ezCOS -unzipApplications
   ```
   L'option `-unzipApplications` décompresse les archives zip téléchargées dans le répertoire `dossiers_candidats`.

1. Remplir automatiquement un fichier Excel avec les informations sur les candidats :
   ```
   ezCOS -auth -populateXLSX
   ```
   L'option `-populateXLSX` complète le fichier Excel avec les informations d'Odyssée ainsi que des informations extraites des dossiers des candidats (fichiers qui commencent par `Edition_Dossier`). Le fichier `templates/comite.xlsx` est utilisé comme modèle pour créer le fichier `comitePopulated.xlsx` qui contient les informations des candidats et des membres du comité.

   A cette étape, il faut renommer le fichier `comitePopulated.xlsx` en utilisant celui défini dans la section [FILES] du fichier `config.ini` (ex: `comite.xlsx`).

   Dans la feuille "Candidats" du fichier Excel, affecter les rapporteurs à chaque candidat en utilisant les listes déroulantes dans la colonne "Rapporteur1" et "Rapporteur2". Les membres du comité sont listés dans la feuille "MembresComite" du fichier Excel.

   Dans la feuille "MembresComite", compléter la section et le statut de chaque membre du comité (Interne / Externe).

1. Générer les avis d'impartialité pour les membres du comité :
   ```
   ezCOS -impartialite
   ```
   L'option `-impartialite` génère les déclarations d'impartialité pour les membres du comité, à partir du modèle `templates/Impartialite.xlsx`. Les déclarations sont sauvegardées dans le répertoire `impartialite`.

1. Génération des modèles de rapports pré-complétés pour les rapporteurs :
   ```
   ezCOS -r1
   ```
   L'option `-r1` génère les modèles de rapports pré-complétés pour les rapporteurs avec les infos des candidats, à partir du modèle `templates/NOMCANDIDATprenom-NOMRAPPORTEUR.docx`. Les rapports sont générés au format DOCX et sauvegardés dans le répertoire `rapportsMembresReunion1`.

1. Affectation des rapporteurs sur Odyssee :
   ```
   ezCOS -auth -assignOdyssee
   ```
   L'option `-assignOdyssee` affecte les rapporteurs sur Odyssée, en utilisant les informations du fichier Excel. Les rapporteurs sont affectés aux candidats sur Odyssée en fonction des choix faits dans les colonnes "Rapporteur1" et "Rapporteur2" du fichier Excel.

   Vérifier le résultat de l'affectation des rapporteurs sur Odyssée.

1. Génération du PV de répartition des rapporteurs :
   ```
   ezCOS -pvRepart
   ```
   L'option `-pvRepart` génère le PV de répartition des rapporteurs, à partir du modèle `templates/PVrepartition.docx`. Le PV est sauvegardé sous le nom `PVrepartitionRapporteurs.docx`.

1. Génération des attestations de participation en visio pour les membres du comité :
   ```
   ezCOS -attestationsVisio
   ```
   L'option `-attestationsVisio` génère les attestations de participation en visio pour les membres du comité, à partir du modèle `templates/Attestation_visioconference_membre.docx`. Les attestations sont générées au format DOCX et sauvegardées dans le répertoire `attestationsVisio`.

1. Téléchargement des rapports des rapporteurs depuis Odyssée :
   ```
   ezCOS -auth -reports
   ```
   L'option `-reports` télécharge les rapports des rapporteurs depuis Odyssée. Les rapports sont téléchargés au format PDF et sauvegardés dans le répertoire `rapportsOdyssee`.

1. Extraction des avis des rapports des rapporteurs et mise à jour du fichier Excel :
   ```
   ezCOS -reports2xlsx
   ```
   L'option `-reports2xlsx` extrait les avis des rapports des rapporteurs et met à jour le fichier Excel. Les avis sont extraits des rapports PDF téléchargés à l'étape précédente, puis reportés dans les colonnes "AvisRapp1" et "AvisRapp2" du fichier Excel. Par précaution, le script crée une copie du fichier Excel avant de le modifier, avec le suffixe "-2". Les avis sont ensuite à copier dans le fichier Excel original. Cette option reste expérimentale et peut produire des erreurs d'extraction.

1. Avis par candidat et vote suite à la première réunion du comité de sélection :
   ```
   ezCOS -auth -avisPremiereReunion
   ```
   L'option `-avisPremiereReunion` utilise la colonne "AvisAudition" du fichier Excel pour générer un avis global. Les lignes de cette colonne doivent être remplies avec des codes (ex: "A", "B", "C", "D") séparés par des virgules, et définis dans la feuille AvisDetailles du fichier Excel. Les candidats auditionnés doivent avoir au moins un avis "A" dans cette colonne.
   
   Les résultats du vote sont dans les colonnes correspondantes du fichier Excel.

   Vérifier le résultat de la mise à jour des avis et des votes sur Odyssée.

1. Génération des convocations pour les candidats auditionnés :
   ```
   ezCOS -convocCandidats
   ```
   L'option `-convocCandidats` génère les convocations pour les candidats auditionnés, à partir du modèle `templates/Convocation_candidats.docx`. Les convocations sont générées au format DOCX et sauvegardées dans le répertoire `convocationsAuditions`.

1. Conversion au format PDF des convocations pour les candidats auditionnés :
   ```
   ezCOS -convocCandidatsPDF
   ```
   L'option `-convocCandidatsPDF` convertit au format PDF les convocations pour les candidats auditionnés, à partir des fichiers DOCX générés à l'étape précédente. Les convocations au format PDF sont sauvegardées dans le répertoire `convocationsAuditions`.

1. Envoi des convocations aux candidats auditionnés par email :
   ```
   ezCOS -envoimailauditionnes
   ```
   L'option `-envoimailauditionnes` prépare les convocations aux candidats auditionnés par email, en utilisant les adresses email extraites des dossiers des candidats. Les convocations au format PDF ajoutées en pièce jointe.

   Le template `templates/mailCandidatsAuditionnes.txt` est utilisé pour le corps du mail.

   Les mails ne sont pas envoyés directement, mais préparés dans le client de messagerie Thunderbird. Vérifier les mails préparés dans Thunderbird avant de les envoyer.

   Vérifier que vous avez bien configuré `convocations-candidats` dans le fichier de config.ini (mettre le chemin absolu), ainsi que le chemin vers thunderbird (`thunderbird-bin`).