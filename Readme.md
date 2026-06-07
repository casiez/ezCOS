# ezCOS

Outil pour faciliter la gestion des comités de sélection (COS) dans les établissements d'enseignement supérieur. Il permet de télécharger les dossiers des candidats depuis Odyssee, de les décompresser et de remplir automatiquement un fichier Excel avec les informations pertinentes sur les candidats... (voir ci-dessous)

ezCOS s'appuie sur [odysseeapi](https://github.com/casiez/OdysseeAPI/) pour interagir avec la plateforme Odyssee.

## Installation


## Etapes

1. Télécharger les dossiers des candidats depuis Odyssee :
   ```
   python ezCOS.py -auth -downloadApplications
   ```
   L'option `-auth` permet de s'authentifier auprès d'Odyssee pour accéder aux dossiers des candidats. L'option `-downloadApplications` lance le processus de téléchargement des dossiers. Les archives zip des dossiers sont téléchargés dans le répertoire `dossiers_candidats`.

1. Décompresser les dossiers des candidats :
   ```
   python ezCOS.py -unzipApplications
   ```
   L'option `-unzipApplications` lance le processus de décompression des archives zip téléchargées. Les dossiers décompressés se trouvent dans le répertoire `dossiers_candidats`.

1. Remplir automatiquement un fichier Excel avec les informations sur les candidats :
   ```
   python ezCOS.py -auth -populateXLSX
   ```
   L'option `-populateXLSX` lance le processus de remplissage du fichier Excel avec les informations extraites des dossiers des candidats. Le fichier `templates/comite.xlsx` est utilisé comme modèle pour créer le fichier `comitePopulated.xlsx` qui contient les informations des candidats et des membres du comité.

   A cette étape, il faut renommer le fichier `comitePopulated.xlsx` en utilisant celui défini dans la section [FILES] du fichier `config.ini` (ex: `comite.xlsx`).

   Dans la feuille "Candidats" du fichier Excel, affecter les rapporteurs à chaque candidat en utilisant les listes déroulantes dans la colonne "Rapporteur1" et "Rapporteur2". Les membres du comité sont listés dans l'onglet "MembresComite" du fichier Excel.

   Dans l'onglet "MembresComite", compléter la section et le statut de chaque membre du comité (Interne / Externe).

1. Générer les avis d'impartialité pour les membres du comité :
   ```
   python ezCOS.py -impartialite
   ```
   L'option `-impartialite` lance le processus de génération des déclarations d'impartialité pour les membres du comité, à partir du modèle `templates/Impartialite.xlsx`. Les décalrations sont sauvegardées dans le répertoire `impartialite`.

1. Génération des modèles de rapports pré-complétés pour les rapporteurs :
   ```
   python ezCOS.py -r1
   ```
   L'option `-r1` lance le processus de génération des modèles de rapports pré-complétés pour les rapporteurs, avec les infos des candidats et à partir du modèle `templates/NOMCANDIDATprenom-NOMRAPPORTEUR.docx`. Les rapports sont générés au format DOCX et sauvegardés dans le répertoire `rapportsMembresReunion1`.

1. Affectation des rapporteurs sur Odyssee :
   ````
   python ezCOS.py -auth -assignOdyssee
   ```
   L'option `-assignOdyssee` lance le processus d'affectation des rapporteurs sur Odyssee, en utilisant les informations du fichier Excel. Les rapporteurs sont affectés aux candidats sur Odyssee en fonction des choix faits dans les colonnes "Rapporteur1" et "Rapporteur2" du fichier Excel.

   Vérifier le résultat de l'affectation des rapporteurs sur Odyssee.

1. Génération du PV de répartition des rapporteurs :
   ```
   python ezCOS.py -pvRepart
   ```
   L'option `-pvRepart` lance le processus de génération du PV de répartition des rapporteurs, à partir du modèle `templates/PVrepartition.docx`. Le PV est sauvegardé sous le nom `PVrepartitionRapporteurs.docx`.

1. Génération des attestations de participation en visio pour les membres du comité :
   ```
   python ezCOS.py -attestationsVisio
   ```
   L'option `-attestationsVisio` lance le processus de génération des attestations de participation en visio pour les membres du comité, à partir du modèle `templates/Attestation_visioconference_membre.docx`. Les attestations sont générées au format DOCX et sauvegardées dans le répertoire `attestationsVisio`.

1. Téléchargement des rapports des rapporteurs depuis Odyssee :
   ```
   python ezCOS.py -auth -reports
   ```
   L'option `-reports` lance le processus de téléchargement des rapports des rapporteurs depuis Odyssee. Les rapports sont téléchargés au format PDF et sauvegardés dans le répertoire `rapportsOdyssee`.

1. Extraction des avis des rapports des rapporteurs et mise à jour du fichier Excel :
   ```
   python ezCOS.py -reports2xlsx
   ```
   L'option `-reports2xlsx` lance le processus d'extraction des avis des rapports des rapporteurs et de mise à jour du fichier Excel. Les avis sont extraits des rapports PDF téléchargés à l'étape précédente, puis reportés dans les colonnes "AvisRapp1" et "AvisRapp2" du fichier Excel. Par précaution, le script crée une copie du fichier Excel avant de le modifier, avec le suffixe "-2". Les avis sont ensuite à copier dans le fichier Excel original.

1. Avis par candidat et vote suite à la première réunion du comité de sélection :
   ```
   python ezCOS.py -auth -avisPremiereReunion
   ```
   L'option `-avisPremiereReunion` utilise la colonne "AvisAudition" du fichier Excel pour générer un avis global. Les lignes de cette colonne doivent être remplies avec des codes (ex: "A", "B", "C", "D") séparés par des virgules, et définis dans la feuille AvisDetailles du fichier Excel. Les candidats auditionnés doivent avoir au moinns un avis "A" dans cette colonne.
   
   Les résultats du vote sont dans les colonnes correspondantes du fichier Excel.

   Vérifier le résultat de la mise à jour des avis et des votes sur Odyssee.

1. Génération des convocations pour les candidats auditionnés :
   ```
   python ezCOS.py -convocCandidats
   ```
   L'option `-convocCandidats` lance le processus de génération des convocations pour les candidats auditionnés, à partir du modèle `templates/Convocation_candidats.docx`. Les convocations sont générées au format DOCX et sauvegardées dans le répertoire `convocationsAuditions`.

1. Conversion au format PDF des convocations pour les candidats auditionnés :
   ```
   python ezCOS.py -convocCandidatsPDF
   ```
   L'option `-convocCandidatsPDF` lance le processus de conversion au format PDF des convocations pour les candidats auditionnés, à partir des fichiers DOCX générés à l'étape précédente. Les convocations au format PDF sont sauvegardées dans le répertoire `convocationsAuditions`.

1. Envoi des convocations aux candidats auditionnés par email :
   ```
   python ezCOS.py -envoimailauditionnes
   ```
   L'option `-envoimailauditionnes` lance le processus d'envoi des convocations aux candidats auditionnés par email, en utilisant les adresses email extraites des dossiers des candidats. Les convocations au format PDF sont envoyées en pièce jointe.

   Les mails ne sont pas envoyés directement, mais préparés dans le client de messagerie Thunderbird. Vérifier les mails préparés dans Thunderbird avant de les envoyer.

   Vérifier que vous avez bien configuré `convocations-candidats` dans le fichier de config.ini (mettre le chemin absolu), ainsi que le chemin vers thunderbird (`thunderbird-bin`).