# ============================================================
#  ____  ____      _   _  ____ ____  _   _
# |  _ \/ ___|    | \ | |/ ___|  _ \| | | |
# | |_) \___ \    |  \| | |   | | | | | | |
# |  __/ ___) |   | |\  | |___| |_| | |_| |
# |_|   |____/    |_| \_|\____|____/ \___/
#
# ============================================================
#  Script      : PS-NCDU Web Edition (arbre live navigable)
#  Description : Disk Usage Analyzer - Application web locale
#  Version     : 6.28
#  Date        : 2026-09-15
#  Auteur      : Eric Guiffault (eric@guiffault.com)
#  Societe     : CL SASU
# ------------------------------------------------------------
#  Compatibilite : PowerShell 5.1+ | FullLanguage REQUIS (serveur web)
#  Dependances   : Aucune - 100% PowerShell natif (HttpListener)
# ------------------------------------------------------------
#  NOUVEAUTES v6.28 :
#    I18N8 : Localisation complete. Les 15 dernieres cles (explications de
#            profondeur, du filtre, de la fenetre d'analyse, des exclusions,
#            messages de file et de fichiers) sont traduites dans les 30
#            langues recentes. Les 42 langues portent desormais les 89 cles ;
#            plus aucun repli sur l'anglais. Relecture native toujours
#            bienvenue pour les langues les moins courantes.
#  NOUVEAUTES v6.27 :
#    FIX14 : Le bouton de tri (et la profondeur au pied) ne se retraduisaient
#            pas quand on changeait de langue AVANT le premier scan. Ces
#            libelles sont ecrits par script (pas par data-i18n) et n'etaient
#            rafraichis que via renderTree, inactif sans scan. setLang les
#            rafraichit desormais directement.
#  NOUVEAUTES v6.26 :
#    I18N7 : 17 cles ajoutees au fil des versions (profondeur du scan, tri,
#            explorateur, exclusions, Interrompre, lecteurs libres, etc.)
#            n'existaient qu'en anglais et en francais. Elles sont desormais
#            traduites dans les 10 autres langues d'origine (zh hi es ar bn
#            pt ru ur id de). Plus aucune langue d'origine ne manque de cle.
#  NOUVEAUTES v6.25 :
#    I18N6 : Libelles de navigation et de scan traduits pour les 30 langues
#            ajoutees (dossier parent, Pret, Termine, En cours, etats des
#            dossiers, messages d'erreur, seuils d'affichage, etc.). Il ne
#            reste en repli anglais, pour ces langues, que quelques longs
#            paragraphes d'explication (aide de profondeur, filtre, tooltips).
#            Traductions au meilleur effort, relecture native conseillee.
#  NOUVEAUTES v6.24 :
#    I18N5 : Legende et aide traduites pour les 30 langues ajoutees en v6.23
#            (etats des dossiers, fichier, barre, et les notes d'aide). Le
#            panneau Legende / Aide n'est donc plus en anglais dans ces
#            langues. Traductions au meilleur effort, relecture native
#            conseillee pour les langues les moins courantes.
#  NOUVEAUTES v6.23 :
#    SORT2 : Tri par taille par defaut (y compris pendant le scan, lisse).
#            Le bouton bascule vers le tri par nom si besoin.
#    I18N4 : 30 langues supplementaires (42 au total) pour couvrir plus de
#            80%% de la population mondiale : ja ko it tr vi pl nl uk ro cs
#            el sv hu fa th ms fil sw ta te mr gu kn ml pa he ha my am km.
#            Sens d'ecriture RTL applique aussi pour le persan et l'hebreu.
#            Le coeur de l'interface est traduit ; les textes longs (aide,
#            explications) retombent sur l'anglais pour ces nouvelles langues.
#  NOUVEAUTES v6.22 :
#    SORT1 : Bouton de tri Nom / Taille. Le scan avance dans l'ordre
#            alphabetique (remplissage stable de haut en bas), mais on veut
#            aussi voir vite les gros dossiers. Compromis : par defaut tri
#            par nom pendant le scan et par taille a la fin ; un bouton force
#            l'un ou l'autre a tout moment. En tri par taille pendant le scan,
#            le re-classement est lisse (au plus une fois par seconde) pour
#            que les gros elements remontent sans que la liste saute sans arret.
#            Choix memorise.
#  NOUVEAUTES v6.21 :
#    FIX13 : Progression de haut en bas. Le serveur scannait les dossiers de
#            premier niveau dans l'ordre brut du systeme de fichiers, alors
#            que la page les affiche tries par nom. Les deux ordres divergaient,
#            d'ou un remplissage qui sautait partout et semblait commencer au
#            milieu. Les L1 sont desormais tries cote serveur dans le meme
#            ordre que l'affichage (ordinal, minuscules) : remplissage regulier.
#  NOUVEAUTES v6.20 :
#    BROWSE1 : Explorateur de dossiers integre. Le dialogue Windows natif
#              etant peu fiable pour un serveur local (regles de premier
#              plan, et faux sens en acces distant), "Parcourir" ouvre
#              desormais un explorateur rendu dans la page : lecteurs,
#              navigation par clic, dossier parent, puis "Choisir ce
#              dossier". Route serveur /api/browse. Marche a coup sur.
#    ICONS1  : Icones par type de fichier. Une table couvrant ~120 des
#              extensions les plus courantes (images, video, audio, PDF,
#              bureautique, archives, code, exe, polices, etc.) remplace
#              l'icone unique, pour identifier les types d'un coup d'oeil.
#  NOUVEAUTES v6.19 :
#    FIX12 : Selecteur de dossier natif repare. Il etait montre depuis un
#            thread de fond dont le scriptblock revenait vers le runspace
#            principal (bloque) -> blocage, rien a l'ecran. Desormais : en
#            PS 5.1 (console STA) le dialogue s'ouvre directement sur le
#            thread courant ; en PS 7 (MTA) via un runspace STA dedie. Une
#            fenetre proprietaire TopMost le force au premier plan (sinon il
#            s'ouvrait derriere le navigateur). Si l'ouverture echoue, un
#            message invite a saisir le chemin a la main.
#  NOUVEAUTES v6.18 :
#    STOP1 : Interruption d'un scan. Un bouton "Interrompre" apparait
#            pendant le scan. Comme le moteur est mono-thread bloquant,
#            l'annulation est cooperative : fermer la connexion (bouton
#            Interrompre, ou lancer un nouveau scan) fait echouer la
#            prochaine ecriture SSE du serveur, ce qui leve un drapeau
#            $script:CancelScan verifie dans les boucles (E1, E3,
#            Register-Subtree). Le scan s'arrete alors de lui-meme et le
#            serveur redevient disponible en ~1 s.
#  NOUVEAUTES v6.17 :
#    PERF5 : Scan illimite entrelace. Avant, E1 enumerait TOUT le disque
#            (Get-ChildItem -Recurse) avant qu'une seule taille n'apparaisse
#            (plusieurs minutes sans rien a l'ecran sur C:\). Desormais E1
#            n'enumere que le 1er niveau, puis chaque sous-arbre est enumere
#            juste avant d'etre mesure (E3) : les tailles se remplissent
#            dossier par dossier en quelques secondes. Enumeration via .NET
#            streaming (plus rapide que Get-ChildItem).
#    UX12  : Profondeur du scan (illimitee / N) affichee dans le pied
#            pendant le scan, pour savoir tout de suite qu'un scan complet
#            est en cours et pourquoi il est long.
#  NOUVEAUTES v6.16 :
#    UX11 : Fenetre d'analyse reorganisee et plus compacte. Deux colonnes
#           (destination a gauche, options a droite), selecteur de langue
#           remonte dans l'en-tete, espacements resserres. Repli sur une
#           seule colonne sur petit ecran.
#  NOUVEAUTES v6.15 :
#    REC1 : Gestion des recents. Dedoublonnage a l'affichage, retrait
#           individuel (croix) et effacement complet de l'historique.
#    EXC1 : Exclusions visibles et personnalisables. Section depliable
#           listant les dossiers systeme toujours ignores, plus un champ
#           pour en exclure d'autres le temps d'un scan.
#  NOUVEAUTES v6.14 :
#    UX10 : Fenetre d'analyse plus ergonomique. Fermeture par Echap ou par
#           une croix (quand un scan est deja affiche), Entree pour lancer,
#           validation du chemin en direct (pastille verte/rouge) et au
#           lancement, memorisation des derniers reglages (chemin,
#           profondeur, filtre), barre d'occupation par lecteur.
#  NOUVEAUTES v6.13 :
#    ACC1 : Accents. L'interface et les sorties console portent desormais
#           tous leurs accents en francais, conformement a la regle projet.
#  NOUVEAUTES v6.12 :
#    I18N3: Selecteur de langue dans la fenetre d'analyse (12 langues).
#           Le choix est memorise (localStorage) et prime sur la detection
#           du systeme au chargement suivant. RTL applique pour ar/ur.
#  NOUVEAUTES v6.11 :
#    FIX11: Antislash quadruples corriges a la racine. ConvertTo-JsonSafe
#           sur-echappait (un antislash -> quatre au lieu de deux), ce qui
#           doublait a chaque re-scan d'un raccourci recent. Corrige +
#           Normalize-Path et affichage nettoient les entrees deja corrompues.
#  NOUVEAUTES v6.10 :
#    UX9  : Tailles affichees avec deux decimales (ex. 23,71 Go), separateur
#           decimal adapte a la langue. Plus de granularite.
#  NOUVEAUTES v6.9 :
#    FIX10: Vraie cause du "disparaissent" : le tri par taille reordonnait
#           en continu pendant le scan (un dossier grossit -> remonte ;
#           petit en octets mais riche en fichiers -> file en bas, hors vue).
#           Desormais ordre stable (alphabetique) pendant le scan, tri par
#           taille une fois le scan termine. Plus rien ne saute.
#  NOUVEAUTES v6.8 :
#    FIX9 : Les repertoires ne "disparaissent" plus pendant le scan. La
#           taille partielle est memorisee (croissante) sur le noeud, donc
#           un dossier garde sa place et sa taille au lieu de retomber en
#           bas de liste quand le pointeur de scan passe au suivant.
#  NOUVEAUTES v6.7 :
#    UX8  : Animation de chargement (balayage) sur les lignes en attente,
#           pour montrer clairement que ca travaille pendant les longues
#           phases de decouverte/calcul (gros disques > 60 s).
#    FIX8 : Noms normalises a l'affichage (ex. racine "D:\" au lieu de
#           double antislash), meme sur-echappement contourne.
#  NOUVEAUTES v6.6 :
#    UX7  : Avancement en direct du repertoire actif a TOUS les niveaux.
#           L'evenement 'scanning' porte les partiels de chaque ancetre
#           suivi, donc quand on descend dans le repertoire en cours, le
#           sous-repertoire actif montre aussi sa taille/compteur en direct.
#  NOUVEAUTES v6.5 :
#    UX6  : Sur la ligne du repertoire en cours de scan, avancement en
#           direct (taille et nombre de fichiers deja lus dans ce
#           repertoire), mis a jour chaque seconde, au lieu de "scan...".
#  NOUVEAUTES v6.4 :
#    FIX7 : Pourcentages bornes a 100 %% (bug 11959 %%). Le total et les
#           compteurs du dossier courant sont calcules depuis son contenu
#           (enfants), robuste quand la taille remontee est encore partielle.
#  NOUVEAUTES v6.3 :
#    UI5  : Numero de version affiche en clair dans l'en-tete.
#    GREY1: Bouton "Scanner les dossiers gris" dans l'en-tete quand le
#           dossier courant en contient. Les met tous en file et les
#           scanne un a un, en restant sur le dossier courant.
#  NOUVEAUTES v6.2 :
#    CNT2 : Par dossier, affichage du nombre de sous-repertoires ET de
#           fichiers (comptage recursif). Sous-repertoires calcules cote
#           client depuis la structure, sans surcout serveur.
#  NOUVEAUTES v6.1 :
#    CNT1 : Nombre de fichiers par dossier (comptage recursif) affiche
#           sur chaque ligne et dans l'en-tete du dossier courant.
#           Les fichiers directs de la racine sont inclus dans le total.
#  NOUVEAUTES v6.0 :
#    I18N1: Interface multilingue. Detection de la langue du systeme
#           (CurrentUICulture), repli sur l'anglais si non disponible.
#    I18N2: Traduit dans les 12 langues les plus parlees : anglais,
#           mandarin, hindi, espagnol, francais, arabe, bengali,
#           portugais, russe, ourdou, indonesien, allemand. RTL pour
#           l'arabe et l'ourdou.
#  NOUVEAUTES v5.15 :
#    LOG1 : Le log "[WEB] Scan termine" affiche le vrai nombre de
#           dossiers (etait fige a 0 depuis que E5 est saute en v5.10).
#  NOUVEAUTES v5.14 :
#    FIX6 : Les lecteurs RESEAU ne sont plus interroges (taille /
#           disponibilite). Via VPN, un partage lent ou mort bloquait
#           cet appel au chargement, gelant le serveur mono-thread et
#           empechant tout scan de demarrer (ecran vide, statut fige).
#  NOUVEAUTES v5.13 :
#    QUEUE1: Dossier gris clique pendant un scan : point BLEU "en file"
#            (distinct de l'orange "prevu dans le scan en cours").
#    LEG2  : Legende du footer passee en panneau a onglets (Legende /
#            Aide), avec l'etat "en file" et un guide d'utilisation.
#    NOTE  : la permutation (mettre le scan principal en pause pour
#            prioriser un gris) n'est pas faisable : moteur bloquant
#            mono-thread, non pausable. La file reste la reponse.
#  NOUVEAUTES v5.12 :
#    FILES1: Les fichiers sont maintenant listes. A l'ouverture d'un
#            dossier, ses fichiers directs sont charges a la demande,
#            fusionnes avec les sous-dossiers et tries par taille
#            (plafonne aux 1000 plus gros). Le disque n'est pas relu en
#            entier : seulement le dossier ouvert.
#    LEG1  : Legende depliable dans la barre de statut (sens des points
#            de couleur, anneau, dossier/fichier, barre de proportion).
#    CONC1 : Rappel : un seul scan a la fois (le disque ne gagne rien a
#            en paralleliser deux). Un clic gris pendant un scan est mis
#            en file et lance a la fin.
#  NOUVEAUTES v5.11 :
#    UI3  : Fenetre d'analyse repensee. Selecteur de dossier natif
#           Windows (Parcourir), acces rapides vers les dossiers
#           utilisateurs, historique des chemins recents, lecteurs.
#    UI4  : Profondeur en reglette avec explication claire (ce qui est
#           pre-charge, pas un vrai levier de vitesse : les tailles sont
#           toujours exactes ; on descend a la demande en cliquant).
#    HON1 : Le faux "mode de scan" (qui ne changeait rien) est retire.
#           Remplace par un filtre d'affichage honnete : masquer les
#           petits dossiers pour la lisibilite, sans impact sur le scan.
#  NOUVEAUTES v5.10 :
#    PERF3 : Mode web : E4 (propagation) et E5 (re-lecture disque +
#            tri + construction rapport) sautes. L'arbre live tire ses
#            tailles de E3, ces etapes ne servaient plus. Gros gain,
#            surtout la fin de scan (plus de 2e lecture des dossiers).
#    PERF4 : Tableaux chauds (nextLevel, dirsInScope, items) passes en
#            List[object] : fin du cout quadratique sur grosses arbo.
#  NOUVEAUTES v5.9 :
#    UI2  : Refonte ergonomique complete. Layout application plein
#           ecran : barre superieure (fil d'Ariane + actions), ligne
#           de progression fine, arbre pleine largeur, barre de statut
#           en bas (repertoire en cours, etape, chrono). Reglages en
#           modale.
#  NOUVEAUTES v5.8 :
#    SUB2 : Clic sur un gris PENDANT le scan principal : mis en file
#           (marque orange) et scanne automatiquement a la fin du scan.
#           Clic apres le scan : scan immediat. Fusion robuste (la vue
#           suit la racine reellement scannee).
#  NOUVEAUTES v5.7 :
#    SUB1 : Clic sur un dossier gris (non scanne / hors profondeur /
#           protege / exclu) -> scan immediat sur place, 2 niveaux,
#           fusionne dans l'arbre. Bloque si un scan est deja en cours
#           (serveur mono-thread).
#  NOUVEAUTES v5.6 :
#    UX9  : Ligne "En cours" plus lisible : chemin decoupe en segments
#           grises, repertoire en cours de lecture mis en evidence
#           (dernier segment dans une pastille coloree)
#  NOUVEAUTES v5.5 :
#    DOT1 : Point de statut colore sur chaque dossier de l'arbre :
#           vert = scanne, orange = prevu au scan, gris = ni scanne
#           ni prevu (exclu, jonction, protege ACL, hors profondeur)
#    DOT2 : Proteges (ACL) et non-scannes (hors profondeur) affiches
#           dans l'arbre avec point gris
#  NOUVEAUTES v5.4 :
#    MERGE1 : Interface unique "tout-en-un". L'arbre navigable live
#             EST la vue finale : plus de rapport HTML separe ni de
#             bouton "Rapport complet". Scan et lecture au meme endroit.
#    MERGE2 : Barres de proportion + pourcentage par sous-dossier
#             dans l'arbre (vue disque complete)
#  NOUVEAUTES v5.3 :
#    UX7  : Repertoire "En cours" emis 1x/seconde (au lieu de ~6x)
#    UX8  : Ligne "En cours" sur une seule ligne : debut tronque,
#           fin affichee, chemin complet au survol (title)
#    FIX5 : Normalisation des antislashs dans l'arbre live (affichage
#           propre \ -> \, sur-echappement de ConvertTo-JsonSafe contourne)
#  NOUVEAUTES v5.2 :
#    UX6  : Balayage anime sur la barre de progression : montre en
#           permanence l'activite meme quand le pourcentage stagne
#           (gros repertoire en cours de traitement)
#  NOUVEAUTES v5.1 :
#    TRACE1 : Repertoire reellement en cours de lecture emis en direct
#             (evenement 'scanning'), affiche en clair dans l'UI
#    TRACE2 : Dans l'arbre, la branche menant au repertoire en cours
#             de scan est marquee active a chaque niveau de navigation
#  NOUVEAUTES v5.0 :
#    TREE1 : Arborescence navigable qui se construit en direct pendant
#            le scan. Structure emise par E1 au fil de la decouverte,
#            tailles remplies par E3 sous-arbre par sous-arbre.
#    TREE2 : Navigation permanente pendant le scan (descendre/remonter)
#            dans les repertoires deja scannes OU en cours.
#    TREE3 : 3 etats par repertoire :
#            - termine (taille connue, icone dossier + taille)
#            - actif  (en cours de scan, icone animee, navigable)
#            - en attente (structure connue, taille a venir)
#    TREE4 : Bouton "Rapport complet" en fin de scan (vue detaillee)
#    NOTE  : granularite d'activite = sous-arbre de 1er niveau
#            (le moteur calcule un sous-arbre d'un bloc)
#  NOUVEAUTES v4.3 :
#    UX4  : Arborescence live : chaque repertoire de 1er niveau
#           apparait pendant le scan avec une icone "en cours",
#           puis bascule en "termine" + taille une fois scanne
#    UX5  : Moteur instrumente (evenements dir-start / dir-done
#           emis vers le navigateur via SSE), sans impact en mode console
#  NOUVEAUTES v4.2 :
#    UX1  : Au lancement, bascule directe dans la vue d'affichage
#           (fini l'attente sur le formulaire)
#    UX2  : Barre de progression globale + repertoire / etape en cours
#           affiches en direct pendant le scan, avec chrono
#    UX3  : L'arborescence s'affiche des la fin du scan dans la foulee
#  NOUVEAUTES v4.1 :
#    WEB5 : Dossier non scanne -> bouton "Scanner ce dossier"
#           Relance le scan de ce chemin en un clic (au lieu du
#           message invitant a relancer manuellement)
#    WEB6 : Auto-lancement du scan via ?scan=<chemin>&depth=<n>
#           sur la page de reglages (utilise par le bouton ci-dessus)
#    WEB7 : Jeton de session injecte dans le rapport (PSNCDU_TOKEN)
#  NOUVEAUTES v4.0 :
#    WEB1 : Application web locale (System.Net.HttpListener)
#           UI visible des le lancement, avant tout scan
#    WEB2 : Tous les reglages dans l'interface web
#           (chemin, profondeur, mode) - plus de Read-Host
#    WEB3 : Aller-retour natif navigateur <-> PowerShell
#           Scan lance depuis l'UI, progression live via SSE
#    WEB4 : Serveur sur 127.0.0.1 + jeton de session aleatoire
#    BRAND: Retrait branding tiers, passage CL SASU / eric@guiffault.com
#    NOTE : Le serveur exige FullLanguage. En ConstrainedLanguage,
#           le script s'arrete avec un message clair (voir bas du fichier)
#  NOUVEAUTES v3.2 :
#    PERF1 : Enumeration .NET en streaming (FileInfo brut)
#            Remplace Get-ChildItem en methode 1 des 3 enumerateurs
#            Gain typique x10 a x30 sur grosses arborescences
#    PERF2 : PowerShell 7 : recursion native EnumerationOptions
#            (un seul appel, skip ReparsePoint). Repli DFS en 5.1
#    UI1   : Suggestion PowerShell 7 dans le footer si detecte
#            sous Windows PowerShell 5.1 (relance plus rapide)
#  NOUVEAUTES v3.1 :
#    FIX1 : Junctions Windows ignorees (evite double comptage)
#           Mes documents, Menu Démarrer, Local Settings, etc.
#    FIX2 : cmd.exe methode 3 - encodage UTF-8 (chcp 65001)
#           Accents correctement geres (é, è, à, ù, etc.)
#    FIX3 : ConvertTo-JsonSafe - caracteres de controle
#           null byte, tab, backspace, form feed echappes
#    FIX4 : Test-IsJunction - detection ReparsePoint
# ------------------------------------------------------------
#  OPTIMISATIONS v2.9 conservees :
#    OPT1 : Zero hashtable fichier en E3
#    OPT2 : Format-Size uniquement a la serialisation JSON
#    OPT3 : Fusion par niveaux en E1
#  FALLBACK v3.0 conserve :
#    M1 : Get-ChildItem
#    M2 : System.IO.DirectoryInfo
#    M3 : cmd.exe dir (chcp 65001 pour les accents)
# ------------------------------------------------------------
#  Historique :
#    v2.9  - Optimisations perf : -77% temps de scan
#    v3.0  - Fallback 3 niveaux pour droits NTFS partiels
#    v3.1  - Fix junctions + encodage accents + JSON propre
#    v3.2  - Enumeration .NET streaming + suggestion PowerShell 7
#    v4.0  - Application web locale + reglages dans l'UI + de-branding
#    v4.1  - Bouton "Scanner ce dossier" sur les dossiers non scannes
#    v4.2  - Vue d'affichage directe + progression globale et repertoire
#    v4.3  - Arborescence live : repertoires affiches un a un pendant le scan
#    v5.0  - Arbre navigable qui se construit en direct pendant le scan
#    v5.1  - Tracage precis du repertoire en cours de scan
#    v5.2  - Balayage anime sur la barre de progression
#    v5.3  - Emission 1x/s, ligne En cours sur une ligne, antislashs propres
#    v5.4  - Interface unique tout-en-un + barres de proportion
#    v5.5  - Points de statut colores + proteges/non-scannes dans l'arbre
#    v5.6  - Ligne En cours decoupee en segments, repertoire courant mis en avant
#    v5.7  - Clic sur dossier gris = scan a la demande 2 niveaux, fusionne
#    v5.8  - File d'attente des clics gris pendant le scan + fusion robuste
#    v5.9  - Refonte ergonomique : layout application plein ecran
#    v5.10 - Perf : E4/E5 sautes en mode web + tableaux chauds en List
#    v5.11 - Fenetre d'analyse repensee : picker natif, acces rapides, filtre
#    v5.12 - Fichiers listes a la demande + legende depliable
#    v5.13 - File d'attente en bleu + legende a onglets (legende/aide)
#    v5.14 - Fix : lecteurs reseau non interroges (blocage serveur via VPN)
#    v5.15 - Log : vrai nombre de dossiers en fin de scan web
#    v6.0  - Interface multilingue (12 langues, detection OS, repli anglais)
#    v6.1  - Nombre de fichiers par dossier (comptage recursif)
#    v6.2  - Sous-repertoires + fichiers par dossier
#    v6.3  - Version dans l'en-tete + bouton scan groupe des dossiers gris
#    v6.4  - Fix pourcentage > 100 %% + total/compteurs depuis le contenu
#    v6.5  - Avancement en direct sur la ligne du repertoire actif
#    v6.6  - Avancement en direct a tous les niveaux (carte des ancetres)
#    v6.7  - Animation de chargement sur lignes en attente + noms propres
#    v6.8  - Fix repertoires qui disparaissent (taille partielle memorisee)
#    v6.9  - Ordre stable pendant le scan, tri par taille a la fin
#    v6.10 - Tailles a deux decimales
#    v6.11 - Fix antislash quadruples (sur-echappement JSON a la racine)
#    v6.12 - Selecteur de langue dans la fenetre d'analyse
#    v6.13 - Accents (interface et sorties console)
#    v6.14 - Ergonomie fenetre : Echap/croix, Entree, validation, memoire, barre lecteurs
#    v6.15 - Gestion des recents et exclusions editables
#    v6.16 - Fenetre d'analyse reorganisee en deux colonnes, plus compacte
#    v6.17 - Scan illimite entrelace (tailles immediates) + profondeur au pied
#    v6.18 - Interruption d'un scan (annulation cooperative + bouton Interrompre)
#    v6.19 - Selecteur de dossier natif repare (STA correct + fenetre au premier plan)
#    v6.20 - Explorateur de dossiers integre + icones par type de fichier
#    v6.21 - Progression de haut en bas (tri des L1 aligne sur l'affichage)
#    v6.22 - Bouton de tri Nom / Taille (tri par taille lisse pendant le scan)
#    v6.23 - Tri par taille par defaut + 30 langues supplementaires (42 au total)
#    v6.24 - Legende et aide traduites pour les 30 nouvelles langues
#    v6.25 - Libelles de navigation/scan traduits pour les 30 langues (dont dossier parent)
#    v6.26 - 17 cles oubliees traduites dans les 10 langues d'origine (Depth, tri, etc.)
#    v6.27 - Bouton de tri et profondeur retraduits au changement de langue
#    v6.28 - Localisation complete : 89 cles dans les 42 langues
# ============================================================

$SCRIPT_VERSION  = "6.28"
$SCRIPT_DATE     = "2026-09-15"
$USER_EMAIL      = "eric@guiffault.com"
$SCRIPT_AUTHOR   = "Eric Guiffault"
$DEFAULT_PATH    = "C:\Users"
$DEFAULT_DEPTH   = 3
$SERVER_PORT     = 8787   # port de repli si libre, sinon auto (voir bas)

$APP_LOGO_SVG = @'
<svg viewBox="0 0 44 44" xmlns="http://www.w3.org/2000/svg" style="height:40px;width:40px;display:block" aria-label="PS-NCDU">
  <circle cx="22" cy="22" r="18" fill="none" stroke="currentColor" stroke-width="6" opacity="0.18"/>
  <circle cx="22" cy="22" r="18" fill="none" stroke="currentColor" stroke-width="6"
          stroke-dasharray="79 34" stroke-dashoffset="28.3" stroke-linecap="round"
          transform="rotate(-90 22 22)"/>
  <circle cx="22" cy="22" r="6" fill="currentColor"/>
</svg>
'@

$EXCLUDED_DIRS = @(
    "C:\Windows\WinSxS",
    "C:\Windows\Installer",
    "C:\`$Recycle.Bin",
    "C:\System Volume Information",
    "C:\Recovery",
    "C:\ProgramData\Microsoft\Windows Defender",
    "C:\Windows\SoftwareDistribution"
)

$SCAN_MODES = @{
    1 = @{ Name="Equilibre     - Ignore sous-dossiers < 100 MB [RECOMMANDE]"; Speed="~25-40s"; Accuracy="Tres bonne (>95%)" }
    2 = @{ Name="Precis        - Ignore sous-dossiers < 10 MB";               Speed="~40-80s"; Accuracy="Excellente (>99%)" }
    3 = @{ Name="Complet       - Scan integral de tous les fichiers";          Speed="~80-150s";Accuracy="100% exacte"       }
}

$scriptTemp = "$env:TEMP\psncdu"
if (-not (Test-Path $scriptTemp)) { New-Item -ItemType Directory -Path $scriptTemp -Force | Out-Null }
$logFile  = "$scriptTemp\psncdu_debug.log"
$htmlFile = "$scriptTemp\psncdu_report.html"
"============================================================" | Set-Content -Path $logFile -ErrorAction SilentlyContinue

function Write-Log {
    param([string]$Message,[ValidateSet("INFO","WARN","ERROR","DEBUG")][string]$Level="INFO")
    try {
        $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
        Add-Content -Path $logFile -Value "[$ts] [$Level] $Message" -ErrorAction SilentlyContinue
    } catch {}
}
function Write-Step {
    param([string]$Step,[datetime]$StepStart,[string]$Extra="")
    $elapsed=[int]((Get-Date)-$StepStart).TotalSeconds
    $msg="$Step (${elapsed}s)"; if($Extra){$msg+=" - $Extra"}
    Write-Log $msg; Write-Host "  $msg" -ForegroundColor DarkGray
}

$script:LastProgressUpdate = [datetime]::MinValue
$script:ScanGlobalStart    = $null
# v4.0 : quand le scan est declenche par le serveur web, ce sink pousse
# la progression vers le navigateur (SSE). Null = mode console classique.
$script:ProgressSink       = $null
# v4.3 : sink d'evenements repertoire (arborescence live cote navigateur).
# Null = mode console : aucun impact.
$script:DirEventSink       = $null
# v5.0 : canal d'arbre navigable (structure + tailles incrementales)
$script:TreeSink           = $null

function Send-Tree {
    param([string]$EventName,[string]$Json)
    if ($null -ne $script:TreeSink) {
        try { & $script:TreeSink $EventName $Json } catch {}
    }
}

function Send-DirEvent {
    param([string]$Phase,[string]$Path,[string]$Name,[long]$Size = -1,[string]$Kind = 'ok')
    if ($null -ne $script:DirEventSink) {
        try { & $script:DirEventSink $Phase $Path $Name $Size $Kind } catch {}
    }
}

function Update-Progress {
    param([string]$Activity,[string]$Status,[string]$Operation="",[int]$Pct,[bool]$LogIt=$false,[bool]$Force=$false)
    $now = Get-Date
    if (-not $Force -and ($now - $script:LastProgressUpdate).TotalMilliseconds -lt 500) { return }
    $script:LastProgressUpdate = $now
    $elapsed=""
    if ($null -ne $script:ScanGlobalStart) {
        $secs=[int]($now-$script:ScanGlobalStart).TotalSeconds
        if($secs -lt 60){$elapsed=" [${secs}s]"}else{$mins=[int]($secs/60);$secs=$secs%60;$elapsed=" [${mins}m${secs}s]"}
    }
    if ($null -ne $script:ProgressSink) {
        # Mode web : on pousse vers SSE au lieu d'ecrire dans la console
        try { & $script:ProgressSink "$Activity$elapsed" $Status $Operation $Pct } catch {}
    } else {
        Write-Progress -Activity "$Activity$elapsed" -Status $Status -CurrentOperation $Operation -PercentComplete $Pct
    }
    if ($LogIt) { Write-Log "[PROGRESS $Pct%$elapsed] $Status$(if($Operation){" | $Operation"})" }
}

function Get-FullUserName {
    try {
        $wmi = Get-WmiObject -Class Win32_UserAccount -Filter "Name='$($env:USERNAME)'" -ErrorAction SilentlyContinue
        if ($wmi -and -not [string]::IsNullOrWhiteSpace($wmi.FullName)) { return $wmi.FullName }
    } catch {}
    try {
        $adsi = [ADSI]"WinNT://$($env:USERDOMAIN)/$($env:USERNAME),user"
        $fn   = $adsi.FullName
        if (-not [string]::IsNullOrWhiteSpace($fn)) { return $fn.ToString() }
    } catch {}
    return $env:USERNAME
}

Write-Log "PS-NCDU v$SCRIPT_VERSION ($SCRIPT_DATE) Demarrage"
Write-Log "PSVersion    : $($PSVersionTable.PSVersion)"
Write-Log "LanguageMode : $($ExecutionContext.SessionState.LanguageMode)"
Write-Log "Auteur       : $SCRIPT_AUTHOR ($USER_EMAIL)"
Write-Log "User         : $($env:USERNAME)"
Write-Log "Machine      : $($env:COMPUTERNAME)"

# ============================================================
# Utilitaires chemin
# ============================================================
function Normalize-Path {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $Path }
    $Path = $Path.Trim() -replace '/','\' 
    if ($Path -match '^\\\\') {
        $Path = '\\' + ($Path.Substring(2) -replace '\\{2,}','\')
        return $Path.TrimEnd('\')
    }
    $Path = $Path.TrimEnd('\')
    if ($Path -match '^[A-Za-z]:$') { $Path += '\' }
    return $Path
}

function Get-ParentPath {
    param([string]$Path)
    if ([string]::IsNullOrEmpty($Path)) { return $null }
    if ($Path -match '^\\\\') {
        $Path = $Path.TrimEnd('\')
        $parts = $Path.Substring(2) -split '\\'
        if ($parts.Count -le 2) { return $null }
        return '\\' + ($parts[0..($parts.Count-2)] -join '\')
    }
    $Path = ($Path -replace '\\{2,}','\').TrimEnd('\'); $idx = $Path.LastIndexOf('\')
    if ($idx -le 0) { return $null }
    $parent = $Path.Substring(0,$idx)
    if ($parent -match '^[A-Za-z]:$') { $parent += '\' }
    if ($parent -eq $Path) { return $null }
    return $parent
}

function Get-RootDepth {
    param([string]$RootPath)
    if ($RootPath -match '^\\\\') { return 0 }
    $parts = ($RootPath.TrimEnd('\') -split '\\') | Where-Object { $_ -ne '' }
    return ($parts | Measure-Object).Count
}

function Get-PathDepth {
    param([string]$Path,[string]$RootPath,[int]$RootDepth)
    $pathClean=$Path.TrimEnd('\'); $rootClean=$RootPath.TrimEnd('\')
    if ($pathClean -match '^\\\\') {
        if ($pathClean.ToLower() -eq $rootClean.ToLower()) { return 0 }
        $relative=$pathClean.Substring($rootClean.Length).TrimStart('\')
        if ([string]::IsNullOrEmpty($relative)) { return 0 }
        return ($relative -split '\\' | Where-Object { $_ -ne '' } | Measure-Object).Count
    }
    $parts=($pathClean -split '\\') | Where-Object { $_ -ne '' }
    return ($parts | Measure-Object).Count - $RootDepth
}

function Test-IsExcluded {
    param([string]$Path,[string[]]$ExcludedList)
    $pn=$Path.TrimEnd('\').ToLower()
    foreach ($ex in $ExcludedList) {
        $en=$ex.TrimEnd('\').ToLower()
        if ($pn -eq $en -or $pn.StartsWith($en+'\')) { return $true }
    }
    return $false
}

# ✅ v3.1 FIX4 : Detection junction/symlink/reparse point
function Test-IsJunction {
    param([string]$Path)
    try {
        $attr = (Get-Item -Path $Path -Force -ErrorAction Stop).Attributes
        return ($attr -band [System.IO.FileAttributes]::ReparsePoint) -ne 0
    } catch { return $false }
}

function Test-NetworkPath {
    param([string]$Path)
    if ($Path -match '^\\\\') {
        try { return (Test-Path -Path $Path -ErrorAction Stop) }
        catch { Write-Log "[NET] Erreur acces '$Path' : $_" -Level ERROR; return $false }
    }
    return $true
}

function Format-Size {
    param([long]$Size)
    if ($Size -lt 0)       { return "?" }
    if ($Size -ge 1TB)     { return "{0:N2} TB" -f ($Size/1TB) }
    elseif ($Size -ge 1GB) { return "{0:N2} GB" -f ($Size/1GB) }
    elseif ($Size -ge 1MB) { return "{0:N2} MB" -f ($Size/1MB) }
    elseif ($Size -ge 1KB) { return "{0:N2} KB" -f ($Size/1KB) }
    else                   { return "{0} B"     -f $Size }
}

function ConvertTo-HtmlEncoded {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return "" }
    $Text=$Text -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;' -replace "'","&#39;"
    return $Text
}

# ✅ v3.1 FIX3 : JsonSafe - caracteres de controle + accents
function ConvertTo-JsonSafe {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return "" }
    $Text = $Text.Replace('\','\\')
    $Text = $Text -replace '"',  '\"'
    $Text = $Text -replace "`r", ''
    $Text = $Text -replace "`n", ''
    $Text = $Text -replace "`t", '\t'
    # Supprimer caracteres de controle non imprimables (U+0000-U+001F sauf tab)
    $Text = $Text -replace '[\x00-\x08\x0B\x0C\x0E-\x1F]', ''
    return $Text
}

# ============================================================
# ✅ v3.1 FIX2+FIX4 : cmd.exe avec UTF-8 et filtre junctions
# ============================================================
function Invoke-CmdDirSafe {
    param([string]$Path, [string]$DirArgs = "/b /ad")
    $prevEnc = [Console]::OutputEncoding
    try {
        # FIX2 : Forcer UTF-8 pour les accents (é, è, à, etc.)
        [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        $cmdOut = & cmd /c "chcp 65001 >nul 2>nul & dir $DirArgs `"$Path`" 2>nul"
        return $cmdOut
    }
    catch { return $null }
    finally {
        [Console]::OutputEncoding = $prevEnc
    }
}

# ============================================================
# ✅ v3.2 PERF1+PERF2 : Enumeration .NET en streaming
# Renvoie des System.IO.FileInfo bruts (FullName, Name, Length,
# Extension, DirectoryName deja remplis par WIN32_FIND_DATA :
# aucun appel systeme supplementaire par fichier).
#   - PowerShell 7+ : recursion native EnumerationOptions
#     (un seul appel, skip ReparsePoint = pas de double comptage)
#   - Windows PowerShell 5.1 : DFS manuel par pile, skip junctions
# Utilisee comme methode 1 de Get-RecursiveFiles.
# Leve une exception si l'enumeration echoue -> repli sur M2/M3.
# ============================================================
function Get-FilesFastRecurse {
    param([string]$Root)

    if ($PSVersionTable.PSVersion.Major -ge 6) {
        # PS7 : recursion native. AttributesToSkip = ReparsePoint
        # empeche de suivre les junctions (equivalent FIX1).
        $opts = New-Object System.IO.EnumerationOptions
        $opts.RecurseSubdirectories = $true
        $opts.IgnoreInaccessible    = $true
        $opts.AttributesToSkip      = [System.IO.FileAttributes]::ReparsePoint
        $di = New-Object System.IO.DirectoryInfo($Root)
        # @() materialise dans le try : une erreur eventuelle est
        # capturee par l'appelant et declenche le repli M2/M3.
        return @($di.EnumerateFiles('*', $opts))
    }

    # Windows PowerShell 5.1 : DFS manuel, streaming FileInfo brut.
    $result = New-Object System.Collections.ArrayList
    $stack  = New-Object System.Collections.Stack
    [void]$stack.Push($Root)
    while ($stack.Count -gt 0) {
        $dir = $stack.Pop()
        try { $di = New-Object System.IO.DirectoryInfo($dir) } catch { continue }
        try {
            foreach ($f in $di.EnumerateFiles()) { [void]$result.Add($f) }
        } catch {}
        try {
            foreach ($d in $di.EnumerateDirectories()) {
                # FIX1 : ne pas suivre les junctions / reparse points
                if (($d.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -eq 0) {
                    [void]$stack.Push($d.FullName)
                }
            }
        } catch {}
    }
    return $result
}

# ============================================================
# ✅ v3.1 FIX1+FIX4 : Enumeration sous-dossiers
# SANS junctions (evite double comptage)
# AVEC fallback 3 niveaux + encodage UTF-8
# ✅ v3.2 PERF1 : methode 1 via .NET EnumerateDirectories
# ============================================================
function Get-SubDirectories {
    param([string]$Path)

    # ── Methode 1 : .NET EnumerateDirectories (v3.2 PERF1) ───
    try {
        # DirectoryInfo brut : .Attributes et .FullName deja remplis
        $items = @((New-Object System.IO.DirectoryInfo($Path)).EnumerateDirectories())
        # ✅ FIX1 : Filtrer les junctions/symlinks/reparse points
        $jCount = 0
        $filtered = @()
        foreach ($item in $items) {
            if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                $jCount++
                Write-Log "[DIR-M1] Junction ignoree : $($item.FullName)" -Level DEBUG
            } else {
                $filtered += $item
            }
        }
        if ($jCount -gt 0) {
            Write-Log "[DIR-M1] '$Path' : $jCount junction(s) ignoree(s) sur $($items.Count)"
        }
        return @{ Items=$filtered; Method="GCI"; Denied=$false; Junctions=$jCount }
    }
    catch [System.UnauthorizedAccessException] {
        Write-Log "[DIR-M1] .NET refuse '$Path' - essai M2" -Level DEBUG
    }
    catch { Write-Log "[DIR-M1] .NET erreur '$Path' : $_ - essai M2" -Level DEBUG }

    # ── Methode 2 : System.IO.DirectoryInfo ─────────────────
    try {
        $di    = New-Object System.IO.DirectoryInfo($Path)
        $subs  = $di.GetDirectories()
        $jCount = 0
        $items = @()
        foreach ($sub in $subs) {
            # FIX1 : Filtrer les ReparsePoints
            if (($sub.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                $jCount++
                Write-Log "[DIR-M2] Junction ignoree : $($sub.FullName)" -Level DEBUG
            } else {
                $items += New-Object PSObject -Property @{ FullName=$sub.FullName; Name=$sub.Name }
            }
        }
        if ($jCount -gt 0) { Write-Log "[DIR-M2] '$Path' : $jCount junction(s) ignoree(s)" }
        Write-Log "[DIR-M2] DirectoryInfo OK '$Path' ($($items.Count) items)"
        return @{ Items=$items; Method="NET"; Denied=$false; Junctions=$jCount }
    }
    catch [System.UnauthorizedAccessException] {
        Write-Log "[DIR-M2] .NET refuse '$Path' - essai M3" -Level DEBUG
    }
    catch { Write-Log "[DIR-M2] .NET erreur '$Path' : $_ - essai M3" -Level DEBUG }

    # ── Methode 3 : cmd.exe UTF-8 ───────────────────────────
    try {
        # FIX2 : UTF-8 via Invoke-CmdDirSafe
        $cmdOut = Invoke-CmdDirSafe -Path $Path -DirArgs "/b /ad"
        if ($null -ne $cmdOut) {
            $jCount = 0
            $items = @()
            foreach ($line in ($cmdOut | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
                $fp = Join-Path $Path $line
                # FIX1 : Verifier junction
                try {
                    $attr = (Get-Item $fp -Force -ErrorAction Stop).Attributes
                    if (($attr -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                        $jCount++
                        Write-Log "[DIR-M3] Junction ignoree : $fp" -Level DEBUG
                        continue
                    }
                } catch {}
                $items += New-Object PSObject -Property @{ FullName=$fp; Name=$line }
            }
            if ($jCount -gt 0) { Write-Log "[DIR-M3] '$Path' : $jCount junction(s) ignoree(s)" }
            Write-Log "[DIR-M3] cmd.exe UTF-8 OK '$Path' ($($items.Count) items)"
            return @{ Items=$items; Method="CMD"; Denied=$false; Junctions=$jCount }
        }
    }
    catch { Write-Log "[DIR-M3] cmd.exe erreur '$Path' : $_" -Level DEBUG }

    Write-Log "[DIR] PROTEGE (toutes methodes) : '$Path'" -Level INFO
    return @{ Items=@(); Method="NONE"; Denied=$true; Junctions=0 }
}

# ============================================================
# ✅ v3.1 : Enumeration fichiers directs avec fallback + UTF-8
# ============================================================
function Get-DirectFiles {
    param([string]$Path)

    # Methode 1 : .NET EnumerateFiles (v3.2 PERF1)
    try {
        # FileInfo brut : .Length / .Extension deja remplis, zero syscall en plus
        return @((New-Object System.IO.DirectoryInfo($Path)).EnumerateFiles())
    }
    catch [System.UnauthorizedAccessException] {}
    catch { Write-Log "[FILE-M1] Erreur '$Path' : $_" -Level DEBUG }

    # Methode 2
    try {
        $di = New-Object System.IO.DirectoryInfo($Path)
        return $di.GetFiles() | ForEach-Object {
            New-Object PSObject -Property @{
                FullName=$_.FullName; Name=$_.Name
                Length=$_.Length; Extension=$_.Extension; DirectoryName=$Path
            }
        }
    }
    catch [System.UnauthorizedAccessException] {}
    catch { Write-Log "[FILE-M2] Erreur '$Path' : $_" -Level DEBUG }

    # Methode 3 : cmd.exe UTF-8
    try {
        $cmdOut = Invoke-CmdDirSafe -Path $Path -DirArgs "/b /a-d"
        if ($null -ne $cmdOut) {
            return $cmdOut |
                   Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                   ForEach-Object {
                       $fp = Join-Path $Path $_
                       try {
                           $fi = New-Object System.IO.FileInfo($fp)
                           New-Object PSObject -Property @{
                               FullName=$fp; Name=$_
                               Length=$fi.Length; Extension=$fi.Extension; DirectoryName=$Path
                           }
                       } catch { $null }
                   } | Where-Object { $null -ne $_ }
        }
    }
    catch { Write-Log "[FILE-M3] Erreur '$Path' : $_" -Level DEBUG }
    return @()
}

# ============================================================
# ✅ v3.1 : Scan recursif avec fallback + UTF-8
# ============================================================
function Get-RecursiveFiles {
    param([string]$Path)

    # Methode 1 : enumeration .NET streaming (v3.2 PERF1+PERF2)
    # PS7 = recursion native ; 5.1 = DFS manuel. Skip junctions inclus.
    try {
        $fast = Get-FilesFastRecurse -Root $Path
        Write-Log "[RFILE-M1] .NET enum OK '$Path' ($(@($fast).Count) fichiers)" -Level DEBUG
        return $fast
    }
    catch [System.UnauthorizedAccessException] {}
    catch { Write-Log "[RFILE-M1] Erreur '$Path' : $_ - essai M2" -Level DEBUG }

    # Methode 2 : DFS .NET avec filtre junctions
    try {
        $result   = @()
        $dirQueue = @($Path)
        while ($dirQueue.Count -gt 0) {
            $current  = $dirQueue[0]
            $dirQueue = if ($dirQueue.Count -gt 1) { $dirQueue[1..($dirQueue.Count-1)] } else { @() }
            try {
                $di = New-Object System.IO.DirectoryInfo($current)
                foreach ($f in $di.GetFiles()) {
                    $result += New-Object PSObject -Property @{
                        FullName=$f.FullName; Name=$f.Name
                        Length=$f.Length; Extension=$f.Extension; DirectoryName=$current
                    }
                }
                # FIX1 : Ne pas suivre les junctions
                foreach ($s in $di.GetDirectories()) {
                    if (($s.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -eq 0) {
                        $dirQueue += $s.FullName
                    }
                }
            }
            catch {}
        }
        if ($result.Count -gt 0) {
            Write-Log "[RFILE-M2] .NET DFS OK '$Path' ($($result.Count) fichiers)"
            return $result
        }
    }
    catch { Write-Log "[RFILE-M2] Erreur '$Path' : $_" -Level DEBUG }

    # Methode 3 : cmd.exe UTF-8
    try {
        $cmdOut = Invoke-CmdDirSafe -Path $Path -DirArgs "/s /b /a-d"
        if ($null -ne $cmdOut) {
            $result = $cmdOut |
                      Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                      ForEach-Object {
                          try {
                              $fi = New-Object System.IO.FileInfo($_)
                              New-Object PSObject -Property @{
                                  FullName=$_; Name=$fi.Name
                                  Length=$fi.Length; Extension=$fi.Extension; DirectoryName=$fi.DirectoryName
                              }
                          } catch { $null }
                      } | Where-Object { $null -ne $_ }
            Write-Log "[RFILE-M3] cmd.exe UTF-8 OK '$Path' ($($result.Count) fichiers)"
            return $result
        }
    }
    catch { Write-Log "[RFILE-M3] Erreur '$Path' : $_" -Level DEBUG }
    return @()
}

# ============================================================
# v6.17 : enumeration d'un sous-arbre a la demande (mode illimite)
# .NET streaming, saute junctions/exclus, emet les noeuds vers l'arbre
# et enregistre chaque dossier dans les tables de tailles. Appelee dans
# E3 juste avant de mesurer chaque sous-arbre de 1er niveau : les tailles
# s'affichent au fil de l'eau au lieu d'attendre l'enumeration complete.
# ============================================================
function Register-Subtree {
    param([string]$Root,[hashtable]$Own,[hashtable]$Tot,[string]$Act,[string]$L1Name,[int]$PctL1)
    $stack = New-Object System.Collections.Stack
    [void]$stack.Push($Root)
    $cnt = 0
    while ($stack.Count -gt 0) {
        $cur = $stack.Pop()
        try {
            foreach ($sub in ([System.IO.DirectoryInfo]::new($cur)).EnumerateDirectories()) {
                if (($sub.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { continue }
                $fn = $sub.FullName
                if (Test-IsExcluded -Path $fn -ExcludedList $EXCLUDED_DIRS) {
                    Send-Tree 'special' "{""path"":""$(ConvertTo-JsonSafe $fn)"",""parent"":""$(ConvertTo-JsonSafe $cur)"",""name"":""$(ConvertTo-JsonSafe $sub.Name)"",""kind"":""excluded""}"
                    continue
                }
                if (-not $Own.ContainsKey($fn)) { $Own[$fn]=[long]0; $Tot[$fn]=[long]0 }
                Send-Tree 'node' "{""path"":""$(ConvertTo-JsonSafe $fn)"",""parent"":""$(ConvertTo-JsonSafe $cur)"",""name"":""$(ConvertTo-JsonSafe $sub.Name)""}"
                $cnt++
                if (($cnt % 500) -eq 0) { Update-Progress $Act "E3/5 : $L1Name - $cnt dossiers..." "" $PctL1 $false; if ($script:CancelScan) { return $cnt } }
                [void]$stack.Push($fn)
            }
        } catch {}
    }
    return $cnt
}

# ============================================================
# SCAN v3.1
# ============================================================
function Start-FastScan {
    param([string]$RootPath,[int]$MaxDepth=3,[hashtable]$Mode=$null)

    $ACT       = "PS-NCDU v$SCRIPT_VERSION"
    $scanStart = Get-Date
    $script:ScanGlobalStart = $scanStart
    $rootDepth = Get-RootDepth -RootPath $RootPath
    $modeName  = if($null -ne $Mode){$Mode["Name"]}else{"Complet"}
    $isUNC     = $RootPath -match '^\\\\' 
    $pathType  = if($isUNC){"UNC reseau"}else{"Local"}
    $unlimited = ($MaxDepth -eq 0)
    $depthLabel= if($unlimited){"ILLIMITEE"}else{"$MaxDepth niveaux"}

    Write-Log "==========================================";Write-Log "DEBUT SCAN v$SCRIPT_VERSION : $RootPath"
    Write-Log "Type : $pathType | Mode : $modeName | Profondeur : $depthLabel"
    Write-Log "Fix : Junctions ignorees | UTF-8 cmd.exe | JSON propre";Write-Log "=========================================="

    # E1
    $t = Get-Date
    Update-Progress $ACT "E1/5 : Liste des dossiers [$depthLabel]..." "" 1 $true $true
    Send-Tree 'root' "{""path"":""$(ConvertTo-JsonSafe $RootPath)"",""name"":""$(ConvertTo-JsonSafe (Split-Path $RootPath -Leaf))""}"

    $allLevels       = @(@($RootPath))
    $excludedFound   = @()
    $accessDenied    = @()
    $junctionsTotal  = 0
    $methodStats     = @{ GCI=0; NET=0; CMD=0; NONE=0 }

    if ($unlimited) {
        # v6.17 : en illimite, on n'enumere ICI que le premier niveau. Le reste
        # de l'arborescence est enumere sous-arbre par sous-arbre dans E3, juste
        # avant d'en mesurer les tailles (entrelacement). Sans cela, E1 devait
        # parcourir tout le disque avant qu'une seule taille n'apparaisse.
        Write-Log "[E1] Mode illimite : enumeration L1 seule (reste entrelace en E3)"
        Update-Progress $ACT "E1/5 : Dossiers de premier niveau..." "" 10 $true $true
        $flatList = @($RootPath)
        $l1sr = Get-SubDirectories -Path $RootPath
        $junctionsTotal += $l1sr["Junctions"]
        if (-not $l1sr["Denied"]) {
            foreach ($sub in $l1sr["Items"]) {
                if (Test-IsExcluded -Path $sub.FullName -ExcludedList $EXCLUDED_DIRS) {
                    if(-not($excludedFound -contains $sub.FullName)){$excludedFound+=$sub.FullName}
                } else {
                    $flatList += $sub.FullName
                    Send-Tree 'node' "{""path"":""$(ConvertTo-JsonSafe $sub.FullName)"",""parent"":""$(ConvertTo-JsonSafe $RootPath)"",""name"":""$(ConvertTo-JsonSafe $sub.Name)""}"
                }
            }
        } else {
            if(-not($accessDenied -contains $RootPath)){$accessDenied+=$RootPath}
        }
        $allLevels = @($flatList)
        Write-Log "[E1] Mode illimite : $($flatList.Count-1) dossiers L1 | $junctionsTotal junctions ignorees"
    } else {
        for ($d=0; $d -lt $MaxDepth; $d++) {
            if ($script:CancelScan) { break }
            $currentLevel = $allLevels[$d]
            if ($null -eq $currentLevel -or $currentLevel.Count -eq 0) { break }
            $tLevel=$Get=Get-Date; $nextLevel=New-Object System.Collections.Generic.List[object]; $totalDirs=$currentLevel.Count; $dirsDone=0
            $levelNum=$d+1; $levelBase=1+[int](($d*24)/$MaxDepth); $levelRng=[int](24/$MaxDepth)

            foreach ($dir in $currentLevel) {
                $dirsDone++
                $pct=$levelBase+[int](($dirsDone*$levelRng)/($totalDirs+1))
                Update-Progress $ACT "E1/5 : Niveau $levelNum/$MaxDepth | $dirsDone/$totalDirs" "Scan : $dir" $pct ($dirsDone%200 -eq 0)

                $subResult = Get-SubDirectories -Path $dir
                $methodStats[$subResult["Method"]]++
                $junctionsTotal += $subResult["Junctions"]

                if ($subResult["Denied"]) {
                    if(-not($accessDenied -contains $dir)){$accessDenied+=$dir;Write-Log "[E1] PROTEGE : $dir" -Level INFO}
                } else {
                    foreach ($sub in $subResult["Items"]) {
                        if(Test-IsExcluded -Path $sub.FullName -ExcludedList $EXCLUDED_DIRS){
                            if(-not($excludedFound -contains $sub.FullName)){$excludedFound+=$sub.FullName}
                        } else {
                            [void]$nextLevel.Add($sub.FullName)
                            Send-Tree 'node' "{""path"":""$(ConvertTo-JsonSafe $sub.FullName)"",""parent"":""$(ConvertTo-JsonSafe $dir)"",""name"":""$(ConvertTo-JsonSafe $sub.Name)""}"
                        }
                    }
                }
            }
            $allLevels+=,$nextLevel
            Write-Log "[E1] Niveau $levelNum/$MaxDepth : $($nextLevel.Count) en $([int]((Get-Date)-$tLevel).TotalSeconds)s | junctions: $junctionsTotal"
            if($nextLevel.Count -eq 0){break}
        }
    }

    # Fusion niveaux (OPT3) - List pour eviter le O(n^2) sur grosses arbo
    $dirsInScope=New-Object System.Collections.Generic.List[object]
    foreach ($level in $allLevels){ foreach ($dir in $level){ [void]$dirsInScope.Add($dir) } }
    $nbScope=$dirsInScope.Count

    Write-Log "[E1] Stats : GCI=$($methodStats['GCI']) | .NET=$($methodStats['NET']) | CMD=$($methodStats['CMD']) | PROTEGE=$($methodStats['NONE']) | Junctions ignorees=$junctionsTotal"

    # Dossiers non scannes niveau MaxDepth+1
    $dirsUnscanned=@()
    if(-not $unlimited){
        $lastLevel=$allLevels[$allLevels.Count-1]
        if($null -ne $lastLevel){
            foreach ($dir in $lastLevel){
                $depth=Get-PathDepth -Path $dir -RootPath $RootPath -RootDepth $rootDepth
                if($depth -eq $MaxDepth){
                    $subResult=Get-SubDirectories -Path $dir
                    if(-not $subResult["Denied"]){
                        foreach($sub in $subResult["Items"]){
                            if(-not(Test-IsExcluded -Path $sub.FullName -ExcludedList $EXCLUDED_DIRS)){$dirsUnscanned+=$sub.FullName}
                        }
                    } else { if(-not($accessDenied -contains $dir)){$accessDenied+=$dir} }
                }
            }
        }
    }
    # v5.5 : les proteges (ACL) et non-scannes (au-dela profondeur)
    # apparaissent dans l'arbre avec un point gris
    if($null -ne $script:TreeSink){
        foreach($u in $dirsUnscanned){
            $pu=Get-ParentPath $u; $nu=Split-Path $u -Leaf
            Send-Tree 'special' "{""path"":""$(ConvertTo-JsonSafe $u)"",""parent"":""$(ConvertTo-JsonSafe $pu)"",""name"":""$(ConvertTo-JsonSafe $nu)"",""kind"":""unscanned""}"
        }
        foreach($dd in $accessDenied){
            $pdd=Get-ParentPath $dd; $ndd=Split-Path $dd -Leaf
            Send-Tree 'special' "{""path"":""$(ConvertTo-JsonSafe $dd)"",""parent"":""$(ConvertTo-JsonSafe $pdd)"",""name"":""$(ConvertTo-JsonSafe $ndd)"",""kind"":""denied""}"
        }
    }
    Write-Step "[E1] TERMINEE" $t "Scope : $nbScope | NonScannes : $($dirsUnscanned.Count) | Proteges : $($accessDenied.Count) | Junctions ignorees : $junctionsTotal"

    # E2
    $t=Get-Date
    Update-Progress $ACT "E2/5 : Init ($nbScope dossiers)..." "" 25 $true $true
    $dirOwnSizes=@{}; $dirTotalSizes=@{}; $initDone=0
    foreach ($dir in $dirsInScope){
        $initDone++; $dirOwnSizes[$dir]=[long]0; $dirTotalSizes[$dir]=[long]0
        Update-Progress $ACT "E2/5 : Init $initDone / $nbScope" "" (25+[int](($initDone*5)/($nbScope+1))) ($initDone%5000 -eq 0)
    }
    Write-Step "[E2] TERMINEE" $t "$nbScope entrees"

    # E3
    $t=Get-Date; $totalFilesScanned=0; $totalSizeScanned=[long]0
    $liveTotals=@{}   # v5.0 : totaux cumules par dossier pour l'arbre live (mode web)
    $liveCounts=@{}   # v6.1 : nombre de fichiers cumule par dossier (mode web)
    $lastScanEmit=[datetime]::MinValue; $lastScanEmitted=$null   # v5.1 : throttle 'scanning'
    Update-Progress $ACT "E3/5 : Scan tailles..." "" 30 $true $true

    try {
        $rootFiles=Get-DirectFiles -Path $RootPath
        $nbRF=0; $sizeRF=[long]0
        foreach ($f in $rootFiles){if($null -eq $f){continue};$fl=[long]$f.Length;$sizeRF+=$fl;$nbRF++;$dirOwnSizes[$RootPath]+=$fl}
        Write-Log "[E3] Racine : $nbRF fichiers = $(Format-Size $sizeRF)"
        if($null -ne $script:TreeSink){
            # v6.1 : les fichiers directs de la racine comptent aussi dans son total/compte
            $liveTotals[$RootPath]=$sizeRF; $liveCounts[$RootPath]=$nbRF
            Send-Tree 'totals' "{""items"":[{""path"":""$(ConvertTo-JsonSafe $RootPath)"",""size"":$sizeRF,""count"":$nbRF}]}"
        }
    } catch { Write-Log "[E3] Erreur racine : $_" -Level WARN }

    $level1Dirs=@()
    $l1Result=Get-SubDirectories -Path $RootPath
    if(-not $l1Result["Denied"]){ foreach($s in $l1Result["Items"]){$level1Dirs+=$s.FullName} }
    # v6.21 : trier les L1 dans le MEME ordre que l'affichage (ordinal, minuscules,
    # par nom de dossier) pour que la progression se remplisse de haut en bas au
    # lieu d'eparpiller le remplissage dans la liste. Le client trie par
    # nom.toLowerCase() en ordinal ; on reproduit exactement cet ordre ici.
    if ($level1Dirs.Count -gt 1) {
        $l1keys  = [string[]]@($level1Dirs | ForEach-Object { (Split-Path $_ -Leaf).ToLowerInvariant() })
        $l1items = [string[]]$level1Dirs
        [System.Array]::Sort($l1keys, $l1items, [System.StringComparer]::Ordinal)
        $level1Dirs = $l1items
    }
    $totalL1=$level1Dirs.Count; $doneL1=0
    Write-Log "[E3] $totalL1 dossiers L1 | Fallback actif + filtre junctions"

    foreach ($l1dir in $level1Dirs){
        if ($script:CancelScan) { break }
        $doneL1++; $l1Name=Split-Path $l1dir -Leaf
        $pctL1=30+[int](($doneL1*45)/($totalL1+1)); if($pctL1 -gt 74){$pctL1=74}

        if(Test-IsExcluded -Path $l1dir -ExcludedList $EXCLUDED_DIRS){
            Write-Log "[E3] L1 EXCLU : '$l1Name'"
            Update-Progress $ACT "E3/5 : L1 $doneL1/$totalL1 - $l1Name [EXCLU]" "" $pctL1 $true $true
            Send-Tree 'special' "{""path"":""$(ConvertTo-JsonSafe $l1dir)"",""parent"":""$(ConvertTo-JsonSafe $RootPath)"",""name"":""$(ConvertTo-JsonSafe $l1Name)"",""kind"":""excluded""}"
            continue
        }
        # FIX1 : Skip si c est une junction au niveau L1
        if (Test-IsJunction -Path $l1dir) {
            Write-Log "[E3] L1 JUNCTION ignoree : '$l1dir'"
            Update-Progress $ACT "E3/5 : L1 $doneL1/$totalL1 - $l1Name [JUNCTION]" "" $pctL1 $true $true
            Send-Tree 'special' "{""path"":""$(ConvertTo-JsonSafe $l1dir)"",""parent"":""$(ConvertTo-JsonSafe $RootPath)"",""name"":""$(ConvertTo-JsonSafe $l1Name)"",""kind"":""junction""}"
            continue
        }

        Update-Progress $ACT "E3/5 : L1 $doneL1/$totalL1 - $l1Name | $(Format-Size $totalSizeScanned)" "Collecte..." $pctL1 $true $true
        Send-Tree 'active' "{""path"":""$(ConvertTo-JsonSafe $l1dir)""}"
        # v6.17 : en illimite, enumere le sous-arbre de ce L1 juste avant de le
        # mesurer (entrelacement) -> l'arbre et les tailles se remplissent tout
        # de suite, sous-arbre par sous-arbre, au lieu d'attendre tout le disque.
        if ($unlimited) { [void](Register-Subtree -Root $l1dir -Own $dirOwnSizes -Tot $dirTotalSizes -Act $ACT -L1Name $l1Name -PctL1 $pctL1) }
        $touched=@{}   # dossiers de ce sous-arbre dont le total a change

        try {
            $filesInL1=Get-RecursiveFiles -Path $l1dir
            $nbF1=0; $sz1=[long]0; $fdL1=0

            foreach ($file in $filesInL1){
                if($null -eq $file){continue}
                $pd=$null
                try{$pd=$file.DirectoryName}catch{$pd=Get-ParentPath $file.FullName}
                if($null -eq $pd){continue}
                if(Test-IsExcluded -Path $pd -ExcludedList $EXCLUDED_DIRS){continue}

                # v5.1 : trace le repertoire en cours + v6.5 : avancement du L1 actif
                if($null -ne $script:TreeSink){
                    $nowT=[datetime]::Now
                    if(($nowT-$lastScanEmit).TotalMilliseconds -gt 1000){
                        $lastScanEmit=$nowT
                        # v6.6 : partiels de tous les ancetres suivis (avancement a chaque niveau)
                        $ancParts=New-Object System.Collections.Generic.List[object]
                        $aa=$pd; $hh=0
                        while(-not[string]::IsNullOrEmpty($aa) -and $hh -lt 60){
                            if($dirTotalSizes.ContainsKey($aa)){
                                $asz=if($liveTotals.ContainsKey($aa)){$liveTotals[$aa]}else{0}
                                $acn=if($liveCounts.ContainsKey($aa)){$liveCounts[$aa]}else{0}
                                [void]$ancParts.Add("{""path"":""$(ConvertTo-JsonSafe $aa)"",""size"":$asz,""count"":$acn}")
                            }
                            $aa=Get-ParentPath $aa; $hh++
                        }
                        Send-Tree 'scanning' "{""path"":""$(ConvertTo-JsonSafe $pd)"",""anc"":[$($ancParts -join ',')]}"
                    }
                }

                $fl=[long]$file.Length; $sz1+=$fl; $nbF1++; $fdL1++; $totalFilesScanned++; $totalSizeScanned+=$fl

                Update-Progress $ACT "E3/5 : L1 $doneL1/$totalL1 - $l1Name | $(Format-Size $totalSizeScanned)" "$fdL1 fichiers ($(Format-Size $sz1))" $pctL1 ($fdL1%50000 -eq 0)
                if (($fdL1 % 20000) -eq 0 -and $script:CancelScan) { break }

                if($dirOwnSizes.ContainsKey($pd)){
                    $dirOwnSizes[$pd]+=$fl
                } else {
                    $anc=$pd; $hops=0; $found=$false
                    while(-not[string]::IsNullOrEmpty($anc) -and $hops -lt 50){
                        $anc=Get-ParentPath $anc; $hops++
                        if(-not[string]::IsNullOrEmpty($anc) -and $dirOwnSizes.ContainsKey($anc)){
                            $dirOwnSizes[$anc]+=$fl; $found=$true; break
                        }
                    }
                    if(-not $found){$dirOwnSizes[$RootPath]+=$fl}
                }

                # v5.0 (mode web) : cumul du total sur tous les ancetres suivis
                if($null -ne $script:TreeSink){
                    $a=$pd; $h=0
                    while(-not[string]::IsNullOrEmpty($a) -and $h -lt 60){
                        if($dirTotalSizes.ContainsKey($a)){
                            if($liveTotals.ContainsKey($a)){$liveTotals[$a]+=$fl}else{$liveTotals[$a]=$fl}
                            if($liveCounts.ContainsKey($a)){$liveCounts[$a]++}else{$liveCounts[$a]=1}
                            $touched[$a]=$true
                        }
                        $a=Get-ParentPath $a; $h++
                    }
                }
            }
            Write-Log "[E3] L1 $doneL1/$totalL1 '$l1Name' : $nbF1 = $(Format-Size $sz1) | Cumul : $totalFilesScanned = $(Format-Size $totalSizeScanned)"

            if($null -ne $script:TreeSink){
                # Le L1 recoit toujours un total (0 si vide) pour basculer en "termine"
                if(-not $touched.ContainsKey($l1dir)){ if(-not $liveTotals.ContainsKey($l1dir)){$liveTotals[$l1dir]=0}; $touched[$l1dir]=$true }
                $items=New-Object System.Collections.Generic.List[object]
                foreach($tp in $touched.Keys){ [void]$items.Add("{""path"":""$(ConvertTo-JsonSafe $tp)"",""size"":$($liveTotals[$tp]),""count"":$(if($liveCounts.ContainsKey($tp)){$liveCounts[$tp]}else{0})}") }
                Send-Tree 'totals' "{""items"":[$($items -join ',')]}"
            }
        }
        catch {
            Write-Log "[E3] Erreur '$l1dir' : $_" -Level ERROR
            Send-Tree 'totals' "{""items"":[{""path"":""$(ConvertTo-JsonSafe $l1dir)"",""size"":-1}]}"
        }
    }
    Write-Step "[E3] TERMINEE" $t "$totalFilesScanned fichiers = $(Format-Size $totalSizeScanned)"

    # E4 + E5 : uniquement en mode console (rapport statique).
    # En mode web, l'arbre live tire ses tailles de E3 (liveTotals cote client),
    # donc on saute la propagation ET la re-lecture disque de E5 (gros gain vitesse).
    $allScans=@{}; $reportCount=$nbScope; $rootTotalStr=Format-Size $totalSizeScanned
    if($null -eq $script:TreeSink){

    # E4
    $t=Get-Date; Update-Progress $ACT "E4/5 : Propagation..." "" 75 $true $true
    foreach ($dir in $dirsInScope){$dirTotalSizes[$dir]=$dirOwnSizes[$dir]}
    $propDone=0
    foreach ($dir in $dirsInScope){
        $propDone++; $ownSize=$dirOwnSizes[$dir]; if($ownSize -eq 0){continue}
        $pct=75+[int](($propDone*10)/($nbScope+1)); if($pct -gt 84){$pct=84}
        Update-Progress $ACT "E4/5 : $propDone / $nbScope" "Racine : $(Format-Size $dirTotalSizes[$RootPath])" $pct ($propDone%5000 -eq 0)
        $current=Get-ParentPath $dir; $hops=0
        while(-not[string]::IsNullOrEmpty($current) -and $hops -lt 50){
            if($dirTotalSizes.ContainsKey($current)){$dirTotalSizes[$current]+=$ownSize}
            $current=Get-ParentPath $current; $hops++
        }
    }
    Write-Step "[E4] TERMINEE" $t "Racine : $(Format-Size $dirTotalSizes[$RootPath])"

    # E5
    $t=Get-Date; Update-Progress $ACT "E5/5 : Index + rapport..." "" 85 $true $true
    $childIndex=@{}; $idxDone=0
    foreach ($dir in $dirsInScope){
        $idxDone++; $parent=Get-ParentPath $dir
        if($null -ne $parent){if(-not $childIndex.ContainsKey($parent)){$childIndex[$parent]=@()};$childIndex[$parent]+=$dir}
        Update-Progress $ACT "E5/5 : Index $idxDone / $nbScope" "" (85+[int](($idxDone*3)/($nbScope+1))) ($idxDone%5000 -eq 0)
    }
    foreach ($u in $dirsUnscanned){$p=Get-ParentPath $u;if($null -ne $p){if(-not $childIndex.ContainsKey($p)){$childIndex[$p]=@()};if(-not($childIndex[$p] -contains "UNSCANNED:$u")){$childIndex[$p]+="UNSCANNED:$u"}}}
    foreach ($d in $accessDenied){$p=Get-ParentPath $d;if($null -ne $p){if(-not $childIndex.ContainsKey($p)){$childIndex[$p]=@()};if(-not($childIndex[$p] -contains "DENIED:$d")){$childIndex[$p]+="DENIED:$d"}}}
    foreach ($e in $excludedFound){$p=Get-ParentPath $e;if($null -ne $p){if(-not $childIndex.ContainsKey($p)){$childIndex[$p]=@()};if(-not($childIndex[$p] -contains "EXCLU:$e")){$childIndex[$p]+="EXCLU:$e"}}}

    Write-Log "[E5] Index : $($childIndex.Count) parents"
    Update-Progress $ACT "E5/5 : Construction ($nbScope dossiers)..." "" 90 $true $true

    $allScans=@{}; $reportCount=0; $skipCount=0
    foreach ($dirPath in $dirsInScope){
        if(-not $unlimited){
            $dirDepth=Get-PathDepth -Path $dirPath -RootPath $RootPath -RootDepth $rootDepth
            if($dirDepth -gt $MaxDepth){$skipCount++;continue}
        }
        $entries=@()
        if($childIndex.ContainsKey($dirPath)){
            foreach ($ce in $childIndex[$dirPath]){
                if($ce -like "EXCLU:*"){$ep=$ce.Substring(6);$entries+=@{Name=Split-Path $ep -Leaf;FullPath=$ep;Size=[long]-1;IsDir=$true;Excluded=$true;Unscanned=$false;Denied=$false;Ext=""}}
                elseif($ce -like "UNSCANNED:*"){$up=$ce.Substring(10);$entries+=@{Name=Split-Path $up -Leaf;FullPath=$up;Size=[long]-2;IsDir=$true;Excluded=$false;Unscanned=$true;Denied=$false;Ext=""}}
                elseif($ce -like "DENIED:*"){$dp=$ce.Substring(7);$entries+=@{Name=Split-Path $dp -Leaf;FullPath=$dp;Size=[long]-3;IsDir=$true;Excluded=$false;Unscanned=$false;Denied=$true;Ext=""}}
                else{$cn=$ce.Substring($dirPath.TrimEnd('\').Length).TrimStart('\');$entries+=@{Name=$cn;FullPath=$ce;Size=$dirTotalSizes[$ce];IsDir=$true;Excluded=$false;Unscanned=$false;Denied=$false;Ext=""}}
            }
        }
        if($dirOwnSizes[$dirPath] -gt 0){
            try {
                $directFiles=Get-DirectFiles -Path $dirPath
                foreach ($f in $directFiles){
                    if($null -eq $f){continue}
                    $fl=[long]$f.Length
                    $entries+=@{Name=$f.Name;FullPath=$f.FullName;Size=$fl;IsDir=$false;Excluded=$false;Unscanned=$false;Denied=$false;Ext=$f.Extension.ToLower()}
                }
            } catch {}
        }
        $sN=$entries|Where-Object{-not $_["Excluded"] -and -not $_["Unscanned"] -and -not $_["Denied"]}|Sort-Object -Property{$_["Size"]} -Descending
        $sU=$entries|Where-Object{$_["Unscanned"] -eq $true}
        $sD=$entries|Where-Object{$_["Denied"]    -eq $true}
        $sE=$entries|Where-Object{$_["Excluded"]  -eq $true}
        $sorted=@();foreach($e in $sN){$sorted+=$e};foreach($e in $sU){$sorted+=$e};foreach($e in $sD){$sorted+=$e};foreach($e in $sE){$sorted+=$e}
        $allScans[$dirPath]=@{Items=$sorted;Total=$dirTotalSizes[$dirPath]};$reportCount++
        Update-Progress $ACT "E5/5 : $reportCount / $nbScope" "" (90+[int](($reportCount*9)/($nbScope+1))) ($reportCount%2000 -eq 0)
    }
    Write-Step "[E5] TERMINEE" $t "$reportCount construits | $($dirsUnscanned.Count) non scannes | $($accessDenied.Count) proteges | $junctionsTotal junctions ignorees"
    $rootTotalStr=Format-Size $dirTotalSizes[$RootPath]
    }
    else { Write-Log "[E4/E5] Mode web : propagation et rapport statique sautes (arbre live), $nbScope dossiers" }

    Update-Progress $ACT "Termine" "" 100 $true $true; Write-Progress -Activity $ACT -Completed

    $totalElap=[int]((Get-Date)-$scanStart).TotalSeconds
    Write-Log "SCAN TERMINE v$SCRIPT_VERSION en ${totalElap}s | $pathType | $depthLabel | Dossiers : $reportCount | Fichiers : $totalFilesScanned | Proteges : $($accessDenied.Count) | Junctions : $junctionsTotal | Taille : $rootTotalStr"

    return @{
        Scans         = $allScans
        DirCount      = $reportCount
        Excluded      = $excludedFound
        Unscanned     = $dirsUnscanned
        AccessDenied  = $accessDenied
        ModeName      = $modeName
        ElapsedSec    = $totalElap
        MaxDepth      = $MaxDepth
        Unlimited     = $unlimited
        PathType      = $pathType
        MethodStats   = $methodStats
        JunctionsSkipped = $junctionsTotal
    }
}

# ============================================================
# Generation HTML v3.1
# ============================================================
function Get-HtmlReport {
    param(
        [string]$RootPath,
        [hashtable]$AllScans,
        [string[]]$UnscannedDirs,
        [string[]]$AccessDeniedDirs,
        [int]$MaxDepth,
        [bool]$Unlimited,
        [string]$ModeName,
        [string]$ModeAccuracy,
        [int]$ElapsedSec,
        [string]$PathType,
        [string]$FullUserName,
        [string]$ScanDateTime,
        [hashtable]$MethodStats,
        [int]$JunctionsSkipped
    )

    $ACT="PS-NCDU v$SCRIPT_VERSION"; $t=Get-Date
    Write-Log "[HTML] Generation v$SCRIPT_VERSION : $($AllScans.Count) dossiers"
    Update-Progress $ACT "HTML : JSON..." "" 2 $true $true

    $jsonParts=@(); $jsonCount=0
    foreach ($scanPath in $AllScans.Keys){
        $jsonCount++
        Update-Progress $ACT "HTML : JSON $jsonCount / $($AllScans.Count)" "" (2+[int](($jsonCount*80)/($AllScans.Count+1))) ($jsonCount%5000 -eq 0)
        $scanData=$AllScans[$scanPath]; $items=$scanData["Items"]; $total=$scanData["Total"]
        $ip=@()
        foreach ($item in $items){
            $sizeStr=Format-Size $item["Size"]
            # FIX3 : ConvertTo-JsonSafe propre (caracteres de controle)
            $ip+="{"+"""name"":""$(ConvertTo-JsonSafe $item["Name"])"","+"""fullPath"":""$(ConvertTo-JsonSafe $item["FullPath"])"","+"""size"":$($item["Size"]),"+"""sizeStr"":""$(ConvertTo-JsonSafe $sizeStr)"","+"""isDir"":$(if($item["IsDir"]){"true"}else{"false"}),"+"""excluded"":$(if($item["Excluded"]){"true"}else{"false"}),"+"""unscanned"":$(if($item["Unscanned"]){"true"}else{"false"}),"+"""denied"":$(if($item["Denied"]){"true"}else{"false"}),"+"""ext"":""$(ConvertTo-JsonSafe $item["Ext"])"","+"""type"":""$(if($item["IsDir"]){"DIR"}else{"FILE"})"""+"}"
        }
        $jsonParts+="""$(ConvertTo-JsonSafe $scanPath)"":{""items"":["+($ip -join ",")+"],""total"":$total}"
    }
    $jsonScans="{"+($jsonParts -join ",")+"}"
    $exParts=@(); foreach($ex in $EXCLUDED_DIRS){$exParts+="""$(ConvertTo-JsonSafe $ex)"""}
    $jsonExcluded="["+($exParts -join ",")+"]"

    Write-Step "[HTML] JSON serialise" $t "$jsonCount dossiers"
    Update-Progress $ACT "HTML : Generation page..." "" 85 $true $true; $t=Get-Date

    $rootPathSafe  = ConvertTo-JsonSafe $RootPath
    $psVer         = $PSVersionTable.PSVersion.ToString()
    $logEnc        = ConvertTo-HtmlEncoded $logFile
    $rootEnc       = ConvertTo-HtmlEncoded $RootPath
    $nbScans       = $AllScans.Count
    $emailEnc      = ConvertTo-HtmlEncoded $USER_EMAIL
    $authorEnc     = ConvertTo-HtmlEncoded $SCRIPT_AUTHOR
    $modeEncHtml   = ConvertTo-HtmlEncoded $ModeName
    $modeEncAcc    = ConvertTo-HtmlEncoded $ModeAccuracy
    $logoSvg       = $APP_LOGO_SVG
    $maxDepthJs    = $MaxDepth
    $unlimitedJs   = if($Unlimited){"true"}else{"false"}
    $nbUnscanned   = $UnscannedDirs.Count
    $nbDenied      = $AccessDeniedDirs.Count
    $pathTypeEnc   = ConvertTo-HtmlEncoded $PathType
    $pathIcon      = if($RootPath -match '^\\\\'){"&#127760;"}else{"&#128190;"}
    $depthLabel    = if($Unlimited){"Illimitee"}else{"$MaxDepth niveaux"}
    $depthLabelEnc = ConvertTo-HtmlEncoded $depthLabel
    $fullUserEnc   = ConvertTo-HtmlEncoded $FullUserName
    $scanDTEnc     = ConvertTo-HtmlEncoded $ScanDateTime
    $machineEnc    = ConvertTo-HtmlEncoded $env:COMPUTERNAME
    # Info junctions et fallback dans footer
    $footerExtra = ""
    if ($JunctionsSkipped -gt 0) { $footerExtra += " &middot; $JunctionsSkipped junctions ignorees" }
    if ($null -ne $MethodStats -and ($MethodStats["NET"] -gt 0 -or $MethodStats["CMD"] -gt 0)) {
        $footerExtra += " &middot; Fallback .NET=$($MethodStats['NET']) CMD=$($MethodStats['CMD'])"
    }

    # ✅ v3.2 UI1 : moteur utilise + suggestion PowerShell 7 si on tourne en 5.1
    $engineLabelEnc = ConvertTo-HtmlEncoded "PowerShell $psVer"
    $engineHintHtml = ""
    if ($PSVersionTable.PSVersion.Major -lt 6 -and -not [string]::IsNullOrWhiteSpace($PSCommandPath)) {
        $pwshPath = $null
        $cmd = Get-Command pwsh.exe -ErrorAction SilentlyContinue
        if ($cmd) {
            $pwshPath = $cmd.Source
        } else {
            foreach ($p in @(
                "$env:ProgramFiles\PowerShell\7\pwsh.exe",
                "${env:ProgramFiles(x86)}\PowerShell\7\pwsh.exe"
            )) { if (Test-Path $p) { $pwshPath = $p; break } }
        }
        if ($pwshPath) {
            Write-Log "[HTML] PowerShell 7 detecte : $pwshPath - suggestion footer affichee"
            $relaunch     = "pwsh -NoProfile -File `"$PSCommandPath`""
            $relaunchHtml = ConvertTo-HtmlEncoded $relaunch
            $engineHintHtml = "<div class=""engine-hint"">&#9889; <strong>PowerShell 7 detecte</strong> sur ce poste. Le scan actuel tourne sous Windows PowerShell 5.1. Pour un scan plus rapide, relancez avec :<br><code>$relaunchHtml</code></div>"
        }
    }
    $exListHtml="";foreach($ex in $EXCLUDED_DIRS){$exListHtml+="<li><code>$(ConvertTo-HtmlEncoded $ex)</code></li>"}
    $dnListHtml="";foreach($dn in $AccessDeniedDirs){$dnListHtml+="<li><code>$(ConvertTo-HtmlEncoded $dn)</code></li>"}

    $html = @"
<!DOCTYPE html>
<html lang="fr" data-theme="light">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width,initial-scale=1.0">
    <title>PS-NCDU v$SCRIPT_VERSION - $rootEnc</title>
    <style>
    @font-face{font-family:'Public Sans';font-style:normal;font-weight:400;src:url(https://fonts.gstatic.com/s/publicsans/v21/ijwRs572Xtc6ZYQws9YVwllKVnqpTiU.woff2) format('woff2')}
    @font-face{font-family:'Public Sans';font-style:italic;font-weight:400;src:url(https://fonts.gstatic.com/s/publicsans/v21/ijwTs572Xtc6ZYQws9YVwnNDTJLax9k0.woff2) format('woff2')}
    @font-face{font-family:'Public Sans';font-style:normal;font-weight:600;src:url(https://fonts.gstatic.com/s/publicsans/v21/ijwRs572Xtc6ZYQws9YVwllKVnqpTiU.woff2) format('woff2')}
    @font-face{font-family:'Public Sans';font-style:normal;font-weight:700;src:url(https://fonts.gstatic.com/s/publicsans/v21/ijwRs572Xtc6ZYQws9YVwllKVnqpTiU.woff2) format('woff2')}
    :root{--bg:#f0f4f8;--surface:#ffffff;--card:#e8edf3;--card-hover:#dde4ed;--border:#cdd6e0;--border-light:#dde4ed;--text:#0f1923;--text-muted:#4a6080;--text-dim:#8da0b8;--accent:#0063be;--accent-light:#004fa3;--accent-glow:rgba(0,99,190,.08);--success:#007a5e;--warning:#b86b00;--danger:#c0202e;--unscanned:#5e3d8f;--shadow:0 2px 8px rgba(0,0,0,.1);--shadow-sm:0 1px 3px rgba(0,0,0,.08);--radius:8px;--radius-sm:5px;--radius-pill:20px;--logo-color:#001965}
    [data-theme="dark"]{--bg:#0f1923;--surface:#162032;--card:#1e2d42;--card-hover:#243350;--border:#2a3f5f;--border-light:#1e3050;--text:#D3B276;--text-muted:#a08a5a;--text-dim:#6b5a3a;--accent:#0063be;--accent-light:#D3B276;--accent-glow:rgba(211,178,118,.15);--success:#00c48c;--warning:#f5a623;--danger:#e8394a;--unscanned:#9b7fd4;--shadow:0 2px 12px rgba(0,0,0,.45);--shadow-sm:0 1px 4px rgba(0,0,0,.3);--logo-color:#D3B276}
    *{box-sizing:border-box;margin:0;padding:0}
    body{font-family:'Public Sans','Segoe UI',Roboto,Arial,sans-serif;background:var(--bg);color:var(--text);min-height:100vh;font-size:14px;transition:background .3s,color .3s}
    .header{background:var(--surface);height:68px;padding:0 24px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;box-shadow:var(--shadow);position:sticky;top:0;z-index:200;gap:16px}
    .header-left{display:flex;align-items:center;gap:12px;flex-shrink:0}
    .app-title-main{font-size:1.3em;font-weight:700;letter-spacing:.5px;color:var(--text)}
    .app-title-main span{color:var(--accent-light)}
    .app-title-sub{font-size:.72em;color:var(--text-muted);letter-spacing:.3px}
    .version-chip{font-size:.7em;background:var(--accent-glow);color:var(--accent-light);border:1px solid var(--accent-light);padding:2px 10px;border-radius:var(--radius-pill);font-weight:600;white-space:nowrap}
    .header-right{display:flex;align-items:center;gap:14px;flex-shrink:0}
    .user-block{display:flex;align-items:center;gap:10px;background:var(--card);border:1px solid var(--border);border-radius:var(--radius);padding:6px 14px}
    .user-avatar{width:34px;height:34px;border-radius:50%;background:var(--accent-light);color:var(--surface);display:flex;align-items:center;justify-content:center;font-weight:700;font-size:.95em;flex-shrink:0}
    [data-theme="dark"] .user-avatar{color:#0f1923}
    .user-info{display:flex;flex-direction:column;gap:1px;text-align:left}
    .user-fullname{font-size:.85em;font-weight:600;color:var(--text);white-space:nowrap;max-width:200px;overflow:hidden;text-overflow:ellipsis}
    .user-details{font-size:.72em;color:var(--text-muted);white-space:nowrap}
    .theme-toggle{background:var(--card);border:1px solid var(--border);color:var(--text-muted);padding:6px 12px;border-radius:var(--radius-sm);cursor:pointer;font-size:.85em;transition:all .2s;white-space:nowrap;font-family:inherit;flex-shrink:0}
    .theme-toggle:hover{background:var(--card-hover);color:var(--text);border-color:var(--accent-light)}
    .nn-logo{cursor:pointer;display:flex;align-items:center;color:var(--logo-color);height:44px;transition:opacity .2s;flex-shrink:0}
    .nn-logo:hover{opacity:.75}
    .nn-logo svg{height:44px;width:auto;color:var(--logo-color)}
    .scan-datebar{background:var(--accent-light);color:var(--surface);padding:4px 24px;font-size:.78em;font-weight:500;display:flex;align-items:center;gap:16px;flex-wrap:wrap;border-bottom:1px solid var(--border)}
    [data-theme="dark"] .scan-datebar{background:var(--card);color:var(--text-muted)}
    .scan-datebar span{opacity:.9}.scan-datebar strong{opacity:1;font-weight:700}
    .toolbar{background:var(--surface);padding:10px 24px;border-bottom:1px solid var(--border);display:flex;gap:10px;align-items:center;flex-wrap:wrap}
    .breadcrumb{display:flex;flex-wrap:wrap;gap:2px;align-items:center;font-size:.85em;flex:1;min-width:0}
    .breadcrumb a{color:var(--accent-light);text-decoration:none;cursor:pointer;padding:4px 8px;border-radius:var(--radius-sm);transition:background .15s;white-space:nowrap}
    .breadcrumb a:hover{background:var(--accent-glow)}
    .breadcrumb a.active{color:var(--text);font-weight:600;background:var(--card);cursor:default}
    .sep{color:var(--text-dim);padding:0 2px;user-select:none}
    .nav-input{background:var(--card);border:1px solid var(--border);color:var(--text);padding:7px 14px;border-radius:var(--radius-pill);font-size:.85em;font-family:inherit;width:280px;outline:none;transition:all .2s}
    .nav-input:focus{border-color:var(--accent-light);box-shadow:0 0 0 3px var(--accent-glow);width:380px}
    .nav-input::placeholder{color:var(--text-dim)}
    .btn{background:var(--card);color:var(--accent-light);border:1px solid var(--border);padding:7px 16px;border-radius:var(--radius-sm);cursor:pointer;font-size:.875em;font-weight:500;transition:all .15s;white-space:nowrap;font-family:inherit}
    .btn:hover{background:var(--card-hover);border-color:var(--accent-light);box-shadow:var(--shadow-sm)}
    .btn-primary{background:var(--accent);color:#fff;border-color:var(--accent)}
    .btn-primary:hover{background:#0050a0}
    .mode-banner{background:var(--card);padding:8px 24px;font-size:.8em;display:flex;align-items:center;gap:10px;flex-wrap:wrap;border-bottom:1px solid var(--border);color:var(--text-muted)}
    .mode-tag{display:inline-flex;align-items:center;background:var(--accent-glow);color:var(--accent-light);border:1px solid var(--accent-light);padding:2px 10px;border-radius:var(--radius-pill);font-weight:600;font-size:.9em;white-space:nowrap}
    .path-type-badge{display:inline-flex;align-items:center;gap:4px;background:rgba(0,99,190,.1);color:var(--accent);border:1px solid rgba(0,99,190,.3);padding:2px 10px;border-radius:var(--radius-pill);font-weight:600;font-size:.85em;white-space:nowrap}
    .depth-badge{display:inline-flex;align-items:center;gap:4px;padding:2px 10px;border-radius:var(--radius-pill);font-weight:600;font-size:.85em;white-space:nowrap}
    .depth-unlimited{background:rgba(0,122,94,.1);color:var(--success);border:1px solid var(--success)}
    .depth-limited{background:var(--accent-glow);color:var(--accent-light);border:1px solid var(--accent-light)}
    .alert-banner{padding:10px 24px;font-size:.85em;display:flex;align-items:flex-start;gap:10px;border-bottom:1px solid var(--border);flex-wrap:wrap}
    .alert-unscanned{background:rgba(94,61,143,.08);border-left:4px solid var(--unscanned);color:var(--unscanned)}
    .alert-denied{background:rgba(192,32,46,.07);border-left:4px solid var(--danger);color:var(--danger)}
    .alert-banner strong{font-weight:700}
    .hint-link{color:var(--accent-light);cursor:pointer;text-decoration:underline;font-size:.9em;margin-left:6px}
    .alert-detail{display:none;padding:6px 24px 10px 40px;background:var(--card);border-bottom:1px solid var(--border);font-size:.82em;color:var(--text-muted)}
    .alert-detail ul{margin:0;padding:0;list-style:none}
    .alert-detail li{padding:2px 0}
    .alert-detail code{padding:1px 6px;border-radius:3px;background:rgba(192,32,46,.08);color:var(--danger)}
    .excl-section{border-bottom:1px solid var(--border)}
    .excl-toggle{width:100%;background:none;border:none;cursor:pointer;padding:9px 24px;display:flex;align-items:center;gap:8px;font-size:.82em;color:var(--warning);text-align:left;font-family:inherit;transition:background .15s}
    .excl-toggle:hover{background:var(--card)}
    .excl-count{background:rgba(176,107,0,.12);color:var(--warning);border:1px solid var(--warning);padding:1px 8px;border-radius:var(--radius-pill);font-weight:600;font-size:.85em}
    .excl-chevron{margin-left:auto;transition:transform .25s;color:var(--text-dim);font-size:.8em}
    .excl-toggle.open .excl-chevron{transform:rotate(180deg)}
    .excl-body{display:none;padding:6px 24px 10px 40px;background:var(--card);border-top:1px solid var(--border-light)}
    .excl-body.open{display:block}
    .excl-body ul{margin:0;padding:0;list-style:none}
    .excl-body li{padding:2px 0;font-size:.82em;color:var(--text-muted)}
    .excl-body code{background:rgba(176,107,0,.08);color:var(--warning);padding:1px 6px;border-radius:3px}
    .statsbar{background:var(--surface);padding:10px 24px;border-bottom:1px solid var(--border);display:flex;gap:8px;flex-wrap:wrap;align-items:center}
    .chip{display:inline-flex;align-items:center;gap:6px;background:var(--card);border:1px solid var(--border);border-radius:var(--radius-pill);padding:5px 14px;font-size:.8em}
    .chip-label{color:var(--text-muted)}.chip-value{color:var(--accent-light);font-weight:600}
    .chip-warn{border-color:var(--danger)}.chip-warn .chip-value{color:var(--danger)}
    .chip-excl{border-color:var(--warning)}.chip-excl .chip-value{color:var(--warning)}
    .chip-unscanned{border-color:var(--unscanned)}.chip-unscanned .chip-value{color:var(--unscanned)}
    .chip-denied{border-color:var(--danger)}.chip-denied .chip-value{color:var(--danger)}
    .table-wrap{padding:16px 24px;overflow-x:auto}
    table{width:100%;border-collapse:collapse;font-size:.875em;background:var(--surface);border-radius:var(--radius);overflow:hidden;box-shadow:var(--shadow)}
    thead th{background:var(--card);color:var(--text-muted);padding:12px 16px;text-align:left;border-bottom:2px solid var(--accent-light);font-weight:600;font-size:.75em;letter-spacing:.5px;text-transform:uppercase;cursor:pointer;user-select:none;white-space:nowrap;transition:background .15s,color .15s}
    thead th:hover{background:var(--card-hover);color:var(--accent-light)}
    thead th.sort-asc::after{content:' \u25b2';color:var(--accent-light)}
    thead th.sort-desc::after{content:' \u25bc';color:var(--accent-light)}
    tbody tr{border-bottom:1px solid var(--border-light);transition:background .1s}
    tbody tr:hover{background:var(--card)}
    tbody tr:last-child{border-bottom:none}
    tbody tr.row-excluded{background:rgba(176,107,0,.05)}tbody tr.row-excluded:hover{background:rgba(176,107,0,.1)}
    tbody tr.row-unscanned{background:rgba(94,61,143,.05)}tbody tr.row-unscanned:hover{background:rgba(94,61,143,.1)}
    tbody tr.row-denied{background:rgba(192,32,46,.05)}tbody tr.row-denied:hover{background:rgba(192,32,46,.1)}
    td{padding:10px 16px;vertical-align:middle}
    .col-icon{width:28px;text-align:center}.col-size{width:120px;text-align:right;font-family:monospace;font-size:.85em;color:var(--text-muted)}
    .col-bar{width:220px}.col-type{width:70px;text-align:center}
    .dir-link{color:var(--accent-light);text-decoration:none;font-weight:500;cursor:pointer;transition:color .15s}
    .dir-link:hover{color:var(--text);text-decoration:underline}
    .dir-link-unscanned{color:var(--unscanned);text-decoration:none;font-weight:500;cursor:pointer}
    .dir-link-unscanned:hover{color:var(--text);text-decoration:underline}
    .dir-link-denied{color:var(--danger);text-decoration:none;font-weight:500;cursor:default}
    .file-name{color:var(--text)}
    .type-badge{font-size:.7em;padding:2px 7px;border-radius:3px;font-weight:600}
    .type-dir{background:rgba(0,99,190,.1);color:var(--accent-light)}
    .type-file{background:var(--card);color:var(--text-dim)}
    .type-excl{background:rgba(176,107,0,.12);color:var(--warning)}
    .type-unscanned{background:rgba(94,61,143,.12);color:var(--unscanned)}
    .type-denied{background:rgba(192,32,46,.12);color:var(--danger)}
    .excl-label{font-size:.7em;background:rgba(176,107,0,.12);color:var(--warning);padding:2px 7px;border-radius:3px;margin-left:6px;font-weight:500}
    .denied-label{font-size:.7em;background:rgba(192,32,46,.12);color:var(--danger);padding:2px 7px;border-radius:3px;margin-left:6px;font-weight:500}
    .size-unknown{color:var(--warning);font-weight:600;font-style:italic}
    .size-unscanned{color:var(--unscanned);font-weight:600;font-style:italic}
    .size-denied{color:var(--danger);font-weight:600;font-style:italic}
    .bar-wrap{display:flex;align-items:center;gap:8px}
    .bar-bg{background:var(--card);border-radius:4px;height:8px;flex:1;overflow:hidden;border:1px solid var(--border-light)}
    .bar-fill{height:100%;border-radius:4px;transition:width .3s}
    .bar-pct{font-size:.75em;color:var(--text-dim);white-space:nowrap;min-width:38px;text-align:right}
    .bar-unknown{font-size:.75em;color:var(--warning);font-weight:600}
    .bar-unscanned-text{font-size:.75em;color:var(--unscanned);font-weight:600;font-style:italic}
    .bar-denied-text{font-size:.75em;color:var(--danger);font-weight:600}
    .bar-excl{height:8px;background:repeating-linear-gradient(45deg,rgba(176,107,0,.1),rgba(176,107,0,.1) 4px,rgba(176,107,0,.2) 4px,rgba(176,107,0,.2) 8px)}
    .bar-unscanned{height:8px;background:repeating-linear-gradient(45deg,rgba(94,61,143,.1),rgba(94,61,143,.1) 4px,rgba(94,61,143,.2) 4px,rgba(94,61,143,.2) 8px)}
    .bar-denied{height:8px;background:repeating-linear-gradient(45deg,rgba(192,32,46,.1),rgba(192,32,46,.1) 4px,rgba(192,32,46,.2) 4px,rgba(192,32,46,.2) 8px)}
    .msg-page{padding:56px 24px;text-align:center}
    .msg-page .icon{font-size:3em;margin-bottom:16px;display:block}
    .msg-page h3{color:var(--text);margin-bottom:8px;font-weight:500;font-size:1.1em}
    .msg-page p{color:var(--text-muted);font-size:.875em;margin-top:6px}
    .msg-page .hint{margin-top:14px;font-size:.8em;background:rgba(176,107,0,.1);color:var(--warning);border:1px solid rgba(176,107,0,.3);padding:8px 16px;border-radius:var(--radius-sm);display:inline-block}
    .msg-page .hint-unscanned{margin-top:14px;font-size:.8em;background:rgba(94,61,143,.1);color:var(--unscanned);border:1px solid rgba(94,61,143,.3);padding:8px 16px;border-radius:var(--radius-sm);display:inline-block}
    .msg-page .hint-denied{margin-top:14px;font-size:.8em;background:rgba(192,32,46,.1);color:var(--danger);border:1px solid rgba(192,32,46,.3);padding:8px 16px;border-radius:var(--radius-sm);display:inline-block}
    .msg-page code{font-family:monospace;font-size:.85em;color:var(--accent-light);background:var(--card);padding:2px 8px;border-radius:3px;margin-top:8px;display:inline-block}
    .msg-page .rescan-btn{margin-top:16px;background:var(--accent);color:#fff;border:none;padding:11px 22px;border-radius:var(--radius-sm);cursor:pointer;font-size:.95em;font-weight:600;font-family:inherit;transition:background .15s}
    .msg-page .rescan-btn:hover{background:var(--accent-light)}
    .msg-page ul.protected-list{text-align:left;margin:12px auto;display:inline-block;color:var(--text-muted);font-size:.85em;list-style:disc;padding-left:20px}
    .msg-page ul.protected-list li{margin:4px 0}
    .footer{text-align:center;padding:14px 24px;color:var(--text-dim);font-size:.78em;border-top:1px solid var(--border);background:var(--surface);margin-top:16px;line-height:1.8}
    .footer a{color:var(--accent-light);text-decoration:none}
    .footer a:hover{text-decoration:underline}
    .footer-support{margin-top:6px;font-size:.85em;color:var(--text-muted)}
    .engine-hint{margin:12px auto 0;max-width:680px;text-align:left;background:var(--accent-glow);border:1px solid var(--accent-light);border-radius:var(--radius-sm);padding:9px 14px;font-size:.92em;color:var(--text-muted);line-height:1.6}
    .engine-hint strong{color:var(--accent-light)}
    .engine-hint code{display:inline-block;margin-top:5px;background:var(--card);border:1px solid var(--border);border-radius:3px;padding:3px 9px;color:var(--text);font-family:monospace;font-size:.95em;user-select:all}
    ::-webkit-scrollbar{width:6px;height:6px}
    ::-webkit-scrollbar-track{background:var(--bg)}
    ::-webkit-scrollbar-thumb{background:var(--border);border-radius:3px}
    ::-webkit-scrollbar-thumb:hover{background:var(--accent-light)}
    @media(max-width:768px){.col-bar,.col-type{display:none}.nav-input{width:140px}.nav-input:focus{width:220px}.user-block,.user-info{display:none}.scan-datebar{font-size:.72em;padding:4px 12px}}
    </style>
</head>
<body>

<div class="header">
    <div class="header-left">
        <div>
            <div class="app-title-main">PS-<span>NCDU</span></div>
            <div class="app-title-sub">Disk Usage Analyzer</div>
        </div>
        <span class="version-chip">v$SCRIPT_VERSION</span>
    </div>
    <div class="header-right">
        <button class="theme-toggle" onclick="toggleTheme()" id="themeBtn">&#9728; Light</button>
        <div class="user-block">
            <div class="user-avatar" id="userAvatar">?</div>
            <div class="user-info">
                <div class="user-fullname" title="$fullUserEnc">$fullUserEnc</div>
                <div class="user-details">$machineEnc &middot; $($env:USERNAME)</div>
            </div>
        </div>
        <div class="nn-logo" onclick="navigateTo(ROOT_PATH)" title="Retour racine">$logoSvg</div>
    </div>
</div>

<div class="scan-datebar">
    <span>&#128197; Scan du <strong>$scanDTEnc</strong></span>
    <span>&middot;</span>
    <span>$pathIcon <strong>$rootEnc</strong></span>
    <span>&middot;</span>
    <span>&#8987; ${ElapsedSec}s &middot; $nbScans dossiers</span>
    $(if ($JunctionsSkipped -gt 0) { "<span>&middot; &#128279; $JunctionsSkipped jonctions ignorees</span>" })
</div>

<div class="toolbar">
    <div class="breadcrumb" id="breadcrumb"></div>
    <input type="text" class="nav-input" id="pathInput" value="$rootEnc"
           placeholder="C:\chemin ou \\serveur\partage"
           onkeydown="if(event.key==='Enter')navigateTo(this.value)">
    <button class="btn btn-primary" onclick="navigateTo(document.getElementById('pathInput').value)" title="Naviguer">&#128269;</button>
    <button class="btn" onclick="goUp()">&#11014; Remonter</button>
    <button class="btn" onclick="navigateTo(ROOT_PATH)" title="Racine">&#127968;</button>
</div>

<div class="mode-banner">
    <span>&#8505;&#65039;</span>
    <span class="mode-tag">$modeEncHtml</span>
    <span class="path-type-badge">$pathIcon $pathTypeEnc</span>
    <span class="depth-badge $(if($Unlimited){'depth-unlimited'}else{'depth-limited'})">$(if($Unlimited){'&#8734; Profondeur illimitee'}else{"&#128269; Profondeur : $MaxDepth"})</span>
    <span>Precision : <strong style="color:var(--text)">$modeEncAcc</strong></span>
</div>

$(if ($nbDenied -gt 0) {
"<div class='alert-banner alert-denied'>
    <span>&#128274;</span>
    <div>
        <strong>$nbDenied repertoire(s) proteges par ACL NTFS</strong>
        - Inaccessibles avec le compte <strong>$($env:USERNAME)</strong>.
        C'est normal pour les dossiers systeme ou d'autres utilisateurs.
        <span class='hint-link' onclick='toggleDetail(""deniedDetail"")'>&#128196; Voir la liste</span>
    </div>
</div>
<div class='alert-detail' id='deniedDetail'><ul>$dnListHtml</ul></div>"
})

$(if ($nbUnscanned -gt 0) {
"<div class='alert-banner alert-unscanned'>
    <span>&#128269;</span>
    <div><strong>$nbUnscanned repertoire(s) non scanne(s)</strong>
    - Scan arrete a la profondeur <strong>$MaxDepth</strong>. Taille inconnue.
    <span class='hint-link' onclick='toggleDetail(""deepScanHint"")'>&#128161; Scanner plus profond</span></div>
</div>
<div class='alert-detail' id='deepScanHint'>Relancez avec profondeur <strong>0 (illimitee)</strong> ou augmentez le niveau.</div>"
})

$(if ($EXCLUDED_DIRS.Count -gt 0) {
"<div class='excl-section'>
    <button class='excl-toggle' id='exclToggle' onclick='toggleExcl()'>
        <span>&#9888;&#65039;</span>
        <span>Repertoires exclus du scan - taille inconnue</span>
        <span class='excl-count'>$($EXCLUDED_DIRS.Count)</span>
        <span class='excl-chevron'>&#9660;</span>
    </button>
    <div class='excl-body' id='exclBody'><ul>$exListHtml</ul></div>
</div>"
})

<div class="statsbar" id="stats"></div>
<div id="mainContent"></div>

<div class="footer">
    PS-NCDU v$SCRIPT_VERSION &middot; Disk Usage Analyzer &middot;
    $pathIcon $pathTypeEnc &middot; $depthLabelEnc &middot;
    $nbScans dossiers &middot; $scanDTEnc &middot; $authorEnc$footerExtra
    <div class="footer-support">
        Support : <a href="mailto:$emailEnc">$emailEnc</a>
        &nbsp;&middot;&nbsp; Moteur : $engineLabelEnc
        &nbsp;&middot;&nbsp; Log : <code>$logEnc</code>
    </div>
    $engineHintHtml
</div>

<script>
var DATA=($jsonScans),EXCLUDED=($jsonExcluded),ROOT_PATH="$rootPathSafe";
var MAX_DEPTH=$maxDepthJs,UNLIMITED=$unlimitedJs;
var curPath=ROOT_PATH,sortCol='size',sortAsc=false;

(function(){
    var fn="$fullUserEnc",parts=fn.split(' ').filter(function(p){return p.length>0;}),ini='';
    if(parts.length>=2){ini=(parts[0][0]+parts[parts.length-1][0]).toUpperCase();}
    else if(parts.length===1){ini=parts[0].substring(0,2).toUpperCase();}
    else{ini='?';}
    var av=document.getElementById('userAvatar');if(av)av.textContent=ini;
})();

var EXT_ICONS={
    '.pdf':'&#128196;','.doc':'&#128196;','.docx':'&#128196;','.xls':'&#128200;','.xlsx':'&#128200;',
    '.ppt':'&#128202;','.pptx':'&#128202;','.txt':'&#128196;','.csv':'&#128200;','.odt':'&#128196;',
    '.ods':'&#128200;','.odp':'&#128202;','.rtf':'&#128196;','.md':'&#128196;','.log':'&#128203;',
    '.ps1':'&#128187;','.py':'&#128013;','.js':'&#129300;','.ts':'&#129300;','.html':'&#127760;',
    '.htm':'&#127760;','.css':'&#127912;','.php':'&#128187;','.java':'&#9749;','.cs':'&#128187;',
    '.cpp':'&#128187;','.c':'&#128187;','.h':'&#128187;','.go':'&#128187;','.rs':'&#128187;',
    '.rb':'&#128312;','.sh':'&#128192;','.bat':'&#128192;','.cmd':'&#128192;','.sql':'&#128020;',
    '.xml':'&#128196;','.json':'&#128196;','.yaml':'&#128196;','.yml':'&#128196;','.ini':'&#9881;',
    '.conf':'&#9881;','.config':'&#9881;','.env':'&#9881;','.toml':'&#128196;',
    '.jpg':'&#128444;','.jpeg':'&#128444;','.png':'&#128444;','.gif':'&#127916;','.bmp':'&#128444;',
    '.svg':'&#128444;','.ico':'&#128444;','.webp':'&#128444;','.tiff':'&#128444;','.tif':'&#128444;',
    '.psd':'&#127912;','.ai':'&#127912;','.eps':'&#127912;','.raw':'&#128247;',
    '.mp4':'&#127916;','.avi':'&#127916;','.mkv':'&#127916;','.mov':'&#127916;','.wmv':'&#127916;',
    '.flv':'&#127916;','.webm':'&#127916;','.m4v':'&#127916;','.mpg':'&#127916;','.mpeg':'&#127916;',
    '.mp3':'&#127925;','.wav':'&#127925;','.flac':'&#127925;','.aac':'&#127925;','.ogg':'&#127925;',
    '.m4a':'&#127925;','.wma':'&#127925;',
    '.zip':'&#128230;','.rar':'&#128230;','.7z':'&#128230;','.tar':'&#128230;','.gz':'&#128230;',
    '.bz2':'&#128230;','.xz':'&#128230;','.iso':'&#128191;','.dmg':'&#128191;',
    '.exe':'&#9881;','.msi':'&#9881;','.dll':'&#9881;','.so':'&#9881;','.deb':'&#9881;',
    '.rpm':'&#9881;','.apk':'&#128241;','.app':'&#9881;',
    '.db':'&#128020;','.sqlite':'&#128020;','.mdb':'&#128020;','.bak':'&#128190;','.mdf':'&#128020;','.ldf':'&#128020;',
    '.ttf':'&#128210;','.otf':'&#128210;','.woff':'&#128210;','.woff2':'&#128210;',
    '.torrent':'&#128279;','.ics':'&#128197;','.vcf':'&#128100;','.eml':'&#128140;','.msg':'&#128140;','.pst':'&#128140;',
    'dir':'&#128193;','net':'&#127760;','denied':'&#128274;','unknown':'&#128196;'
};
function getFileIcon(ext,isDir,isDenied,path){
    if(isDenied)return EXT_ICONS['denied'];
    if(isDir){if(path&&path.startsWith('\\\\'))return EXT_ICONS['net'];return EXT_ICONS['dir'];}
    return EXT_ICONS[ext.toLowerCase()]||EXT_ICONS['unknown'];
}

function toggleTheme(){
    var h=document.documentElement,isLight=h.getAttribute('data-theme')==='light';
    h.setAttribute('data-theme',isLight?'dark':'light');
    document.getElementById('themeBtn').innerHTML=isLight?'&#9790; Dark':'&#9728; Light';
    try{localStorage.setItem('nn-theme',isLight?'dark':'light');}catch(e){}
}
(function(){
    try{var t=localStorage.getItem('nn-theme');if(t){document.documentElement.setAttribute('data-theme',t);var b=document.getElementById('themeBtn');if(b)b.innerHTML=t==='light'?'&#9728; Light':'&#9790; Dark';}}catch(e){}
})();

function toggleExcl(){var b=document.getElementById('exclToggle'),d=document.getElementById('exclBody');if(b&&d){b.classList.toggle('open');d.classList.toggle('open');}}
function toggleDetail(id){var el=document.getElementById(id);if(el)el.style.display=(el.style.display==='block')?'none':'block';}

function escHtml(s){if(!s)return '';return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
function normPath(p){
    if(!p)return '';var isUnc=p.startsWith('\\\\');p=p.replace(/\//g,'\\');
    if(isUnc){p='\\\\'+p.substring(2).replace(/\\{2,}/g,'\\');}else{p=p.replace(/\\{2,}/g,'\\');}
    if(p.length>3&&!isUnc)p=p.replace(/\\+`$`/,'');
    if(/^[A-Za-z]:`$`/.test(p))p+='\\';return p;
}
function parentOf(p){
    p=normPath(p);var isUnc=p.startsWith('\\\\');
    if(/^[A-Za-z]:\\`$`/.test(p))return null;
    if(isUnc){var parts=p.substring(2).split('\\').filter(Boolean);if(parts.length<=2)return null;return '\\\\'+parts.slice(0,parts.length-1).join('\\');}
    var i=p.lastIndexOf('\\');if(i<0)return null;var r=p.substring(0,i);if(/^[A-Za-z]:`$`/.test(r))r+='\\';return r||null;
}
function formatSize(b){b=Number(b);if(b<0)return '?';if(b>=1099511627776)return (b/1099511627776).toFixed(2)+' TB';if(b>=1073741824)return (b/1073741824).toFixed(2)+' GB';if(b>=1048576)return (b/1048576).toFixed(2)+' MB';if(b>=1024)return (b/1024).toFixed(2)+' KB';return b+' B';}
function barColor(sz){if(sz>=1073741824)return 'var(--danger)';if(sz>=104857600)return 'var(--warning)';if(sz>=10485760)return '#c47d00';return 'var(--success)';}
function isExcluded(path){path=normPath(path).toLowerCase();for(var i=0;i<EXCLUDED.length;i++){var ex=normPath(EXCLUDED[i]).toLowerCase();if(path===ex||path.startsWith(ex+'\\'))return true;}return false;}
function findScan(p){p=normPath(p).toLowerCase();for(var k in DATA){if(normPath(k).toLowerCase()===p)return DATA[k];}return null;}
function navigateTo(p){if(!p||!p.trim())return;p=normPath(p.trim());curPath=p;document.getElementById('pathInput').value=p;renderBreadcrumb(p);renderContent(p);}
function rescanFolder(p){
    if(typeof PSNCDU_TOKEN==='undefined'){alert('Rescan indisponible : rapport ouvert hors du serveur PS-NCDU.');return;}
    var d=(typeof MAX_DEPTH==='number'&&MAX_DEPTH>0)?MAX_DEPTH:3;
    window.location='/?token='+PSNCDU_TOKEN+'&scan='+encodeURIComponent(p)+'&depth='+d;
}
function goUp(){var p=parentOf(curPath);if(p)navigateTo(p);}

function renderBreadcrumb(path){
    path=normPath(path);var isUnc=path.startsWith('\\\\');var html='',built='';
    if(isUnc){
        var parts=path.substring(2).split('\\').filter(Boolean);
        if(parts.length>=1){built='\\\\'+parts[0];html+='<span style="color:var(--text-dim)">&#127760;</span> ';if(parts.length===1){html+='<a class="active">'+escHtml('\\\\'+parts[0])+'</a>';}else{html+='<a onclick="navigateTo(\''+built.replace(/\\/g,'\\\\')+'\')">'+escHtml('\\\\'+parts[0])+'</a>';}}
        for(var i=1;i<parts.length;i++){built+='\\'+parts[i];html+=' <span class="sep">&#8250;</span> ';if(i===parts.length-1){html+='<a class="active">'+escHtml(parts[i])+'</a>';}else{var b=built;html+='<a onclick="navigateTo(\''+b.replace(/\\/g,'\\\\')+'\')">' +escHtml(parts[i])+'</a>';}}
    }else{
        var parts2=path.replace(/\\+`$`/,'').split('\\').filter(Boolean);
        if(/^[A-Za-z]:/.test(path)){built=parts2[0]+'\\';if(parts2.length===1){html+='<a class="active">'+escHtml(parts2[0])+'</a>';}else{html+='<a onclick="navigateTo(\''+built.replace(/\\/g,'\\\\')+'\')">'+escHtml(parts2[0])+'</a>';}parts2.shift();}
        parts2.forEach(function(p,i){built+=p+'\\';html+=' <span class="sep">&#8250;</span> ';if(i===parts2.length-1){html+='<a class="active">'+escHtml(p)+'</a>';}else{var b=built;html+='<a onclick="navigateTo(\''+b.replace(/\\/g,'\\\\')+'\')">' +escHtml(p)+'</a>';}});
    }
    document.getElementById('breadcrumb').innerHTML=html||'&#127968; Racine';
}

function renderContent(path){
    if(isExcluded(path)){
        document.getElementById('stats').innerHTML='<div class="chip chip-excl"><span class="chip-label">&#9888;</span><span class="chip-value">Repertoire exclu</span></div>';
        document.getElementById('mainContent').innerHTML='<div class="msg-page"><span class="icon">&#9888;</span><h3>Repertoire exclu du scan</h3><p>Sa taille est inconnue.</p><code>'+escHtml(path)+'</code></div>';
        return;
    }
    var scan=findScan(path);
    var isUnscannedDir=false,isDeniedDir=false;
    if(!scan){
        var par=parentOf(path);
        if(par){var parScan=findScan(par);if(parScan){parScan.items.forEach(function(it){var np=normPath(it.fullPath).toLowerCase(),npath=normPath(path).toLowerCase();if(np===npath){if(it.unscanned)isUnscannedDir=true;if(it.denied)isDeniedDir=true;}});}}
    }
    if(!scan){
        if(isDeniedDir){
            document.getElementById('stats').innerHTML='<div class="chip chip-denied"><span class="chip-label">&#128274;</span><span class="chip-value">Protege ACL NTFS</span></div>';
            document.getElementById('mainContent').innerHTML=
                '<div class="msg-page"><span class="icon">&#128274;</span>'+
                '<h3>Dossier protege par ACL NTFS</h3>'+
                '<p>Ce dossier est inaccessible avec le compte <strong>$($env:USERNAME)</strong>.</p>'+
                '<p style="margin-top:8px;color:var(--text-muted);font-size:.85em">C\'est normal pour :</p>'+
                '<ul class="protected-list">'+
                '<li>Les dossiers systeme Windows</li>'+
                '<li>Les profils d\'autres utilisateurs</li>'+
                '<li>Les dossiers de protection de donnees</li>'+
                '</ul>'+
                '<p class="hint-denied">Pour y acceder : demandez les droits NTFS a votre administrateur</p>'+
                '<code>'+escHtml(path)+'</code></div>';
        }else if(isUnscannedDir){
            document.getElementById('stats').innerHTML='<div class="chip chip-unscanned"><span class="chip-label">&#128269;</span><span class="chip-value">Non scanne - profondeur '+MAX_DEPTH+'</span></div>';
            document.getElementById('mainContent').innerHTML='<div class="msg-page"><span class="icon">&#128194;</span><h3>Dossier non scanne</h3><p>Ce dossier est au-dela de la profondeur <strong>'+MAX_DEPTH+'</strong> du scan actuel.</p><code>'+escHtml(path)+'</code><br><button class="rescan-btn" onclick="rescanFolder(\''+path.replace(/\\/g,'\\\\')+'\')">&#128269; Scanner ce dossier</button></div>';
        }else{
            var msg=UNLIMITED?'Ce dossier n\'est pas dans le rapport.':'Ce dossier depasse la profondeur '+MAX_DEPTH+' du scan actuel.';
            document.getElementById('stats').innerHTML='<div class="chip chip-warn"><span class="chip-label">&#9888;</span><span class="chip-value">Non scanne</span></div>';
            document.getElementById('mainContent').innerHTML='<div class="msg-page"><span class="icon">&#128194;</span><h3>Dossier non scanne</h3><p>'+msg+'</p><code>'+escHtml(path)+'</code><br><button class="rescan-btn" onclick="rescanFolder(\''+path.replace(/\\/g,'\\\\')+'\')">&#128269; Scanner ce dossier</button></div>';
        }
        return;
    }
    var items=scan.items,total=scan.total,nbD=0,nbF=0,nbExcl=0,nbUnscanned=0,nbDenied=0;
    items.forEach(function(i){if(i.excluded)nbExcl++;else if(i.unscanned)nbUnscanned++;else if(i.denied)nbDenied++;else if(i.isDir)nbD++;else nbF++;});
    var st='<div class="chip"><span class="chip-label">Taille scannee</span><span class="chip-value">'+formatSize(total)+'</span></div>'+
        '<div class="chip"><span class="chip-label">Dossiers</span><span class="chip-value">'+nbD+'</span></div>'+
        '<div class="chip"><span class="chip-label">Fichiers</span><span class="chip-value">'+nbF+'</span></div>'+
        '<div class="chip"><span class="chip-label">Elements</span><span class="chip-value">'+items.length+'</span></div>';
    if(nbDenied>0)st+='<div class="chip chip-denied"><span class="chip-label">&#128274; Proteges</span><span class="chip-value">'+nbDenied+'</span></div>';
    if(nbUnscanned>0)st+='<div class="chip chip-unscanned"><span class="chip-label">&#128269; Non scannes</span><span class="chip-value">'+nbUnscanned+'</span></div>';
    if(nbExcl>0)st+='<div class="chip chip-excl"><span class="chip-label">&#9888; Exclus</span><span class="chip-value">'+nbExcl+'</span></div>';
    document.getElementById('stats').innerHTML=st;
    if(items.length===0){document.getElementById('mainContent').innerHTML='<div class="msg-page"><span class="icon">&#128194;</span><h3>Dossier vide</h3><p>Aucun fichier accessible.</p></div>';return;}
    var rows='';
    items.forEach(function(item){
        var isExcl=item.excluded||isExcluded(item.fullPath);
        var isUnscanned=item.unscanned,isDenied=item.denied;
        var fp=item.fullPath.replace(/\\/g,'\\\\');
        var icon=getFileIcon(item.ext||'',item.isDir,isDenied,item.fullPath);
        var isDeep=!UNLIMITED&&item.isDir&&!isExcl&&!isUnscanned&&!isDenied&&!findScan(item.fullPath);
        var tag=isDeep?'<span style="font-size:.7em;color:var(--text-dim);background:var(--card);padding:1px 6px;border-radius:3px;margin-left:6px;border:1px solid var(--border)">+</span>':'';
        var nameCell,sizeCell,barCell,badge,rowClass,tipTitle;
        if(isDenied){
            rowClass='row-denied';tipTitle='Protege par ACL NTFS - droits d acces requis';
            nameCell='<span class="dir-link-denied" title="'+tipTitle+'">'+escHtml(item.name)+'</span><span class="denied-label">&#128274; Protege</span>';
            sizeCell='<span class="size-denied" title="'+tipTitle+'">?</span>';
            barCell='<div class="bar-wrap"><div class="bar-bg"><div class="bar-denied"></div></div><span class="bar-denied-text" title="'+tipTitle+'">&#128274;</span></div>';
            badge='<span class="type-badge type-denied" title="'+tipTitle+'">ACL</span>';
        }else if(isExcl){
            rowClass='row-excluded';tipTitle='Repertoire exclu - taille inconnue';
            nameCell='<a class="dir-link" onclick="navigateTo(\''+fp+'\')" title="'+tipTitle+'">'+escHtml(item.name)+'</a><span class="excl-label">&#9888;</span>';
            sizeCell='<span class="size-unknown">?</span>';
            barCell='<div class="bar-wrap"><div class="bar-bg"><div class="bar-excl"></div></div><span class="bar-unknown">?</span></div>';
            badge='<span class="type-badge type-excl">EXCLU</span>';
        }else if(isUnscanned){
            rowClass='row-unscanned';tipTitle='Taille inconnue - scan limite a la profondeur '+MAX_DEPTH+'. Utilisez profondeur 0.';
            nameCell='<a class="dir-link-unscanned" onclick="navigateTo(\''+fp+'\')" title="'+tipTitle+'">'+escHtml(item.name)+'</a>';
            sizeCell='<span class="size-unscanned" title="'+tipTitle+'">?</span>';
            barCell='<div class="bar-wrap"><div class="bar-bg"><div class="bar-unscanned"></div></div><span class="bar-unscanned-text" title="'+tipTitle+'">?</span></div>';
            badge='<span class="type-badge type-unscanned" title="'+tipTitle+'">?</span>';
        }else{
            rowClass='';var pct=total>0?Math.round((item.size/total)*1000)/10:0;
            nameCell=item.isDir?'<a class="dir-link" onclick="navigateTo(\''+fp+'\')">'+escHtml(item.name)+'</a>'+tag:'<span class="file-name">'+escHtml(item.name)+'</span>';
            sizeCell=escHtml(item.sizeStr);
            barCell='<div class="bar-wrap"><div class="bar-bg"><div class="bar-fill" style="width:'+pct+'%;background:'+barColor(item.size)+'"></div></div><span class="bar-pct">'+pct+'%</span></div>';
            badge=item.isDir?'<span class="type-badge type-dir">DIR</span>':'<span class="type-badge type-file">FILE</span>';
        }
        rows+='<tr class="'+rowClass+'"><td class="col-icon">'+icon+'</td><td class="col-name">'+nameCell+'</td><td class="col-size">'+sizeCell+'</td><td class="col-bar">'+barCell+'</td><td class="col-type">'+badge+'</td></tr>';
    });
    document.getElementById('mainContent').innerHTML=
        '<div class="table-wrap"><table><thead><tr><th class="col-icon"></th><th id="th-name" onclick="sortBy(\'name\')">Nom</th><th id="th-size" onclick="sortBy(\'size\')" style="text-align:right">Taille</th><th>Utilisation</th><th class="col-type">Type</th></tr></thead><tbody>'+rows+'</tbody></table></div>';
    ['th-name','th-size'].forEach(function(id){var el=document.getElementById(id);if(el)el.classList.remove('sort-asc','sort-desc');});
    var thEl=document.getElementById(sortCol==='name'?'th-name':'th-size');
    if(thEl)thEl.classList.add(sortAsc?'sort-asc':'sort-desc');
}

function sortBy(col){
    if(sortCol===col){sortAsc=!sortAsc;}else{sortCol=col;sortAsc=(col==='name');}
    var scan=findScan(curPath);if(!scan)return;
    var normal=scan.items.filter(function(i){return !i.excluded&&!i.unscanned&&!i.denied;});
    var unscanned=scan.items.filter(function(i){return i.unscanned;});
    var denied=scan.items.filter(function(i){return i.denied;});
    var excl=scan.items.filter(function(i){return i.excluded;});
    normal.sort(function(a,b){var va=a[col],vb=b[col];if(typeof va==='string'){va=va.toLowerCase();vb=vb.toLowerCase();}if(va<vb)return sortAsc?-1:1;if(va>vb)return sortAsc?1:-1;return 0;});
    scan.items=normal.concat(unscanned).concat(denied).concat(excl);
    renderContent(curPath);
}
window.onload=function(){navigateTo(ROOT_PATH);};
</script>
</body>
</html>
"@

    Write-Progress -Activity $ACT -Completed
    Write-Step "[HTML] Generation complete" $t
    return $html
}

# ============================================================
# POINT D ENTREE
# ============================================================
# ============================================================
# v4.0 : Application web locale (HttpListener)
# ============================================================

# --- Verif FullLanguage (le serveur exige les types .NET) -----
if ($ExecutionContext.SessionState.LanguageMode -ne 'FullLanguage') {
    Clear-Host
    Write-Host ""
    Write-Host "  PS-NCDU v$SCRIPT_VERSION - Application web" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  ERREUR : LanguageMode = $($ExecutionContext.SessionState.LanguageMode)" -ForegroundColor Red
    Write-Host "  Le serveur web exige le mode FullLanguage (HttpListener)." -ForegroundColor Yellow
    Write-Host "  Lancez ce script sur un poste en FullLanguage, ou utilisez" -ForegroundColor Yellow
    Write-Host "  la v3.2 (rapport HTML statique) qui tourne en ConstrainedLanguage." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Appuyez sur ENTER pour quitter..." -ForegroundColor DarkGray
    $null = Read-Host
    exit 1
}

$script:FullUserName   = Get-FullUserName
$script:Token          = [guid]::NewGuid().ToString('N')
$script:LastResultHtml = $null
$script:ScanBusy       = $false
$script:CancelScan     = $false

# --- Liste des lecteurs pour l'UI ---------------------------
function Get-DrivesJson {
    $parts = @()
    try {
        foreach ($d in [System.IO.DriveInfo]::GetDrives()) {
            $type = ""
            try { $type = [string]$d.DriveType } catch {}
            $ready = $false; $freeGB = 0; $totalGB = 0; $label = ""
            if ($type -eq 'Network') {
                # NE PAS interroger un lecteur reseau : IsReady/TotalSize peuvent
                # bloquer longtemps (VPN, partage lent ou mort) et geler le serveur.
                $ready = $true; $label = ""
            } else {
                try {
                    $ready = $d.IsReady
                    if ($ready) {
                        $freeGB  = [math]::Round($d.AvailableFreeSpace / 1GB, 1)
                        $totalGB = [math]::Round($d.TotalSize / 1GB, 1)
                        $label   = $d.VolumeLabel
                    }
                } catch {}
            }
            $parts += "{""name"":""$(ConvertTo-JsonSafe $d.Name)"",""type"":""$(ConvertTo-JsonSafe $type)"",""label"":""$(ConvertTo-JsonSafe $label)"",""ready"":$(if($ready){'true'}else{'false'}),""freeGB"":$freeGB,""totalGB"":$totalGB}"
        }
    } catch { Write-Log "[WEB] Get-Drives erreur : $_" -Level WARN }
    return "[" + ($parts -join ",") + "]"
}

# --- Acces rapides (users) + historique --------------------
$script:HistoryFile = "$scriptTemp\history.txt"

function Add-History {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    try {
        $hist = @()
        if (Test-Path -LiteralPath $script:HistoryFile) { $hist = @(Get-Content -LiteralPath $script:HistoryFile -ErrorAction SilentlyContinue) }
        $hist = @($Path) + ($hist | Where-Object { $_ -and $_.Trim() -ne '' -and $_.ToLower() -ne $Path.ToLower() })
        $hist | Select-Object -First 15 | Set-Content -LiteralPath $script:HistoryFile -Encoding UTF8 -ErrorAction SilentlyContinue
    } catch {}
}

function Remove-History {
    param([string]$Path,[switch]$All)
    try {
        if ($All) {
            if (Test-Path -LiteralPath $script:HistoryFile) { Remove-Item -LiteralPath $script:HistoryFile -Force -ErrorAction SilentlyContinue }
            return
        }
        if ([string]::IsNullOrWhiteSpace($Path)) { return }
        if (Test-Path -LiteralPath $script:HistoryFile) {
            $keep = @(Get-Content -LiteralPath $script:HistoryFile -ErrorAction SilentlyContinue | Where-Object { $_ -and $_.Trim() -ne '' -and $_.Trim().ToLower() -ne $Path.Trim().ToLower() })
            $keep | Set-Content -LiteralPath $script:HistoryFile -Encoding UTF8 -ErrorAction SilentlyContinue
        }
    } catch {}
}

function Get-QuickJson {
    $users = New-Object System.Collections.Generic.List[object]
    $hist  = New-Object System.Collections.Generic.List[object]
    try {
        $cur = "$env:USERPROFILE"
        if ($cur -and (Test-Path -LiteralPath $cur)) { [void]$users.Add("{""path"":""$(ConvertTo-JsonSafe $cur)"",""name"":""$(ConvertTo-JsonSafe (Split-Path $cur -Leaf))""}") }
        foreach ($u in ([System.IO.DirectoryInfo]::new("C:\Users").EnumerateDirectories())) {
            if (($u.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { continue }
            if ($u.FullName -ieq $cur) { continue }
            [void]$users.Add("{""path"":""$(ConvertTo-JsonSafe $u.FullName)"",""name"":""$(ConvertTo-JsonSafe $u.Name)""}")
        }
    } catch { Write-Log "[WEB] Get-Quick users erreur : $_" -Level WARN }
    try {
        if (Test-Path -LiteralPath $script:HistoryFile) {
            foreach ($h in @(Get-Content -LiteralPath $script:HistoryFile -ErrorAction SilentlyContinue)) {
                if ($h -and $h.Trim() -ne '') { [void]$hist.Add("""$(ConvertTo-JsonSafe $h.Trim())""") }
            }
        }
    } catch {}
    return "{""users"":[$($users -join ',')],""history"":[$($hist -join ',')]}"
}

# --- Fichiers directs d'un dossier (a la demande, trie, plafonne) ---
function Get-FilesJson {
    param([string]$Path)
    $items = New-Object System.Collections.Generic.List[object]
    $capped = 'false'
    try {
        $files = @(Get-DirectFiles -Path $Path) | Where-Object { $null -ne $_ } | Sort-Object -Property Length -Descending
        if ($files.Count -gt 1000) { $capped = 'true' }
        $n = 0
        foreach ($f in $files) {
            if ($n -ge 1000) { break }
            $ext = ""
            try { $ext = ([string]$f.Extension).ToLower() } catch {}
            [void]$items.Add("{""name"":""$(ConvertTo-JsonSafe $f.Name)"",""size"":$([long]$f.Length),""ext"":""$(ConvertTo-JsonSafe $ext)""}")
            $n++
        }
    } catch { return "{""files"":[],""capped"":false,""err"":true}" }
    return "{""files"":[$($items -join ',')],""capped"":$capped}"
}

# --- Selecteur de dossier natif Windows (thread STA) --------
function Show-FolderPicker {
    param([string]$Initial)
    # v6.19 : le dialogue WinForms exige un thread STA. En PS 5.1 la console
    # est deja STA -> on l'ouvre directement (aucun marshaling, donc pas de
    # blocage). En PS 7 (MTA) on ouvre un runspace STA dedie. Une fenetre
    # proprietaire TopMost force le dialogue AU PREMIER PLAN, sinon il peut
    # s'ouvrir derriere le navigateur et sembler absent.
    $core = {
        param($Init)
        Add-Type -AssemblyName System.Windows.Forms
        $owner = New-Object System.Windows.Forms.Form
        $owner.TopMost = $true; $owner.ShowInTaskbar = $false
        $owner.Width = 1; $owner.Height = 1; $owner.StartPosition = 'CenterScreen'
        $owner.Show(); $owner.Activate()
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = 'Selectionnez un dossier a analyser'
        $dlg.ShowNewFolderButton = $false
        if ($Init -and (Test-Path -LiteralPath $Init)) { $dlg.SelectedPath = $Init }
        $picked = ''
        if ($dlg.ShowDialog($owner) -eq [System.Windows.Forms.DialogResult]::OK) { $picked = $dlg.SelectedPath }
        $owner.Close(); $owner.Dispose()
        return $picked
    }
    $path = ''; $err = ''
    try {
        if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -eq [System.Threading.ApartmentState]::STA) {
            $path = & $core $Initial
        } else {
            $rs = [runspacefactory]::CreateRunspace()
            $rs.ApartmentState = [System.Threading.ApartmentState]::STA
            $rs.ThreadOptions  = 'ReuseThread'
            $rs.Open()
            $ps = [powershell]::Create(); $ps.Runspace = $rs
            [void]$ps.AddScript($core.ToString()).AddArgument($Initial)
            $out = $ps.Invoke()
            if ($ps.HadErrors -and $ps.Streams.Error.Count -gt 0) { $err = [string]$ps.Streams.Error[0] }
            foreach ($o in $out) { if ($o -is [string] -and $o) { $path = $o } }
            $ps.Dispose(); $rs.Close(); $rs.Dispose()
        }
    } catch { $err = [string]$_; Write-Log "[WEB] FolderPicker erreur : $_" -Level WARN }
    if ($err) { Write-Log "[WEB] FolderPicker indisponible : $err" -Level WARN }
    return @{ Path = $path; Err = $err }
}

# --- Modes de scan pour l'UI (depuis $SCAN_MODES) -----------
$modesJsonParts = @()
foreach ($k in ($SCAN_MODES.Keys | Sort-Object)) {
    $m = $SCAN_MODES[$k]
    $modesJsonParts += "{""key"":$k,""name"":""$(ConvertTo-JsonSafe $m.Name)"",""speed"":""$(ConvertTo-JsonSafe $m.Speed)"",""acc"":""$(ConvertTo-JsonSafe $m.Accuracy)""}"
}
$MODES_JSON = "[" + ($modesJsonParts -join ",") + "]"

# --- Helpers reponses HTTP ----------------------------------
function Send-Bytes {
    param($Resp, [byte[]]$Bytes, [string]$ContentType)
    try {
        $Resp.StatusCode  = 200
        $Resp.ContentType = $ContentType
        $Resp.ContentLength64 = $Bytes.Length
        $Resp.OutputStream.Write($Bytes, 0, $Bytes.Length)
    } catch { Write-Log "[WEB] Send-Bytes erreur : $_" -Level WARN }
    finally { try { $Resp.OutputStream.Close() } catch {} }
}
function Send-Text {
    param($Resp, [string]$Text, [string]$ContentType = "text/html; charset=utf-8")
    $enc = New-Object System.Text.UTF8Encoding($false)
    Send-Bytes -Resp $Resp -Bytes ($enc.GetBytes($Text)) -ContentType $ContentType
}
function Send-Status {
    param($Resp, [int]$Code, [string]$Text = "")
    try { $Resp.StatusCode = $Code } catch {}
    $enc = New-Object System.Text.UTF8Encoding($false)
    $b = $enc.GetBytes($Text)
    try { $Resp.ContentLength64 = $b.Length; $Resp.OutputStream.Write($b,0,$b.Length) } catch {}
    finally { try { $Resp.OutputStream.Close() } catch {} }
}

# --- Page de reglages (SPA) ---------------------------------
# Here-string a quotes simples : le $ et les ${} du JS restent litteraux.
# On injecte les valeurs via .Replace() sur des marqueurs __XXX__.
$SETTINGS_HTML = @'
<!DOCTYPE html>
<html lang="fr" data-theme="light">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>PS-NCDU v__VERSION__</title>
    <style>
*{box-sizing:border-box;margin:0;padding:0}
html,body{height:100%}
:root{--bg:#f0f4f8;--surface:#fff;--card:#e8edf3;--card-hover:#dde4ed;--border:#cdd6e0;--text:#0f1923;--text-muted:#4a6080;--text-dim:#8da0b8;--accent:#0063be;--accent-light:#004fa3;--accent-glow:rgba(0,99,190,.08);--success:#007a5e;--warning:#e0932b;--danger:#c0202e;--radius:8px;--radius-sm:5px;--radius-pill:20px}
[data-theme="dark"]{--bg:#0f1923;--surface:#162032;--card:#1e2d42;--card-hover:#243350;--border:#2a3f5f;--text:#e6edf5;--text-muted:#9fb3c8;--text-dim:#6b7f96;--accent:#3a9bff;--accent-light:#6cb6ff;--accent-glow:rgba(58,155,255,.15);--success:#00c48c;--warning:#f0a94b;--danger:#e8394a}
body{font-family:'Segoe UI',Roboto,Arial,sans-serif;background:var(--bg);color:var(--text);font-size:14px;overflow:hidden}
.app{display:flex;flex-direction:column;height:100vh}
.appbar{display:flex;align-items:center;gap:16px;padding:0 18px;height:56px;background:var(--surface);border-bottom:1px solid var(--border);flex-shrink:0}
.brand{display:flex;align-items:center;gap:10px;color:var(--accent);flex-shrink:0}
.brand .bname{font-weight:800;font-size:1.05em;color:var(--text)}
.brand .ver{font-size:.72em;color:var(--text-dim);font-weight:600;margin-left:2px;align-self:flex-start}
.brand .bname i{color:var(--accent-light);font-style:normal}
.bc{flex:1;min-width:0;display:flex;align-items:center;flex-wrap:nowrap;overflow:hidden;font-size:.95em}
.bc:empty::before{content:"";color:var(--text-dim);font-size:.9em}
.tcrumb{color:var(--accent-light);cursor:pointer;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:260px}
.tcrumb:hover{text-decoration:underline}
.tsep{color:var(--text-dim);margin:0 6px;flex-shrink:0}
.appact{display:flex;align-items:center;gap:12px;flex-shrink:0}
.pctbadge{font-weight:700;color:var(--accent-light);font-size:1.05em;min-width:42px;text-align:right}
.pctbadge:empty{display:none}
.appact .ttotal{font-size:.9em;color:var(--text-muted);white-space:nowrap}
.appact .ttotal strong{color:var(--text)}
.tactive-tag{background:var(--accent-glow);color:var(--accent-light);border:1px solid var(--accent-light);border-radius:var(--radius-pill);padding:0 8px;font-size:.82em;margin-left:6px}
.iconbtn{background:var(--card);border:1px solid var(--border);color:var(--text-muted);width:34px;height:34px;border-radius:var(--radius-sm);cursor:pointer;font-size:1.05em;line-height:1}
.iconbtn:hover{border-color:var(--accent-light);color:var(--accent-light)}
.progline{flex-shrink:0}
.gbar{position:relative;height:4px;background:var(--card);overflow:hidden}
.gbar>i{display:block;height:100%;width:0;background:linear-gradient(90deg,var(--accent),var(--accent-light));transition:width .3s}
.gsweep{position:absolute;top:0;left:-45%;height:100%;width:45%;background:linear-gradient(90deg,transparent,rgba(255,255,255,.5),transparent);display:none;pointer-events:none}
.gbar.loading .gsweep{display:block;animation:gsweep 1.1s linear infinite}
@keyframes gsweep{from{left:-45%}to{left:100%}}
.treearea{flex:1;overflow-y:auto;padding:10px 12px}
.tlist{max-width:1500px;margin:0 auto}
.trow{display:flex;align-items:center;gap:12px;padding:9px 12px;border-radius:var(--radius-sm);font-size:.92em;border-bottom:1px solid var(--border)}
.trow:last-child{border-bottom:none}
.trow[onclick]{cursor:pointer}
.trow[onclick]:hover{background:var(--card-hover)}
.trow .tname{flex:1;min-width:120px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:var(--text)}
.ticon{width:24px;display:inline-flex;align-items:center;justify-content:center;flex-shrink:0}
.ticon .ic{position:relative;display:inline-flex;line-height:1;font-size:1.1em}
.sdot{position:absolute;right:-3px;bottom:-2px;width:9px;height:9px;border-radius:50%;border:2px solid var(--surface);box-sizing:border-box}
.sdot-green{background:var(--success)}
.sdot-orange{background:var(--warning)}
.sdot-grey{background:var(--text-dim)}
.sdot-blue{background:var(--accent)}
.trow .tbarwrap{width:240px;height:9px;background:var(--card);border-radius:var(--radius-pill);overflow:hidden;flex-shrink:0}
.trow .tbar{display:block;height:100%;background:linear-gradient(90deg,var(--accent),var(--accent-light));border-radius:var(--radius-pill)}
.trow .tpct{width:46px;text-align:right;color:var(--text-muted);font-size:.85em;flex-shrink:0}
.trow .tsize{width:96px;text-align:right;color:var(--text-muted);font-family:monospace;font-size:.9em;white-space:nowrap;flex-shrink:0}
.trow .tcount{width:150px;text-align:right;color:var(--text-dim);font-size:.82em;white-space:nowrap;flex-shrink:0}
.tcnt{color:var(--text-dim);font-size:.88em;margin-left:10px;white-space:nowrap}
.trow.tup{color:var(--text-muted)}
.trow.tup .tname{color:var(--text-muted)}
.trow.active{background:var(--accent-glow)}
.trow.active .tname{color:var(--accent-light);font-weight:600}
.trow.pending{opacity:.85}
.trow.pending .tbarwrap{position:relative;overflow:hidden}
.trow.pending .tbarwrap::after{content:"";position:absolute;top:0;left:0;height:100%;width:100%;background:linear-gradient(90deg,transparent,var(--card-hover),transparent);transform:translateX(-100%);animation:tshimmer 1.4s ease-in-out infinite}
@keyframes tshimmer{100%{transform:translateX(100%)}}
.trow.pending .tsize{color:var(--text-dim)}
.trow.skip{opacity:.72}
.trow.done .ic{color:var(--accent-light)}
.trow.file{opacity:.95}
.trow.file .fic{color:var(--text-dim);font-size:1.05em}
.trow.file .tname{color:var(--text-muted)}
.trow .tbar-file{background:linear-gradient(90deg,var(--text-dim),var(--text-muted));opacity:.7}
.tempty{padding:36px;color:var(--text-dim);font-size:.9em;text-align:center}
.spin{width:14px;height:14px;border:2px solid var(--accent-glow);border-top-color:var(--accent);border-radius:50%;box-sizing:border-box;display:inline-block;animation:psspin .7s linear infinite}
@keyframes psspin{to{transform:rotate(360deg)}}
.statusbar{display:flex;align-items:center;gap:14px;padding:0 16px;height:36px;background:var(--surface);border-top:1px solid var(--border);flex-shrink:0;font-size:.82em;color:var(--text-muted);overflow:hidden}
.stmain{flex:1;min-width:0;display:flex;align-items:center;gap:7px;overflow:hidden;white-space:nowrap}
.stmain:empty{display:none}
.stmain .spin{flex-shrink:0}
.stmain .lbl{color:var(--text-muted);font-weight:600;flex-shrink:0}
.stmain .scanparents{min-width:0;overflow:hidden;text-overflow:ellipsis;color:var(--text-dim)}
.stmain .segleaf{flex-shrink:0;color:var(--accent-light);font-weight:700;background:var(--accent-glow);border:1px solid var(--accent-light);padding:0 8px;border-radius:var(--radius-pill)}
.stmain .segsep{opacity:.5;margin:0 2px}
.ststage{color:var(--text-dim);white-space:nowrap;max-width:34%;overflow:hidden;text-overflow:ellipsis;flex-shrink:0}
.stdepth{color:var(--text-dim);white-space:nowrap;flex-shrink:0}
.stdepth:empty{display:none}
.stelapsed{white-space:nowrap;flex-shrink:0}
.stelapsed strong{color:var(--text-muted)}
.stbtn{flex-shrink:0;background:var(--card);border:1px solid var(--border);color:var(--text-muted);padding:3px 12px;border-radius:var(--radius-sm);cursor:pointer;font-size:.9em;font-family:inherit}
.stbtn:hover{border-color:var(--accent-light);color:var(--accent-light)}
.legend{position:fixed;right:14px;bottom:44px;width:min(440px,92vw);background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);box-shadow:0 8px 32px rgba(0,0,0,.28);z-index:60;display:none;overflow:hidden}
.legend.on{display:block}
.legtabs{display:flex;align-items:center;border-bottom:1px solid var(--border);background:var(--card)}
.legtab{background:none;border:none;border-bottom:2px solid transparent;color:var(--text-muted);padding:10px 16px;cursor:pointer;font-size:.86em;font-family:inherit;font-weight:600}
.legtab.on{color:var(--accent-light);border-bottom-color:var(--accent-light);background:var(--surface)}
.legclose{margin-left:auto;background:none;border:none;color:var(--text-dim);cursor:pointer;font-size:1em;padding:8px 14px}
.legclose:hover{color:var(--danger)}
.legpane{display:none;padding:12px 16px}
.legpane.on{display:block}
.legtitle{font-weight:700;font-size:.95em;margin-bottom:10px;color:var(--text)}
.legrow{display:flex;align-items:center;gap:10px;font-size:.85em;color:var(--text-muted);padding:5px 0}
.legrow .ic{position:relative;display:inline-flex;line-height:1;font-size:1.1em;width:22px;justify-content:center;flex-shrink:0}
.legrow .fic{color:var(--text-dim);font-size:1.05em;width:22px;text-align:center;flex-shrink:0}
.legrow .spin{flex-shrink:0;margin:0 4px}
.legrow2{font-size:.83em;color:var(--text-muted);line-height:1.55;padding:6px 0;border-bottom:1px solid var(--border)}
.legrow2:last-child{border-bottom:none}
.legrow2 b{color:var(--text)}
.legbar{width:22px;height:9px;background:var(--card);border-radius:var(--radius-pill);overflow:hidden;flex-shrink:0;display:inline-block}
.legbar>span{display:block;height:100%;width:65%;background:linear-gradient(90deg,var(--accent),var(--accent-light))}
.legnote{margin-top:10px;font-size:.78em;color:var(--text-dim);line-height:1.5;border-top:1px solid var(--border);padding-top:10px}
.derr{color:var(--danger);font-size:.85em;padding:8px 14px;text-align:center}
.derr:empty{display:none}
.overlay{position:fixed;inset:0;background:rgba(8,16,28,.55);display:none;align-items:flex-start;justify-content:center;z-index:100;overflow-y:auto;padding:44px 16px}
.overlay.on{display:flex}
.setcard{position:relative;background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);box-shadow:0 16px 56px rgba(0,0,0,.34);width:min(760px,96vw);padding:22px 26px;height:fit-content}
.sethead{display:flex;align-items:center;gap:12px;margin-bottom:14px;color:var(--accent)}
.settitle{font-weight:800;font-size:1.1em;color:var(--text)}
.setsub{font-size:.75em;color:var(--text-muted)}
.setcard .chip{margin-left:10px;font-size:.72em;background:var(--accent-glow);color:var(--accent-light);border:1px solid var(--accent-light);padding:2px 10px;border-radius:var(--radius-pill);font-weight:600}
h2{font-size:1.05em;margin-bottom:4px}
.hint{font-size:.82em;color:var(--text-muted);margin-bottom:16px}
label{display:block;font-size:.82em;font-weight:600;color:var(--text-muted);margin:14px 0 6px}
input[type=text],input[type=number]{width:100%;background:var(--card);border:1px solid var(--border);color:var(--text);padding:10px 14px;border-radius:var(--radius-sm);font-size:1em;font-family:inherit;outline:none}
input:focus{border-color:var(--accent-light);box-shadow:0 0 0 3px var(--accent-glow)}
.drives{display:flex;flex-wrap:wrap;gap:8px;margin-top:8px}
.pathrow{display:flex;gap:8px}
.pathrow input{flex:1;min-width:0}
.pathrow .btn{flex-shrink:0}
.quickwrap{margin-top:12px}
.quicklbl{font-size:.72em;font-weight:700;color:var(--text-dim);text-transform:uppercase;letter-spacing:.04em;margin-bottom:6px}
.chips{display:flex;flex-wrap:wrap;gap:8px}
input[type=range]{width:100%;accent-color:var(--accent);margin-top:8px;cursor:pointer}
.chkline{display:flex;align-items:center;gap:8px;font-weight:400;margin:10px 0 0;color:var(--text-muted);cursor:pointer}
.chkline input{width:auto;margin:0}
.explain{font-size:.8em;color:var(--text-muted);margin-top:8px;line-height:1.5;background:var(--card);border-left:3px solid var(--accent-light);padding:8px 12px;border-radius:0 var(--radius-sm) var(--radius-sm) 0}
select{width:100%;background:var(--card);border:1px solid var(--border);color:var(--text);padding:10px 14px;border-radius:var(--radius-sm);font-size:1em;font-family:inherit;outline:none;cursor:pointer;margin-top:6px}
select:focus{border-color:var(--accent-light);box-shadow:0 0 0 3px var(--accent-glow)}
label b{color:var(--accent-light)}
.drive{background:var(--card);border:1px solid var(--border);border-radius:var(--radius-sm);padding:8px 12px;cursor:pointer;font-size:.85em;text-align:left}
.drive:hover{border-color:var(--accent-light);background:var(--card-hover)}
.drive b{color:var(--accent-light)}.drive small{display:block;color:var(--text-dim);font-size:.9em}
.setclose{position:absolute;top:14px;right:14px;width:30px;height:30px;border:none;background:transparent;color:var(--text-dim);font-size:22px;line-height:1;cursor:pointer;border-radius:6px}
.setclose:hover{background:var(--card);color:var(--text)}
.browsecard{position:relative;background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);box-shadow:0 16px 56px rgba(0,0,0,.34);width:min(520px,96vw);max-height:80vh;display:flex;flex-direction:column;padding:18px 20px}
.browsehead{display:flex;align-items:center;justify-content:space-between;margin-bottom:10px}
.browsepath{font-family:monospace;font-size:.85em;color:var(--text-muted);word-break:break-all;margin-bottom:8px;min-height:1em}
.browselist{flex:1;overflow:auto;border:1px solid var(--border);border-radius:var(--radius-sm);padding:4px}
.browseitem{display:flex;align-items:center;gap:8px;padding:7px 10px;border-radius:var(--radius-sm);cursor:pointer;font-size:.92em}
.browseitem:hover{background:var(--card)}
.browseitem .bic{font-size:1.05em}
.browseitem.up{color:var(--accent-light)}
.browseact{display:flex;gap:10px;margin-top:12px}
.pathok{display:inline-flex;align-items:center;justify-content:center;width:20px;font-weight:700;font-size:1.05em}
.pathok.ok{color:#3fb968}.pathok.bad{color:#e0655f}
.dbar{display:block;height:4px;margin-top:7px;background:var(--border);border-radius:2px;overflow:hidden}
.dbar i{display:block;height:100%;background:var(--accent)}
.dbar.warn i{background:#e0a020}.dbar.bad i{background:#e0655f}
.exclwrap{margin:12px 0 0}
.excltoggle{background:none;border:none;color:var(--accent-light);cursor:pointer;font-size:.9em;padding:4px 0}
.exclbody{margin-top:8px}
.excllist{margin:6px 0;padding-left:18px;color:var(--text-dim);font-size:.85em;max-height:110px;overflow:auto}
.exclbody textarea{width:100%;box-sizing:border-box;font-family:monospace;font-size:.85em;resize:vertical}
.rec{display:inline-flex;align-items:stretch;gap:0}
.rec .drive{border-top-right-radius:0;border-bottom-right-radius:0}
.recx{border:1px solid var(--border);border-left:none;background:var(--card);color:var(--text-dim);cursor:pointer;padding:0 8px;border-top-right-radius:var(--radius-sm);border-bottom-right-radius:var(--radius-sm);font-size:.9em;line-height:1}
.recx:hover{color:#e0655f}
.clearhist{background:none;border:none;color:var(--text-dim);cursor:pointer;font-size:.82em;text-decoration:underline;padding:4px 8px}
.modes{display:flex;flex-direction:column;gap:8px;margin-top:8px}
.mode{display:flex;align-items:flex-start;gap:10px;background:var(--card);border:1px solid var(--border);border-radius:var(--radius-sm);padding:10px 12px;cursor:pointer}
.mode:hover{border-color:var(--accent-light)}
.mode.sel{border-color:var(--accent);background:var(--accent-glow)}
.mode input{margin-top:3px}
.mode .mname{font-weight:600;font-size:.9em}
.mode .mmeta{font-size:.78em;color:var(--text-dim)}
.actions{margin-top:18px;display:flex;gap:12px}
.btn{background:var(--accent);color:#fff;border:none;padding:10px 20px;border-radius:var(--radius-sm);cursor:pointer;font-size:.92em;font-weight:600;font-family:inherit;transition:background .15s;white-space:nowrap}
.btn:hover{background:var(--accent-light)}
.btn:disabled{opacity:.5;cursor:not-allowed}
.btn-ghost{background:var(--card);color:var(--text-muted);border:1px solid var(--border)}
.btn-stop{background:var(--card);color:var(--danger);border:1px solid var(--danger)}
.btn-stop:hover{background:var(--danger);color:#fff}
.btn-sort{font-size:.85em;white-space:nowrap}
.err{margin-top:12px;color:var(--danger);font-size:.85em;display:none}
.err.on{display:block}
.setfoot{margin-top:16px;font-size:.75em;color:var(--text-dim);text-align:center}
.setfoot a{color:var(--accent-light);text-decoration:none}
.setgrid{display:grid;grid-template-columns:1fr 1fr;gap:26px;align-items:start}
.setcol>*:first-child{margin-top:0}
.langsel{margin-left:auto;width:auto;min-width:132px;padding:6px 10px;margin-top:0;font-size:.82em}
@media(max-width:680px){.setgrid{grid-template-columns:1fr;gap:0}.setcard{width:min(560px,96vw)}}
    </style>
</head>
<body>
<div class="app">
  <header class="appbar">
    <div class="brand">__LOGO__<span class="bname">PS<i>-</i>NCDU</span><span class="ver">v__VERSION__</span></div>
    <nav class="bc" id="tbc"></nav>
    <div class="appact">
      <button class="btn btn-ghost" id="scangrey" onclick="scanAllGrey()" style="display:none"></button>
      <span class="ttotal" id="ttotal"></span>
      <span class="pctbadge" id="dpct"></span>
      <button class="btn btn-ghost btn-sort" id="sortbtn" onclick="toggleSort()"></button>
      <button class="iconbtn" data-i18n-title="theme" title="Theme clair / sombre" onclick="toggleTheme()">&#9681;</button>
      <button class="btn btn-stop" id="stopscan" onclick="stopScan()" style="display:none" data-i18n="stop">Interrompre</button>
      <button class="btn" id="dback" onclick="showSettings()" data-i18n="newscan">Nouveau scan</button>
    </div>
  </header>
  <div class="progline"><div class="gbar" id="gbar"><i id="gbarfill"></i><span class="gsweep"></span></div></div>
  <main class="treearea">
    <div class="tlist" id="tlist"><div class="tempty">Configurez une analyse pour commencer.</div></div>
    <div class="derr" id="derr"></div>
  </main>
  <footer class="statusbar">
    <span class="stmain" id="dscan"></span>
    <span class="ststage" id="dcur" data-i18n="ready">Pret.</span>
    <span class="stdepth" id="stdepth"></span>
    <span class="stelapsed">&#9201; <strong id="delapsed">0s</strong></span>
    <button class="stbtn" onclick="toggleLegend()" data-i18n="legend">Legende</button>
    <span id="dopwrap" style="display:none"></span>
    <span id="dpath" style="display:none"></span>
  </footer>
  <div class="legend" id="legend">
    <div class="legtabs">
      <button class="legtab on" id="legtab-leg" onclick="legTab('leg')" data-i18n="legend">Legende</button>
      <button class="legtab" id="legtab-aide" onclick="legTab('aide')" data-i18n="help">Aide</button>
      <button class="legclose" onclick="toggleLegend()" title="Fermer">&#10005;</button>
    </div>
    <div class="legpane on" id="legpane-leg">
      <div class="legrow"><span class="ic">&#128193;<span class="sdot sdot-green"></span></span> <span data-i18n="leggreen">Dossier scanne, taille connue</span></div>
      <div class="legrow"><span class="ic">&#128193;<span class="sdot sdot-orange"></span></span> <span data-i18n="legorange">Dossier prevu au scan</span></div>
      <div class="legrow"><span class="ic">&#128193;<span class="sdot sdot-blue"></span></span> <span data-i18n="legblue">Dossier en file</span></div>
      <div class="legrow"><span class="spin"></span> <span data-i18n="legspin">Dossier en cours de scan</span></div>
      <div class="legrow"><span class="ic">&#128193;<span class="sdot sdot-grey"></span></span> <span data-i18n="leggrey">Non scanne, exclu, jonction ou protege</span></div>
      <div class="legrow"><span class="fic">&#128196;</span> <span data-i18n="legfile">Fichier</span></div>
      <div class="legrow"><span class="legbar"><span></span></span> <span data-i18n="legbar">Part de la taille du dossier courant</span></div>
    </div>
    <div class="legpane" id="legpane-aide">
      <div class="legrow2" data-i18n="helpnav">Naviguer</div>
      <div class="legrow2" data-i18n="helpgrey">Dossiers gris</div>
      <div class="legrow2" data-i18n="helpfiles">Fichiers</div>
      <div class="legrow2" data-i18n="helpfilter">Filtre</div>
      <div class="legrow2" data-i18n="helpone">Un seul scan a la fois</div>
    </div>
  </div>
</div>

<div class="overlay on" id="overlay">
  <div class="setcard">
    <div class="sethead">__LOGO__<div><div class="settitle">PS-NCDU</div><div class="setsub" data-i18n="sub">Disk Usage Analyzer</div></div>
      <select id="langsel" class="langsel" onchange="setLang(this.value)" aria-label="Language" title="Language">
        <option value="en">English</option>
        <option value="fr">Francais</option>
        <option value="es">Espanol</option>
        <option value="de">Deutsch</option>
        <option value="pt">Portugues</option>
        <option value="ru">Русский</option>
        <option value="zh">中文</option>
        <option value="ar">العربية</option>
        <option value="hi">हिन्दी</option>
        <option value="bn">বাংলা</option>
        <option value="ur">اردو</option>
        <option value="id">Bahasa Indonesia</option>
        <option value="ja">日本語</option>
        <option value="ko">한국어</option>
        <option value="it">Italiano</option>
        <option value="tr">Türkçe</option>
        <option value="vi">Tiếng Việt</option>
        <option value="pl">Polski</option>
        <option value="nl">Nederlands</option>
        <option value="uk">Українська</option>
        <option value="ro">Română</option>
        <option value="cs">Čeština</option>
        <option value="el">Ελληνικά</option>
        <option value="sv">Svenska</option>
        <option value="hu">Magyar</option>
        <option value="fa">فارسی</option>
        <option value="th">ไทย</option>
        <option value="ms">Bahasa Melayu</option>
        <option value="fil">Filipino</option>
        <option value="sw">Kiswahili</option>
        <option value="ta">தமிழ்</option>
        <option value="te">తెలుగు</option>
        <option value="mr">मराठी</option>
        <option value="gu">ગુજરાતી</option>
        <option value="kn">ಕನ್ನಡ</option>
        <option value="ml">മലയാളം</option>
        <option value="pa">ਪੰਜਾਬੀ</option>
        <option value="he">עברית</option>
        <option value="ha">Hausa</option>
        <option value="my">မြန်မာ</option>
        <option value="am">አማርኛ</option>
        <option value="km">ខ្មែរ</option>
      </select>
      <span class="chip">v__VERSION__</span>
      <button class="setclose" id="setclose" onclick="closeModal()" aria-label="Fermer" title="Fermer" style="display:none">&#215;</button>
    </div>
    <h2 data-i18n="newanalysis">Nouvelle analyse</h2>
    <div class="hint" data-i18n="modalhint">Choisissez le dossier a analyser.</div>
    <div class="setgrid">
      <div class="setcol">
        <label data-i18n="folderlabel">Dossier a analyser</label>
        <div class="pathrow">
          <input id="path" type="text" value="__DEFAULTPATH__" placeholder="C:\ ou \\serveur\partage" spellcheck="false" oninput="onPathInput()" onkeydown="if(event.key==='Enter'){event.preventDefault();startScan();}">
          <span id="pathok" class="pathok" aria-hidden="true"></span>
          <button class="btn btn-ghost" id="pickbtn" onclick="openBrowse()" data-i18n="browse">Parcourir...</button>
        </div>
        <div class="quickwrap" id="usersWrap" style="display:none"><div class="quicklbl" data-i18n="quickaccess">Acces rapide</div><div class="chips" id="users"></div></div>
        <div class="quickwrap" id="histWrap" style="display:none"><div class="quicklbl" data-i18n="recent">Recents</div><div class="chips" id="history"></div></div>
        <div class="quickwrap"><div class="quicklbl" data-i18n="drives">Lecteurs</div><div class="chips" id="drives"></div></div>
      </div>
      <div class="setcol">
        <label><span data-i18n="depthlabel">Profondeur d exploration :</span> <b id="depthval">3</b> <span data-i18n="levels">niveau(x)</span></label>
        <input id="depth" type="range" min="1" max="10" value="3" oninput="onDepth()">
        <label class="chkline"><input type="checkbox" id="depthUnl" onchange="onDepth()"> <span data-i18n="unlimited">Illimitee</span></label>
        <div class="explain" id="depthExp"></div>
        <label data-i18n="displaylabel">Affichage : masquer les petits elements</label>
        <select id="minsize" onchange="onMinSize()">
          <option value="0" data-i18n="showall">Tout afficher</option>
          <option value="1048576" data-i18n="hide1">Masquer sous 1 Mo</option>
          <option value="104857600" data-i18n="hide100">Masquer sous 100 Mo</option>
          <option value="1073741824" data-i18n="hide1g">Masquer sous 1 Go</option>
        </select>
        <div class="explain" data-i18n="filterexp">Filtre.</div>
        <div class="exclwrap">
          <button type="button" class="excltoggle" onclick="toggleExcl()"><span class="chev" id="exclchev">&#9662;</span> <span data-i18n="exclusions">Exclusions</span></button>
          <div class="exclbody" id="exclbody" style="display:none">
            <div class="explain" data-i18n="exclexp">Ces dossiers systeme sont toujours ignores.</div>
            <ul class="excllist" id="excllist"></ul>
            <label data-i18n="exclcustom">Exclure aussi (un chemin par ligne) :</label>
            <textarea id="exclcustom" rows="2" spellcheck="false" placeholder="D:\\Temp"></textarea>
          </div>
        </div>
      </div>
    </div>
    <div class="actions">
      <button class="btn" id="go" onclick="startScan()" data-i18n="analyze">Analyser</button>
      <button class="btn btn-ghost" onclick="quitServer()" data-i18n="quit">Quitter le serveur</button>
    </div>
    <div class="err" id="err"></div>
    <div class="setfoot">Eric Guiffault &middot; <a href="mailto:__EMAIL__">__EMAIL__</a> &middot; PS-NCDU v__VERSION__ &middot; <span data-i18n="genby">Genere par :</span> __MODEL__</div>
  </div>
</div>
<div class="overlay" id="browseOverlay">
  <div class="browsecard">
    <div class="browsehead"><b data-i18n="browsetitle">Parcourir</b><button class="setclose" onclick="closeBrowse()" aria-label="Fermer" title="Fermer">&#215;</button></div>
    <div class="browsepath" id="browsePath"></div>
    <div class="browselist" id="browseList"></div>
    <div class="browseact"><button class="btn" onclick="chooseBrowse()" data-i18n="choose">Choisir ce dossier</button><button class="btn btn-ghost" onclick="closeBrowse()" data-i18n="cancel">Annuler</button></div>
  </div>
</div>
<script>
var TOKEN="__TOKEN__";
var LANG="__LANG__";
var I18N={
en:{sub:"Disk Usage Analyzer",newscan:"New scan",noanalysis:"No analysis",theme:"Light / dark theme",legend:"Legend",help:"Help",ready:"Ready.",launching:"Starting scan...",done:"Scan complete. Browse freely.",inprogress:"In progress:",total:"Total:",scanningtag:"scanning",parent:".. (parent folder)",scandots:"Scanning...",empty:"Empty folder",loadingfiles:"Loading files...",excluded:"excluded",junction:"junction",unscanned:"not scanned",protected:"protected",error:"error",scanshort:"scan...",queued:"queued",clickscan:"Click to scan this folder (2 levels)",filescap:"Showing the 1000 largest files of this folder.",filesafter:"Files will appear once the current scan finishes.",newanalysis:"New analysis",modalhint:"Choose the folder to analyze, then set the depth and display.",folderlabel:"Folder to analyze",browse:"Browse...",quickaccess:"Quick access",recent:"Recent",drives:"Drives",depthlabel:"Exploration depth:",levels:"level(s)",unlimited:"Unlimited (scans everything, can be very long)",displaylabel:"Display: hide small items",showall:"Show all",hide1:"Hide under 1 MB",hide100:"Hide under 100 MB",hide1g:"Hide under 1 GB",filterexp:"Filters the display only, for readability. Does not speed up the scan: sizes are always fully computed. Can be changed anytime.",analyze:"Analyze",quit:"Quit server",genby:"Generated by:",depthunl:"Scans and shows the whole tree. Can be very long and heavy on a large disk.",depthlow:"Light and fast to display: only the first levels are preloaded. Go deeper by clicking a folder.",depthmid:"Good balance: several levels visible at once, still smooth.",depthhigh:"Detailed: many levels preloaded, heavier to display.",depthtail:"Sizes are always exact; beyond this depth, click a grey folder to explore it.",leggreen:"Scanned folder, size known",legorange:"Folder planned in the current scan",legblue:"Folder queued (scanned when the current scan ends)",legspin:"Folder being scanned",leggrey:"Not scanned, excluded, junction or protected. Click to scan it (2 levels)",legfile:"File",legbar:"Share of the current folder size",helpnav:"Navigate: click a folder to enter, use the breadcrumb on top to go up.",helpgrey:"Grey folders: click to scan them (2 more levels). During a scan they queue (blue) and are scanned at the end.",helpfiles:"Files: shown when you open a folder, sorted by size, limited to the 1000 largest.",helpfilter:"Display filter (New scan button): hides small items for readability, without changing the scan.",helpone:"One scan at a time: two disk scans in parallel would slow each other down.",queuedmsg:"Folder queued ({n} waiting). Scanned when the current scan ends.",hiddenmsg:"{n} item(s) below the threshold hidden by the display filter.",invalidpath:"Invalid or inaccessible path.",windowopen:"Window open...",interrupted:"Scan interrupted.",enterpath:"Enter a path.",serverstopped:"Server stopped. You can close this tab.",scangrey:"Scan grey folders",you:"(you)",network:"(network)",drivefree:"free",removeone:"Remove",clearhist:"Clear history",exclusions:"Exclusions",exclexp:"These system folders are always skipped during the scan.",exclcustom:"Also exclude (one path per line):",scandepthfull:"Depth: unlimited",scandepthn:"Depth: {n}",stop:"Stop",pickerfail:"The native folder picker could not open on this machine. Type the path directly.",browsetitle:"Browse",choose:"Choose this folder",cancel:"Cancel",upfolder:"Parent folder",sortby:"Sort",byname:"name",bysize:"size"},
fr:{sub:"Analyseur d'espace disque",newscan:"Nouveau scan",noanalysis:"Aucune analyse",theme:"Thème clair / sombre",legend:"Légende",help:"Aide",ready:"Prêt.",launching:"Lancement du scan...",done:"Scan terminé. Naviguez librement.",inprogress:"En cours :",total:"Total :",scanningtag:"scan en cours",parent:".. (dossier parent)",scandots:"Scan en cours...",empty:"Dossier vide",loadingfiles:"Chargement des fichiers...",excluded:"exclu",junction:"jonction",unscanned:"non scanné",protected:"protégé",error:"erreur",scanshort:"scan...",queued:"en file",clickscan:"Cliquer pour scanner ce dossier (2 niveaux)",filescap:"Affichage limité aux 1000 plus gros fichiers de ce dossier.",filesafter:"Les fichiers s'afficheront à la fin du scan en cours.",newanalysis:"Nouvelle analyse",modalhint:"Choisissez le dossier à analyser, puis réglez la profondeur et l'affichage.",folderlabel:"Dossier à analyser",browse:"Parcourir...",quickaccess:"Accès rapide",recent:"Récents",drives:"Lecteurs",depthlabel:"Profondeur d'exploration :",levels:"niveau(x)",unlimited:"Illimitée (parcourt tout, peut être très long)",displaylabel:"Affichage : masquer les petits éléments",showall:"Tout afficher",hide1:"Masquer sous 1 Mo",hide100:"Masquer sous 100 Mo",hide1g:"Masquer sous 1 Go",filterexp:"Filtre uniquement l'affichage, pour la lisibilité. N'accélère pas le scan : les tailles sont toujours calculées en entier. Modifiable à tout moment.",analyze:"Analyser",quit:"Quitter le serveur",genby:"Généré par :",depthunl:"Parcourt et affiche toute l'arborescence. Peut être très long et lourd sur un gros disque.",depthlow:"Léger et rapide à afficher : seuls les premiers niveaux sont chargés. Descendez en cliquant un dossier.",depthmid:"Bon compromis : plusieurs niveaux visibles d'emblée, affichage fluide.",depthhigh:"Détaillé : beaucoup de niveaux chargés d'avance, plus lourd à afficher.",depthtail:"Les tailles sont toujours exactes ; au-delà, cliquez un dossier gris pour l'explorer.",leggreen:"Dossier scanné, taille connue",legorange:"Dossier prévu dans le scan en cours",legblue:"Dossier en file (scanné à la fin du scan en cours)",legspin:"Dossier en cours de scan",leggrey:"Non scanné, exclu, jonction ou protégé. Cliquer pour le scanner (2 niveaux)",legfile:"Fichier",legbar:"Part de la taille du dossier courant",helpnav:"Naviguer : cliquez un dossier pour entrer, le fil d'Ariane en haut pour remonter.",helpgrey:"Dossiers gris : cliquez pour les scanner (2 niveaux de plus). Pendant un scan ils passent en file (bleu) et sont scannés à la fin.",helpfiles:"Fichiers : affichés quand vous ouvrez un dossier, triés par taille, limités aux 1000 plus gros.",helpfilter:"Filtre d'affichage (bouton Nouveau scan) : masque les petits éléments pour la lisibilité, sans changer le scan.",helpone:"Un seul scan à la fois : deux scans disque en parallèle se ralentiraient.",queuedmsg:"Dossier mis en file ({n} en attente). Scanné à la fin du scan en cours.",hiddenmsg:"{n} élément(s) sous le seuil masqué(s) par le filtre d'affichage.",invalidpath:"Chemin invalide ou inaccessible.",windowopen:"Fenêtre ouverte...",interrupted:"Scan interrompu.",enterpath:"Indiquez un chemin.",serverstopped:"Serveur arrêté. Vous pouvez fermer cet onglet.",scangrey:"Scanner les dossiers gris",you:"(vous)",network:"(réseau)",drivefree:"libres",removeone:"Retirer",clearhist:"Effacer l'historique",exclusions:"Exclusions",exclexp:"Ces dossiers système sont toujours ignorés lors du scan.",exclcustom:"Exclure aussi (un chemin par ligne) :",scandepthfull:"Profondeur : illimitée",scandepthn:"Profondeur : {n}",stop:"Interrompre",pickerfail:"Le sélecteur de dossier natif n'a pas pu s'ouvrir sur ce poste. Saisissez le chemin directement.",browsetitle:"Parcourir",choose:"Choisir ce dossier",cancel:"Annuler",upfolder:"Dossier parent",sortby:"Tri",byname:"nom",bysize:"taille"},
es:{sub:"Analizador de uso de disco",newscan:"Nuevo escaneo",noanalysis:"Sin analisis",theme:"Tema claro / oscuro",legend:"Leyenda",help:"Ayuda",ready:"Listo.",launching:"Iniciando escaneo...",done:"Escaneo completo. Navegue libremente.",inprogress:"En curso:",total:"Total:",scanningtag:"escaneando",parent:".. (carpeta superior)",scandots:"Escaneando...",empty:"Carpeta vacia",loadingfiles:"Cargando archivos...",excluded:"excluido",junction:"union",unscanned:"sin escanear",protected:"protegido",error:"error",scanshort:"esc...",queued:"en cola",clickscan:"Clic para escanear esta carpeta (2 niveles)",filescap:"Mostrando los 1000 archivos mas grandes de esta carpeta.",filesafter:"Los archivos apareceran al terminar el escaneo actual.",newanalysis:"Nuevo analisis",modalhint:"Elija la carpeta a analizar, luego ajuste la profundidad y la visualizacion.",folderlabel:"Carpeta a analizar",browse:"Explorar...",quickaccess:"Acceso rapido",recent:"Recientes",drives:"Unidades",depthlabel:"Profundidad de exploracion:",levels:"nivel(es)",unlimited:"Ilimitada (recorre todo, puede ser muy largo)",displaylabel:"Vista: ocultar elementos pequenos",showall:"Mostrar todo",hide1:"Ocultar bajo 1 MB",hide100:"Ocultar bajo 100 MB",hide1g:"Ocultar bajo 1 GB",filterexp:"Filtra solo la vista, para la legibilidad. No acelera el escaneo: los tamanos siempre se calculan por completo. Modificable en cualquier momento.",analyze:"Analizar",quit:"Detener servidor",genby:"Generado por:",depthunl:"Recorre y muestra todo el arbol. Puede ser muy largo y pesado en un disco grande.",depthlow:"Ligero y rapido: solo los primeros niveles se precargan. Baje haciendo clic en una carpeta.",depthmid:"Buen equilibrio: varios niveles visibles a la vez, fluido.",depthhigh:"Detallado: muchos niveles precargados, mas pesado de mostrar.",depthtail:"Los tamanos siempre son exactos; mas alla, haga clic en una carpeta gris para explorarla.",leggreen:"Carpeta escaneada, tamano conocido",legorange:"Carpeta prevista en el escaneo actual",legblue:"Carpeta en cola (escaneada al terminar el escaneo actual)",legspin:"Carpeta en escaneo",leggrey:"Sin escanear, excluida, union o protegida. Clic para escanearla (2 niveles)",legfile:"Archivo",legbar:"Parte del tamano de la carpeta actual",helpnav:"Navegar: clic en una carpeta para entrar, la ruta de arriba para subir.",helpgrey:"Carpetas grises: clic para escanearlas (2 niveles mas). Durante un escaneo pasan a cola (azul) y se escanean al final.",helpfiles:"Archivos: se muestran al abrir una carpeta, ordenados por tamano, limitados a los 1000 mayores.",helpfilter:"Filtro de vista (boton Nuevo escaneo): oculta elementos pequenos, sin cambiar el escaneo.",helpone:"Un escaneo a la vez: dos escaneos de disco en paralelo se ralentizarian.",queuedmsg:"Carpeta en cola ({n} en espera). Se escanea al terminar el escaneo actual.",hiddenmsg:"{n} elemento(s) bajo el umbral ocultos por el filtro de vista.",invalidpath:"Ruta invalida o inaccesible.",windowopen:"Ventana abierta...",interrupted:"Escaneo interrumpido.",enterpath:"Indique una ruta.",serverstopped:"Servidor detenido. Puede cerrar esta pestana.",scangrey:"Escanear carpetas grises",you:"(usted)",network:"(red)",drivefree:"libres",removeone:"Quitar",clearhist:"Borrar historial",exclusions:"Exclusiones",exclexp:"Estas carpetas del sistema siempre se omiten durante el escaneo.",exclcustom:"Excluir tambien (una ruta por linea):",scandepthfull:"Profundidad: ilimitada",scandepthn:"Profundidad: {n}",stop:"Detener",pickerfail:"El selector de carpetas nativo no pudo abrirse en este equipo. Escribe la ruta directamente.",browsetitle:"Examinar",choose:"Elegir esta carpeta",cancel:"Cancelar",upfolder:"Carpeta superior",sortby:"Ordenar",byname:"nombre",bysize:"tamano"},
de:{sub:"Speicherplatz-Analyse",newscan:"Neuer Scan",noanalysis:"Keine Analyse",theme:"Helles / dunkles Thema",legend:"Legende",help:"Hilfe",ready:"Bereit.",launching:"Scan wird gestartet...",done:"Scan fertig. Frei navigieren.",inprogress:"Laeuft:",total:"Gesamt:",scanningtag:"wird gescannt",parent:".. (uebergeordneter Ordner)",scandots:"Scannen...",empty:"Leerer Ordner",loadingfiles:"Dateien werden geladen...",excluded:"ausgeschlossen",junction:"Verknuepfung",unscanned:"nicht gescannt",protected:"geschuetzt",error:"Fehler",scanshort:"scan...",queued:"in Warteschlange",clickscan:"Klicken, um diesen Ordner zu scannen (2 Ebenen)",filescap:"Zeigt die 1000 groessten Dateien dieses Ordners.",filesafter:"Dateien erscheinen nach dem aktuellen Scan.",newanalysis:"Neue Analyse",modalhint:"Waehlen Sie den Ordner, dann Tiefe und Anzeige einstellen.",folderlabel:"Zu analysierender Ordner",browse:"Durchsuchen...",quickaccess:"Schnellzugriff",recent:"Zuletzt",drives:"Laufwerke",depthlabel:"Erkundungstiefe:",levels:"Ebene(n)",unlimited:"Unbegrenzt (durchsucht alles, kann sehr lange dauern)",displaylabel:"Anzeige: kleine Elemente ausblenden",showall:"Alle anzeigen",hide1:"Unter 1 MB ausblenden",hide100:"Unter 100 MB ausblenden",hide1g:"Unter 1 GB ausblenden",filterexp:"Filtert nur die Anzeige, fuer die Lesbarkeit. Beschleunigt den Scan nicht: Groessen werden immer voll berechnet. Jederzeit aenderbar.",analyze:"Analysieren",quit:"Server beenden",genby:"Erstellt von:",depthunl:"Durchsucht und zeigt den ganzen Baum. Kann sehr lang und schwer sein bei grossem Datentraeger.",depthlow:"Leicht und schnell: nur die ersten Ebenen werden vorgeladen. Tiefer per Klick auf einen Ordner.",depthmid:"Gute Balance: mehrere Ebenen auf einmal sichtbar, fluessig.",depthhigh:"Detailliert: viele Ebenen vorgeladen, schwerer anzuzeigen.",depthtail:"Groessen sind immer exakt; darueber hinaus einen grauen Ordner anklicken.",leggreen:"Gescannter Ordner, Groesse bekannt",legorange:"Ordner im aktuellen Scan geplant",legblue:"Ordner in Warteschlange (nach dem aktuellen Scan)",legspin:"Ordner wird gescannt",leggrey:"Nicht gescannt, ausgeschlossen, Verknuepfung oder geschuetzt. Klicken zum Scannen (2 Ebenen)",legfile:"Datei",legbar:"Anteil an der Groesse des aktuellen Ordners",helpnav:"Navigation: Ordner anklicken zum Oeffnen, Brotkrumen oben zum Hochgehen.",helpgrey:"Graue Ordner: anklicken zum Scannen (2 weitere Ebenen). Waehrend eines Scans in Warteschlange (blau), am Ende gescannt.",helpfiles:"Dateien: erscheinen beim Oeffnen eines Ordners, nach Groesse sortiert, auf die 1000 groessten begrenzt.",helpfilter:"Anzeigefilter (Knopf Neuer Scan): blendet kleine Elemente aus, ohne den Scan zu aendern.",helpone:"Ein Scan gleichzeitig: zwei parallele Datentraeger-Scans wuerden sich bremsen.",queuedmsg:"Ordner in Warteschlange ({n} wartend). Nach dem aktuellen Scan gescannt.",hiddenmsg:"{n} Element(e) unter dem Schwellwert vom Anzeigefilter ausgeblendet.",invalidpath:"Ungueltiger oder unzugaenglicher Pfad.",windowopen:"Fenster geoeffnet...",interrupted:"Scan unterbrochen.",enterpath:"Bitte einen Pfad angeben.",serverstopped:"Server gestoppt. Sie koennen diesen Tab schliessen.",scangrey:"Graue Ordner scannen",you:"(Sie)",network:"(Netzwerk)",drivefree:"frei",removeone:"Entfernen",clearhist:"Verlauf loeschen",exclusions:"Ausschluesse",exclexp:"Diese Systemordner werden beim Scan immer uebersprungen.",exclcustom:"Zusaetzlich ausschliessen (ein Pfad pro Zeile):",scandepthfull:"Tiefe: unbegrenzt",scandepthn:"Tiefe: {n}",stop:"Abbrechen",pickerfail:"Die native Ordnerauswahl konnte auf diesem Rechner nicht geoeffnet werden. Geben Sie den Pfad direkt ein.",browsetitle:"Durchsuchen",choose:"Diesen Ordner waehlen",cancel:"Abbrechen",upfolder:"Uebergeordneter Ordner",sortby:"Sortieren",byname:"Name",bysize:"Groesse"},
pt:{sub:"Analisador de uso de disco",newscan:"Nova varredura",noanalysis:"Sem analise",theme:"Tema claro / escuro",legend:"Legenda",help:"Ajuda",ready:"Pronto.",launching:"Iniciando varredura...",done:"Varredura concluida. Navegue livremente.",inprogress:"Em curso:",total:"Total:",scanningtag:"varrendo",parent:".. (pasta superior)",scandots:"Varrendo...",empty:"Pasta vazia",loadingfiles:"Carregando arquivos...",excluded:"excluido",junction:"juncao",unscanned:"nao varrido",protected:"protegido",error:"erro",scanshort:"var...",queued:"na fila",clickscan:"Clique para varrer esta pasta (2 niveis)",filescap:"Mostrando os 1000 maiores arquivos desta pasta.",filesafter:"Os arquivos aparecerao ao fim da varredura atual.",newanalysis:"Nova analise",modalhint:"Escolha a pasta a analisar, depois ajuste a profundidade e a exibicao.",folderlabel:"Pasta a analisar",browse:"Procurar...",quickaccess:"Acesso rapido",recent:"Recentes",drives:"Unidades",depthlabel:"Profundidade de exploracao:",levels:"nivel(is)",unlimited:"Ilimitada (percorre tudo, pode ser muito longo)",displaylabel:"Exibicao: ocultar itens pequenos",showall:"Mostrar tudo",hide1:"Ocultar abaixo de 1 MB",hide100:"Ocultar abaixo de 100 MB",hide1g:"Ocultar abaixo de 1 GB",filterexp:"Filtra apenas a exibicao, para a legibilidade. Nao acelera a varredura: os tamanhos sao sempre calculados por completo. Alteravel a qualquer momento.",analyze:"Analisar",quit:"Encerrar servidor",genby:"Gerado por:",depthunl:"Percorre e mostra toda a arvore. Pode ser muito longo e pesado num disco grande.",depthlow:"Leve e rapido: so os primeiros niveis sao pre-carregados. Desca clicando numa pasta.",depthmid:"Bom equilibrio: varios niveis visiveis de uma vez, fluido.",depthhigh:"Detalhado: muitos niveis pre-carregados, mais pesado de exibir.",depthtail:"Os tamanhos sao sempre exatos; alem disso, clique numa pasta cinza para explora-la.",leggreen:"Pasta varrida, tamanho conhecido",legorange:"Pasta prevista na varredura atual",legblue:"Pasta na fila (varrida ao fim da varredura atual)",legspin:"Pasta em varredura",leggrey:"Nao varrida, excluida, juncao ou protegida. Clique para varre-la (2 niveis)",legfile:"Arquivo",legbar:"Parte do tamanho da pasta atual",helpnav:"Navegar: clique numa pasta para entrar, a trilha no topo para subir.",helpgrey:"Pastas cinzas: clique para varre-las (2 niveis a mais). Durante uma varredura ficam na fila (azul) e sao varridas no fim.",helpfiles:"Arquivos: aparecem ao abrir uma pasta, ordenados por tamanho, limitados aos 1000 maiores.",helpfilter:"Filtro de exibicao (botao Nova varredura): oculta itens pequenos, sem mudar a varredura.",helpone:"Uma varredura por vez: duas varreduras de disco em paralelo se atrasariam.",queuedmsg:"Pasta na fila ({n} aguardando). Varrida ao fim da varredura atual.",hiddenmsg:"{n} item(ns) abaixo do limite ocultos pelo filtro de exibicao.",invalidpath:"Caminho invalido ou inacessivel.",windowopen:"Janela aberta...",interrupted:"Varredura interrompida.",enterpath:"Indique um caminho.",serverstopped:"Servidor parado. Voce pode fechar esta aba.",scangrey:"Varrer pastas cinzas",you:"(voce)",network:"(rede)",drivefree:"livres",removeone:"Remover",clearhist:"Limpar historico",exclusions:"Exclusoes",exclexp:"Estas pastas do sistema sao sempre ignoradas durante a analise.",exclcustom:"Excluir tambem (um caminho por linha):",scandepthfull:"Profundidade: ilimitada",scandepthn:"Profundidade: {n}",stop:"Parar",pickerfail:"O seletor de pastas nativo nao pode abrir nesta maquina. Digite o caminho diretamente.",browsetitle:"Procurar",choose:"Escolher esta pasta",cancel:"Cancelar",upfolder:"Pasta superior",sortby:"Ordenar",byname:"nome",bysize:"tamanho"},
ru:{sub:"Анализатор дискового пространства",newscan:"Новое сканирование",noanalysis:"Нет анализа",theme:"Светлая / темная тема",legend:"Легенда",help:"Справка",ready:"Готово.",launching:"Запуск сканирования...",done:"Сканирование завершено. Просматривайте свободно.",inprogress:"Выполняется:",total:"Всего:",scanningtag:"сканируется",parent:".. (родительская папка)",scandots:"Сканирование...",empty:"Пустая папка",loadingfiles:"Загрузка файлов...",excluded:"исключено",junction:"соединение",unscanned:"не сканировано",protected:"защищено",error:"ошибка",scanshort:"скан...",queued:"в очереди",clickscan:"Нажмите, чтобы сканировать эту папку (2 уровня)",filescap:"Показаны 1000 крупнейших файлов этой папки.",filesafter:"Файлы появятся после текущего сканирования.",newanalysis:"Новый анализ",modalhint:"Выберите папку, затем задайте глубину и отображение.",folderlabel:"Папка для анализа",browse:"Обзор...",quickaccess:"Быстрый доступ",recent:"Недавние",drives:"Диски",depthlabel:"Глубина обхода:",levels:"уровень(и)",unlimited:"Без ограничения (обходит все, может быть очень долго)",displaylabel:"Показ: скрыть мелкие элементы",showall:"Показать все",hide1:"Скрыть меньше 1 МБ",hide100:"Скрыть меньше 100 МБ",hide1g:"Скрыть меньше 1 ГБ",filterexp:"Фильтрует только отображение, для читаемости. Не ускоряет сканирование: размеры всегда считаются полностью. Можно менять в любой момент.",analyze:"Анализировать",quit:"Остановить сервер",genby:"Создано:",depthunl:"Обходит и показывает все дерево. Может быть очень долго и тяжело на большом диске.",depthlow:"Легко и быстро: загружаются только первые уровни. Глубже по клику на папку.",depthmid:"Хороший баланс: сразу видно несколько уровней, плавно.",depthhigh:"Подробно: много уровней загружено заранее, тяжелее для показа.",depthtail:"Размеры всегда точны; глубже нажмите серую папку, чтобы ее раскрыть.",leggreen:"Папка просканирована, размер известен",legorange:"Папка запланирована в текущем сканировании",legblue:"Папка в очереди (сканируется в конце текущего)",legspin:"Папка сканируется",leggrey:"Не сканирована, исключена, соединение или защищена. Нажмите для сканирования (2 уровня)",legfile:"Файл",legbar:"Доля от размера текущей папки",helpnav:"Навигация: клик по папке для входа, путь сверху для возврата.",helpgrey:"Серые папки: клик для сканирования (еще 2 уровня). Во время сканирования встают в очередь (синие) и сканируются в конце.",helpfiles:"Файлы: показываются при открытии папки, по размеру, до 1000 крупнейших.",helpfilter:"Фильтр показа (кнопка Новое сканирование): скрывает мелкие элементы, не меняя сканирование.",helpone:"Одно сканирование за раз: два параллельных замедлили бы друг друга.",queuedmsg:"Папка в очереди ({n} ожидает). Сканируется в конце текущего.",hiddenmsg:"{n} элемент(ов) ниже порога скрыто фильтром показа.",invalidpath:"Неверный или недоступный путь.",windowopen:"Окно открыто...",interrupted:"Сканирование прервано.",enterpath:"Укажите путь.",serverstopped:"Сервер остановлен. Можно закрыть вкладку.",scangrey:"Сканировать серые папки",you:"(вы)",network:"(сеть)",drivefree:"свободно",removeone:"Удалить",clearhist:"Очистить историю",exclusions:"Исключения",exclexp:"Эти системные папки всегда пропускаются при сканировании.",exclcustom:"Исключить также (по одному пути в строке):",scandepthfull:"Глубина: без ограничений",scandepthn:"Глубина: {n}",stop:"Остановить",pickerfail:"Не удалось открыть системный выбор папки на этом компьютере. Введите путь вручную.",browsetitle:"Обзор",choose:"Выбрать эту папку",cancel:"Отмена",upfolder:"Родительская папка",sortby:"Сортировка",byname:"имя",bysize:"размер"},
zh:{sub:"磁盘占用分析器",newscan:"新扫描",noanalysis:"无分析",theme:"浅色 / 深色主题",legend:"图例",help:"帮助",ready:"就绪。",launching:"正在启动扫描...",done:"扫描完成。可自由浏览。",inprogress:"进行中：",total:"合计：",scanningtag:"扫描中",parent:".. (上级文件夹)",scandots:"扫描中...",empty:"空文件夹",loadingfiles:"正在加载文件...",excluded:"已排除",junction:"联接",unscanned:"未扫描",protected:"受保护",error:"错误",scanshort:"扫描...",queued:"排队中",clickscan:"点击扫描此文件夹（2 层）",filescap:"显示此文件夹中最大的 1000 个文件。",filesafter:"当前扫描结束后将显示文件。",newanalysis:"新分析",modalhint:"选择要分析的文件夹，然后设置深度与显示。",folderlabel:"要分析的文件夹",browse:"浏览...",quickaccess:"快速访问",recent:"最近",drives:"驱动器",depthlabel:"探索深度：",levels:"层",unlimited:"无限（遍历全部，可能很久）",displaylabel:"显示：隐藏小项目",showall:"全部显示",hide1:"隐藏小于 1 MB",hide100:"隐藏小于 100 MB",hide1g:"隐藏小于 1 GB",filterexp:"仅过滤显示，便于阅读。不会加快扫描：大小始终完整计算。可随时更改。",analyze:"分析",quit:"停止服务器",genby:"生成者：",depthunl:"遍历并显示整棵树。在大磁盘上可能很久很重。",depthlow:"轻快显示：仅预加载前几层。点击文件夹可深入。",depthmid:"良好平衡：一次可见多层，仍然流畅。",depthhigh:"详细：预加载很多层，显示更重。",depthtail:"大小始终精确；再深处请点击灰色文件夹以展开。",leggreen:"已扫描文件夹，大小已知",legorange:"当前扫描计划中的文件夹",legblue:"排队中的文件夹（当前扫描结束后扫描）",legspin:"正在扫描的文件夹",leggrey:"未扫描、已排除、联接或受保护。点击扫描（2 层）",legfile:"文件",legbar:"占当前文件夹大小的比例",helpnav:"导航：点击文件夹进入，用顶部面包屑返回。",helpgrey:"灰色文件夹：点击扫描（再 2 层）。扫描期间会排队（蓝色）并在结束时扫描。",helpfiles:"文件：打开文件夹时显示，按大小排序，最多 1000 个最大文件。",helpfilter:"显示过滤（新扫描按钮）：隐藏小项目，不改变扫描。",helpone:"一次一个扫描：两个磁盘扫描并行会互相拖慢。",queuedmsg:"文件夹已排队（{n} 个等待）。当前扫描结束后扫描。",hiddenmsg:"{n} 个低于阈值的项目被显示过滤器隐藏。",invalidpath:"路径无效或不可访问。",windowopen:"窗口已打开...",interrupted:"扫描已中断。",enterpath:"请输入路径。",serverstopped:"服务器已停止。可关闭此标签页。",scangrey:"扫描灰色文件夹",you:"（您）",network:"（网络）",drivefree:"可用",removeone:"移除",clearhist:"清除历史",exclusions:"排除项",exclexp:"这些系统文件夹在扫描时始终被跳过。",exclcustom:"另外排除（每行一个路径）：",scandepthfull:"深度：无限",scandepthn:"深度：{n}",stop:"中止",pickerfail:"本机无法打开原生文件夹选择器。请直接输入路径。",browsetitle:"浏览",choose:"选择此文件夹",cancel:"取消",upfolder:"上级文件夹",sortby:"排序",byname:"名称",bysize:"大小"},
ar:{sub:"محلل استخدام القرص",newscan:"فحص جديد",noanalysis:"لا يوجد تحليل",theme:"سمة فاتحة / داكنة",legend:"مفتاح",help:"مساعدة",ready:"جاهز.",launching:"جارٍ بدء الفحص...",done:"اكتمل الفحص. تصفح بحرية.",inprogress:"جارٍ:",total:"الإجمالي:",scanningtag:"جارٍ الفحص",parent:".. (المجلد الأصل)",scandots:"جارٍ الفحص...",empty:"مجلد فارغ",loadingfiles:"جارٍ تحميل الملفات...",excluded:"مستبعد",junction:"وصلة",unscanned:"غير مفحوص",protected:"محمي",error:"خطأ",scanshort:"فحص...",queued:"في الطابور",clickscan:"انقر لفحص هذا المجلد (مستويان)",filescap:"عرض أكبر 1000 ملف في هذا المجلد.",filesafter:"ستظهر الملفات عند انتهاء الفحص الحالي.",newanalysis:"تحليل جديد",modalhint:"اختر المجلد للتحليل، ثم اضبط العمق والعرض.",folderlabel:"المجلد للتحليل",browse:"استعراض...",quickaccess:"وصول سريع",recent:"الأخيرة",drives:"محركات الأقراص",depthlabel:"عمق الاستكشاف:",levels:"مستوى",unlimited:"غير محدود (يجتاز كل شيء، قد يكون طويلاً جداً)",displaylabel:"العرض: إخفاء العناصر الصغيرة",showall:"إظهار الكل",hide1:"إخفاء أقل من 1 م.ب",hide100:"إخفاء أقل من 100 م.ب",hide1g:"إخفاء أقل من 1 غ.ب",filterexp:"يفلتر العرض فقط، للوضوح. لا يسرّع الفحص: تُحسب الأحجام كاملة دائماً. قابل للتغيير في أي وقت.",analyze:"تحليل",quit:"إيقاف الخادم",genby:"أُنشئ بواسطة:",depthunl:"يجتاز ويعرض الشجرة كاملة. قد يكون طويلاً وثقيلاً على قرص كبير.",depthlow:"خفيف وسريع: تُحمّل المستويات الأولى فقط. انزل بالنقر على مجلد.",depthmid:"توازن جيد: عدة مستويات مرئية دفعة واحدة، سلس.",depthhigh:"مفصّل: مستويات كثيرة محمّلة مسبقاً، أثقل في العرض.",depthtail:"الأحجام دقيقة دائماً؛ أبعد من ذلك انقر مجلداً رمادياً لاستكشافه.",leggreen:"مجلد مفحوص، الحجم معروف",legorange:"مجلد مخطط في الفحص الحالي",legblue:"مجلد في الطابور (يُفحص عند انتهاء الفحص الحالي)",legspin:"مجلد قيد الفحص",leggrey:"غير مفحوص أو مستبعد أو وصلة أو محمي. انقر لفحصه (مستويان)",legfile:"ملف",legbar:"نسبة من حجم المجلد الحالي",helpnav:"التنقل: انقر مجلداً للدخول، ومسار التنقل بالأعلى للصعود.",helpgrey:"المجلدات الرمادية: انقر لفحصها (مستويان إضافيان). أثناء الفحص تدخل الطابور (أزرق) وتُفحص في النهاية.",helpfiles:"الملفات: تظهر عند فتح مجلد، مرتبة بالحجم، بحد أقصى 1000 ملف.",helpfilter:"فلتر العرض (زر فحص جديد): يخفي العناصر الصغيرة دون تغيير الفحص.",helpone:"فحص واحد في كل مرة: فحصان متوازيان للقرص يبطئان بعضهما.",queuedmsg:"المجلد في الطابور ({n} في الانتظار). يُفحص عند انتهاء الفحص الحالي.",hiddenmsg:"{n} عنصر تحت الحد مخفي بفلتر العرض.",invalidpath:"مسار غير صالح أو غير قابل للوصول.",windowopen:"النافذة مفتوحة...",interrupted:"توقف الفحص.",enterpath:"أدخل مساراً.",serverstopped:"تم إيقاف الخادم. يمكنك إغلاق هذا التبويب.",scangrey:"افحص المجلدات الرمادية",you:"(أنت)",network:"(شبكة)",drivefree:"متاح",removeone:"إزالة",clearhist:"مسح السجل",exclusions:"الاستثناءات",exclexp:"يتم دائمًا تخطي مجلدات النظام هذه أثناء الفحص.",exclcustom:"استبعاد أيضًا (مسار واحد لكل سطر):",scandepthfull:"العمق: غير محدود",scandepthn:"العمق: {n}",stop:"إيقاف",pickerfail:"تعذر فتح منتقي المجلدات الأصلي على هذا الجهاز. اكتب المسار مباشرة.",browsetitle:"استعراض",choose:"اختيار هذا المجلد",cancel:"إلغاء",upfolder:"المجلد الأصل",sortby:"ترتيب",byname:"الاسم",bysize:"الحجم"},
hi:{sub:"डिस्क उपयोग विश्लेषक",newscan:"नया स्कैन",noanalysis:"कोई विश्लेषण नहीं",theme:"हल्की / गहरी थीम",legend:"संकेत",help:"सहायता",ready:"तैयार।",launching:"स्कैन शुरू हो रहा है...",done:"स्कैन पूरा। स्वतंत्र रूप से देखें।",inprogress:"जारी:",total:"कुल:",scanningtag:"स्कैन जारी",parent:".. (मूल फ़ोल्डर)",scandots:"स्कैन जारी...",empty:"खाली फ़ोल्डर",loadingfiles:"फ़ाइलें लोड हो रही हैं...",excluded:"बाहर रखा",junction:"जंक्शन",unscanned:"अस्कैन",protected:"संरक्षित",error:"त्रुटि",scanshort:"स्कैन...",queued:"कतार में",clickscan:"इस फ़ोल्डर को स्कैन करने के लिए क्लिक करें (2 स्तर)",filescap:"इस फ़ोल्डर की 1000 सबसे बड़ी फ़ाइलें दिखा रहे हैं।",filesafter:"वर्तमान स्कैन समाप्त होने पर फ़ाइलें दिखेंगी।",newanalysis:"नया विश्लेषण",modalhint:"विश्लेषण हेतु फ़ोल्डर चुनें, फिर गहराई और प्रदर्शन सेट करें।",folderlabel:"विश्लेषण हेतु फ़ोल्डर",browse:"ब्राउज़ करें...",quickaccess:"त्वरित पहुँच",recent:"हाल के",drives:"ड्राइव",depthlabel:"अन्वेषण गहराई:",levels:"स्तर",unlimited:"असीमित (सब कुछ, बहुत लंबा हो सकता है)",displaylabel:"प्रदर्शन: छोटे आइटम छिपाएँ",showall:"सब दिखाएँ",hide1:"1 MB से कम छिपाएँ",hide100:"100 MB से कम छिपाएँ",hide1g:"1 GB से कम छिपाएँ",filterexp:"केवल प्रदर्शन फ़िल्टर करता है, पठनीयता हेतु। स्कैन तेज़ नहीं करता: आकार हमेशा पूरा गणना होते हैं। कभी भी बदल सकते हैं।",analyze:"विश्लेषण करें",quit:"सर्वर बंद करें",genby:"निर्मित:",depthunl:"पूरे वृक्ष को देखता और दिखाता है। बड़ी डिस्क पर बहुत लंबा और भारी हो सकता है।",depthlow:"हल्का और तेज़: केवल पहले स्तर पूर्वलोड होते हैं। फ़ोल्डर पर क्लिक कर नीचे जाएँ।",depthmid:"अच्छा संतुलन: एक साथ कई स्तर दिखते हैं, सहज।",depthhigh:"विस्तृत: कई स्तर पूर्वलोड, दिखाने में भारी।",depthtail:"आकार हमेशा सटीक; इससे आगे किसी धूसर फ़ोल्डर पर क्लिक करें।",leggreen:"स्कैन किया फ़ोल्डर, आकार ज्ञात",legorange:"वर्तमान स्कैन में नियोजित फ़ोल्डर",legblue:"कतार में फ़ोल्डर (वर्तमान स्कैन के अंत में)",legspin:"स्कैन हो रहा फ़ोल्डर",leggrey:"अस्कैन, बाहर, जंक्शन या संरक्षित। स्कैन हेतु क्लिक करें (2 स्तर)",legfile:"फ़ाइल",legbar:"वर्तमान फ़ोल्डर आकार का हिस्सा",helpnav:"नेविगेट: प्रवेश हेतु फ़ोल्डर पर क्लिक करें, ऊपर की पगडंडी से ऊपर जाएँ।",helpgrey:"धूसर फ़ोल्डर: स्कैन हेतु क्लिक (2 और स्तर)। स्कैन के दौरान कतार में (नीला) और अंत में स्कैन।",helpfiles:"फ़ाइलें: फ़ोल्डर खोलने पर दिखती हैं, आकार अनुसार, 1000 सबसे बड़ी तक।",helpfilter:"प्रदर्शन फ़िल्टर (नया स्कैन बटन): छोटे आइटम छिपाता है, स्कैन बदले बिना।",helpone:"एक बार में एक स्कैन: दो समानांतर डिस्क स्कैन एक-दूसरे को धीमा करेंगे।",queuedmsg:"फ़ोल्डर कतार में ({n} प्रतीक्षारत)। वर्तमान स्कैन के अंत में स्कैन।",hiddenmsg:"सीमा से नीचे {n} आइटम प्रदर्शन फ़िल्टर द्वारा छिपे।",invalidpath:"अमान्य या दुर्गम पथ।",windowopen:"विंडो खुली...",interrupted:"स्कैन बाधित।",enterpath:"एक पथ दर्ज करें।",serverstopped:"सर्वर रुका। आप यह टैब बंद कर सकते हैं।",scangrey:"धूसर फ़ोल्डर स्कैन करें",you:"(आप)",network:"(नेटवर्क)",drivefree:"खाली",removeone:"हटाएँ",clearhist:"इतिहास साफ़ करें",exclusions:"बहिष्करण",exclexp:"ये सिस्टम फ़ोल्डर स्कैन के दौरान हमेशा छोड़ दिए जाते हैं।",exclcustom:"अन्य बहिष्कृत (प्रति पंक्ति एक पथ):",scandepthfull:"गहराई: असीमित",scandepthn:"गहराई: {n}",stop:"रोकें",pickerfail:"इस मशीन पर मूल फ़ोल्डर चयनकर्ता नहीं खुल सका। पथ सीधे टाइप करें।",browsetitle:"ब्राउज़ करें",choose:"यह फ़ोल्डर चुनें",cancel:"रद्द करें",upfolder:"मूल फ़ोल्डर",sortby:"क्रमबद्ध",byname:"नाम",bysize:"आकार"},
bn:{sub:"ডিস্ক ব্যবহার বিশ্লেষক",newscan:"নতুন স্ক্যান",noanalysis:"কোনো বিশ্লেষণ নেই",theme:"হালকা / গাঢ় থিম",legend:"নির্দেশিকা",help:"সহায়তা",ready:"প্রস্তুত।",launching:"স্ক্যান শুরু হচ্ছে...",done:"স্ক্যান সম্পন্ন। স্বাধীনভাবে দেখুন।",inprogress:"চলছে:",total:"মোট:",scanningtag:"স্ক্যান চলছে",parent:".. (মূল ফোল্ডার)",scandots:"স্ক্যান চলছে...",empty:"খালি ফোল্ডার",loadingfiles:"ফাইল লোড হচ্ছে...",excluded:"বাদ",junction:"জংশন",unscanned:"অস্ক্যান",protected:"সুরক্ষিত",error:"ত্রুটি",scanshort:"স্ক্যান...",queued:"সারিতে",clickscan:"এই ফোল্ডার স্ক্যান করতে ক্লিক করুন (২ স্তর)",filescap:"এই ফোল্ডারের ১০০০ সবচেয়ে বড় ফাইল দেখানো হচ্ছে।",filesafter:"বর্তমান স্ক্যান শেষ হলে ফাইল দেখা যাবে।",newanalysis:"নতুন বিশ্লেষণ",modalhint:"বিশ্লেষণের ফোল্ডার বেছে নিন, তারপর গভীরতা ও প্রদর্শন সেট করুন।",folderlabel:"বিশ্লেষণের ফোল্ডার",browse:"ব্রাউজ...",quickaccess:"দ্রুত প্রবেশ",recent:"সাম্প্রতিক",drives:"ড্রাইভ",depthlabel:"অন্বেষণ গভীরতা:",levels:"স্তর",unlimited:"সীমাহীন (সব ঘুরে দেখে, খুব দীর্ঘ হতে পারে)",displaylabel:"প্রদর্শন: ছোট আইটেম লুকান",showall:"সব দেখান",hide1:"১ MB এর কম লুকান",hide100:"১০০ MB এর কম লুকান",hide1g:"১ GB এর কম লুকান",filterexp:"শুধু প্রদর্শন ফিল্টার করে, পঠনযোগ্যতার জন্য। স্ক্যান দ্রুত করে না: আকার সবসময় পুরো গণনা হয়। যেকোনো সময় পরিবর্তনযোগ্য।",analyze:"বিশ্লেষণ",quit:"সার্ভার বন্ধ",genby:"তৈরি করেছে:",depthunl:"পুরো গাছ ঘুরে দেখায়। বড় ডিস্কে খুব দীর্ঘ ও ভারী হতে পারে।",depthlow:"হালকা ও দ্রুত: শুধু প্রথম স্তরগুলো প্রিলোড হয়। ফোল্ডারে ক্লিক করে নিচে যান।",depthmid:"ভালো ভারসাম্য: একসাথে কয়েক স্তর দৃশ্যমান, মসৃণ।",depthhigh:"বিস্তারিত: অনেক স্তর প্রিলোড, দেখাতে ভারী।",depthtail:"আকার সবসময় সঠিক; এর বাইরে ধূসর ফোল্ডারে ক্লিক করুন।",leggreen:"স্ক্যান করা ফোল্ডার, আকার জানা",legorange:"বর্তমান স্ক্যানে পরিকল্পিত ফোল্ডার",legblue:"সারিতে থাকা ফোল্ডার (বর্তমান স্ক্যান শেষে)",legspin:"স্ক্যান হওয়া ফোল্ডার",leggrey:"অস্ক্যান, বাদ, জংশন বা সুরক্ষিত। স্ক্যান করতে ক্লিক করুন (২ স্তর)",legfile:"ফাইল",legbar:"বর্তমান ফোল্ডার আকারের অংশ",helpnav:"চলাচল: প্রবেশে ফোল্ডারে ক্লিক করুন, উপরের পথ দিয়ে উপরে যান।",helpgrey:"ধূসর ফোল্ডার: স্ক্যানে ক্লিক (আরও ২ স্তর)। স্ক্যানের সময় সারিতে (নীল), শেষে স্ক্যান।",helpfiles:"ফাইল: ফোল্ডার খুললে দেখা যায়, আকার অনুসারে, ১০০০ বড় পর্যন্ত।",helpfilter:"প্রদর্শন ফিল্টার (নতুন স্ক্যান বোতাম): ছোট আইটেম লুকায়, স্ক্যান না বদলে।",helpone:"একবারে একটি স্ক্যান: দুটি সমান্তরাল ডিস্ক স্ক্যান একে অপরকে ধীর করবে।",queuedmsg:"ফোল্ডার সারিতে ({n} অপেক্ষমাণ)। বর্তমান স্ক্যান শেষে স্ক্যান।",hiddenmsg:"সীমার নিচে {n} আইটেম প্রদর্শন ফিল্টারে লুকানো।",invalidpath:"অবৈধ বা অগম্য পথ।",windowopen:"উইন্ডো খোলা...",interrupted:"স্ক্যান বিঘ্নিত।",enterpath:"একটি পথ দিন।",serverstopped:"সার্ভার থামানো। আপনি এই ট্যাব বন্ধ করতে পারেন।",scangrey:"ধূসর ফোল্ডার স্ক্যান করুন",you:"(আপনি)",network:"(নেটওয়ার্ক)",drivefree:"খালি",removeone:"সরান",clearhist:"ইতিহাস মুছুন",exclusions:"বাদ",exclexp:"এই সিস্টেম ফোল্ডারগুলো স্ক্যানের সময় সবসময় বাদ দেওয়া হয়।",exclcustom:"আরও বাদ দিন (প্রতি লাইনে একটি পথ):",scandepthfull:"গভীরতা: সীমাহীন",scandepthn:"গভীরতা: {n}",stop:"থামান",pickerfail:"এই মেশিনে নেটিভ ফোল্ডার নির্বাচক খোলা যায়নি। পথটি সরাসরি লিখুন।",browsetitle:"ব্রাউজ",choose:"এই ফোল্ডার বেছে নিন",cancel:"বাতিল",upfolder:"মূল ফোল্ডার",sortby:"সাজান",byname:"নাম",bysize:"আকার"},
ur:{sub:"ڈسک استعمال تجزیہ کار",newscan:"نیا اسکین",noanalysis:"کوئی تجزیہ نہیں",theme:"ہلکی / گہری تھیم",legend:"کلید",help:"مدد",ready:"تیار۔",launching:"اسکین شروع ہو رہا ہے...",done:"اسکین مکمل۔ آزادانہ دیکھیں۔",inprogress:"جاری:",total:"کل:",scanningtag:"اسکین جاری",parent:".. (بالائی فولڈر)",scandots:"اسکین جاری...",empty:"خالی فولڈر",loadingfiles:"فائلیں لوڈ ہو رہی ہیں...",excluded:"خارج",junction:"جنکشن",unscanned:"غیر اسکین",protected:"محفوظ",error:"خرابی",scanshort:"اسکین...",queued:"قطار میں",clickscan:"اس فولڈر کو اسکین کرنے کے لیے کلک کریں (2 سطحیں)",filescap:"اس فولڈر کی 1000 سب سے بڑی فائلیں دکھائی جا رہی ہیں۔",filesafter:"موجودہ اسکین ختم ہونے پر فائلیں ظاہر ہوں گی۔",newanalysis:"نیا تجزیہ",modalhint:"تجزیے کے لیے فولڈر منتخب کریں، پھر گہرائی اور نمائش سیٹ کریں۔",folderlabel:"تجزیے کے لیے فولڈر",browse:"براؤز...",quickaccess:"فوری رسائی",recent:"حالیہ",drives:"ڈرائیوز",depthlabel:"جانچ کی گہرائی:",levels:"سطح",unlimited:"لامحدود (سب کچھ، بہت طویل ہو سکتا ہے)",displaylabel:"نمائش: چھوٹے آئٹمز چھپائیں",showall:"سب دکھائیں",hide1:"1 MB سے کم چھپائیں",hide100:"100 MB سے کم چھپائیں",hide1g:"1 GB سے کم چھپائیں",filterexp:"صرف نمائش فلٹر کرتا ہے، پڑھنے کے لیے۔ اسکین تیز نہیں کرتا: سائز ہمیشہ مکمل شمار ہوتے ہیں۔ کسی بھی وقت تبدیل کر سکتے ہیں۔",analyze:"تجزیہ کریں",quit:"سرور بند کریں",genby:"تخلیق کردہ:",depthunl:"پورا درخت دیکھتا اور دکھاتا ہے۔ بڑی ڈسک پر بہت طویل اور بھاری ہو سکتا ہے۔",depthlow:"ہلکا اور تیز: صرف پہلی سطحیں پیشگی لوڈ۔ فولڈر پر کلک کر کے نیچے جائیں۔",depthmid:"اچھا توازن: بیک وقت کئی سطحیں نظر آتی ہیں، ہموار۔",depthhigh:"تفصیلی: کئی سطحیں پیشگی لوڈ، دکھانے میں بھاری۔",depthtail:"سائز ہمیشہ درست؛ اس سے آگے کسی سرمئی فولڈر پر کلک کریں۔",leggreen:"اسکین شدہ فولڈر، سائز معلوم",legorange:"موجودہ اسکین میں منصوبہ بند فولڈر",legblue:"قطار میں فولڈر (موجودہ اسکین کے اختتام پر)",legspin:"اسکین ہوتا فولڈر",leggrey:"غیر اسکین، خارج، جنکشن یا محفوظ۔ اسکین کے لیے کلک کریں (2 سطحیں)",legfile:"فائل",legbar:"موجودہ فولڈر کے سائز کا حصہ",helpnav:"نیویگیشن: داخل ہونے کے لیے فولڈر پر کلک کریں، اوپر کے راستے سے واپس جائیں۔",helpgrey:"سرمئی فولڈرز: اسکین کے لیے کلک (2 مزید سطحیں)۔ اسکین کے دوران قطار میں (نیلا) اور آخر میں اسکین۔",helpfiles:"فائلیں: فولڈر کھولنے پر ظاہر، سائز کے مطابق، 1000 سب سے بڑی تک۔",helpfilter:"نمائش فلٹر (نیا اسکین بٹن): چھوٹے آئٹمز چھپاتا ہے، اسکین بدلے بغیر۔",helpone:"ایک وقت میں ایک اسکین: دو متوازی ڈسک اسکین ایک دوسرے کو سست کریں گے۔",queuedmsg:"فولڈر قطار میں ({n} منتظر)۔ موجودہ اسکین کے اختتام پر اسکین۔",hiddenmsg:"حد سے کم {n} آئٹم نمائش فلٹر سے چھپے۔",invalidpath:"غلط یا ناقابل رسائی راستہ۔",windowopen:"ونڈو کھلی...",interrupted:"اسکین منقطع۔",enterpath:"ایک راستہ درج کریں۔",serverstopped:"سرور رک گیا۔ آپ یہ ٹیب بند کر سکتے ہیں۔",scangrey:"سرمئی فولڈرز اسکین کریں",you:"(آپ)",network:"(نیٹ ورک)",drivefree:"خالی",removeone:"ہٹائیں",clearhist:"تاریخ صاف کریں",exclusions:"استثنیٰ",exclexp:"یہ سسٹم فولڈر اسکین کے دوران ہمیشہ چھوڑ دیے جاتے ہیں۔",exclcustom:"مزید خارج کریں (فی سطر ایک راستہ):",scandepthfull:"گہرائی: لامحدود",scandepthn:"گہرائی: {n}",stop:"روکیں",pickerfail:"اس مشین پر مقامی فولڈر منتخب کنندہ نہیں کھل سکا۔ راستہ براہ راست لکھیں۔",browsetitle:"براؤز",choose:"یہ فولڈر منتخب کریں",cancel:"منسوخ",upfolder:"بنیادی فولڈر",sortby:"ترتیب",byname:"نام",bysize:"سائز"},
id:{sub:"Penganalisis Penggunaan Disk",newscan:"Pemindaian baru",noanalysis:"Tidak ada analisis",theme:"Tema terang / gelap",legend:"Legenda",help:"Bantuan",ready:"Siap.",launching:"Memulai pemindaian...",done:"Pemindaian selesai. Jelajahi bebas.",inprogress:"Berlangsung:",total:"Total:",scanningtag:"memindai",parent:".. (folder induk)",scandots:"Memindai...",empty:"Folder kosong",loadingfiles:"Memuat berkas...",excluded:"dikecualikan",junction:"junction",unscanned:"belum dipindai",protected:"terlindungi",error:"galat",scanshort:"pindai...",queued:"antre",clickscan:"Klik untuk memindai folder ini (2 tingkat)",filescap:"Menampilkan 1000 berkas terbesar di folder ini.",filesafter:"Berkas muncul setelah pemindaian saat ini selesai.",newanalysis:"Analisis baru",modalhint:"Pilih folder untuk dianalisis, lalu atur kedalaman dan tampilan.",folderlabel:"Folder untuk dianalisis",browse:"Telusuri...",quickaccess:"Akses cepat",recent:"Terkini",drives:"Drive",depthlabel:"Kedalaman penjelajahan:",levels:"tingkat",unlimited:"Tak terbatas (menelusuri semua, bisa sangat lama)",displaylabel:"Tampilan: sembunyikan item kecil",showall:"Tampilkan semua",hide1:"Sembunyikan di bawah 1 MB",hide100:"Sembunyikan di bawah 100 MB",hide1g:"Sembunyikan di bawah 1 GB",filterexp:"Hanya memfilter tampilan, untuk keterbacaan. Tidak mempercepat pemindaian: ukuran selalu dihitung penuh. Dapat diubah kapan saja.",analyze:"Analisis",quit:"Hentikan server",genby:"Dibuat oleh:",depthunl:"Menelusuri dan menampilkan seluruh pohon. Bisa sangat lama dan berat pada disk besar.",depthlow:"Ringan dan cepat: hanya tingkat awal yang dimuat. Turun dengan mengklik folder.",depthmid:"Keseimbangan baik: beberapa tingkat terlihat sekaligus, lancar.",depthhigh:"Rinci: banyak tingkat dimuat awal, lebih berat ditampilkan.",depthtail:"Ukuran selalu akurat; di luar itu klik folder abu-abu untuk menjelajahinya.",leggreen:"Folder terpindai, ukuran diketahui",legorange:"Folder direncanakan dalam pemindaian saat ini",legblue:"Folder antre (dipindai saat pemindaian saat ini selesai)",legspin:"Folder sedang dipindai",leggrey:"Belum dipindai, dikecualikan, junction atau terlindungi. Klik untuk memindainya (2 tingkat)",legfile:"Berkas",legbar:"Bagian dari ukuran folder saat ini",helpnav:"Navigasi: klik folder untuk masuk, tapak di atas untuk naik.",helpgrey:"Folder abu-abu: klik untuk memindai (2 tingkat lagi). Selama pemindaian masuk antre (biru) dan dipindai di akhir.",helpfiles:"Berkas: muncul saat membuka folder, urut ukuran, dibatasi 1000 terbesar.",helpfilter:"Filter tampilan (tombol Pemindaian baru): menyembunyikan item kecil, tanpa mengubah pemindaian.",helpone:"Satu pemindaian sekaligus: dua pemindaian disk paralel akan saling memperlambat.",queuedmsg:"Folder diantre ({n} menunggu). Dipindai saat pemindaian saat ini selesai.",hiddenmsg:"{n} item di bawah ambang disembunyikan oleh filter tampilan.",invalidpath:"Jalur tidak valid atau tidak dapat diakses.",windowopen:"Jendela terbuka...",interrupted:"Pemindaian terhenti.",enterpath:"Masukkan jalur.",serverstopped:"Server dihentikan. Anda dapat menutup tab ini.",scangrey:"Pindai folder abu-abu",you:"(Anda)",network:"(jaringan)",drivefree:"kosong",removeone:"Hapus",clearhist:"Bersihkan riwayat",exclusions:"Pengecualian",exclexp:"Folder sistem ini selalu dilewati saat pemindaian.",exclcustom:"Kecualikan juga (satu jalur per baris):",scandepthfull:"Kedalaman: tak terbatas",scandepthn:"Kedalaman: {n}",stop:"Hentikan",pickerfail:"Pemilih folder bawaan tidak dapat dibuka di mesin ini. Ketik jalur secara langsung.",browsetitle:"Jelajahi",choose:"Pilih folder ini",cancel:"Batal",upfolder:"Folder induk",sortby:"Urutkan",byname:"nama",bysize:"ukuran"},ja:{"sub": "ディスク使用量アナライザー", "newscan": "新しいスキャン", "total": "合計：", "newanalysis": "新しい分析", "folderlabel": "分析するフォルダー", "browse": "参照...", "quickaccess": "クイックアクセス", "recent": "最近", "drives": "ドライブ", "depthlabel": "探索の深さ：", "unlimited": "無制限（すべてスキャン、非常に長くなる場合あり）", "displaylabel": "表示：小さい項目を隠す", "showall": "すべて表示", "analyze": "分析", "quit": "サーバーを終了", "legend": "凡例", "help": "ヘルプ", "stop": "中断", "choose": "このフォルダーを選択", "cancel": "キャンセル", "sortby": "並べ替え", "byname": "名前", "bysize": "サイズ", "scangrey": "グレーのフォルダーをスキャン", "leggreen": "スキャン済みフォルダー、サイズ判明", "legorange": "現在のスキャンで予定されたフォルダー", "legblue": "待機中のフォルダー（現在のスキャン後に処理）", "legspin": "スキャン中のフォルダー", "leggrey": "未スキャン、除外、ジャンクションまたは保護。クリックしてスキャン（2階層）", "legfile": "ファイル", "legbar": "現在のフォルダーサイズに占める割合", "helpnav": "操作：フォルダーをクリックして開き、上部のパンくずで上へ戻ります。", "helpgrey": "グレーのフォルダー：クリックでスキャン（さらに2階層）。スキャン中は待機（青）になり、最後に処理されます。", "helpfiles": "ファイル：フォルダーを開くと表示、サイズ順、最大1000件まで。", "helpfilter": "表示フィルター（新しいスキャンのボタン）：スキャンは変えず、小さい項目を隠して見やすくします。", "helpone": "同時に1つのスキャンのみ：2つの並行ディスクスキャンは互いに遅くなります。", "noanalysis": "分析なし", "theme": "ライト / ダークテーマ", "ready": "準備完了。", "launching": "スキャンを開始中...", "done": "スキャン完了。自由に移動できます。", "inprogress": "進行中：", "scanningtag": "スキャン中", "parent": ".. (親フォルダー)", "scandots": "スキャン中...", "empty": "空のフォルダー", "loadingfiles": "ファイルを読み込み中...", "excluded": "除外", "junction": "ジャンクション", "unscanned": "未スキャン", "protected": "保護", "error": "エラー", "scanshort": "スキャン...", "queued": "待機中", "levels": "階層", "hide1": "1 MB未満を隠す", "hide100": "100 MB未満を隠す", "hide1g": "1 GB未満を隠す", "genby": "生成：", "invalidpath": "無効またはアクセスできないパス。", "windowopen": "ウィンドウを開きました...", "interrupted": "スキャンを中断しました。", "enterpath": "パスを入力してください。", "serverstopped": "サーバーを停止しました。このタブを閉じられます。", "you": "(あなた)", "network": "(ネットワーク)", "drivefree": "空き", "removeone": "削除", "clearhist": "履歴を消去", "exclusions": "除外", "scandepthfull": "深さ：無制限", "scandepthn": "深さ：{n}", "browsetitle": "参照", "upfolder": "親フォルダー", "clickscan": "クリックしてこのフォルダーをスキャン（2階層）", "depthhigh": "詳細：多くの階層を事前読み込み、表示は重め。", "depthlow": "軽量で高速表示：最初の階層のみ読み込み。フォルダーをクリックして下へ。", "depthmid": "バランス良好：複数階層を一度に表示、動作は滑らか。", "depthtail": "サイズは常に正確です。この深さより先はグレーのフォルダーをクリックして探索。", "depthunl": "ツリー全体を走査して表示。大きなディスクでは非常に長く重くなります。", "exclcustom": "追加で除外（1行に1パス）：", "exclexp": "これらのシステムフォルダーはスキャン時に常にスキップされます。", "filesafter": "ファイルは現在のスキャン完了後に表示されます。", "filescap": "このフォルダーの最大1000ファイルまで表示。", "filterexp": "表示のみを絞り込みます（見やすさのため）。スキャンは速くなりません：サイズは常に完全に計算されます。いつでも変更可能。", "hiddenmsg": "{n} 件がしきい値未満のため表示フィルターで非表示。", "modalhint": "分析するフォルダーを選び、深さと表示を設定してください。", "pickerfail": "このマシンではネイティブのフォルダー選択を開けません。パスを直接入力してください。", "queuedmsg": "フォルダーを待機に追加（{n} 件待ち）。現在のスキャン後に処理します。"},ko:{"sub": "디스크 사용량 분석기", "newscan": "새 스캔", "total": "합계:", "newanalysis": "새 분석", "folderlabel": "분석할 폴더", "browse": "찾아보기...", "quickaccess": "빠른 액세스", "recent": "최근", "drives": "드라이브", "depthlabel": "탐색 깊이:", "unlimited": "무제한(전체 스캔, 매우 오래 걸릴 수 있음)", "displaylabel": "표시: 작은 항목 숨기기", "showall": "모두 표시", "analyze": "분석", "quit": "서버 종료", "legend": "범례", "help": "도움말", "stop": "중단", "choose": "이 폴더 선택", "cancel": "취소", "sortby": "정렬", "byname": "이름", "bysize": "크기", "scangrey": "회색 폴더 스캔", "leggreen": "스캔된 폴더, 크기 확인됨", "legorange": "현재 스캔에 예정된 폴더", "legblue": "대기 중인 폴더(현재 스캔 후 처리)", "legspin": "스캔 중인 폴더", "leggrey": "미스캔, 제외, 정션 또는 보호됨. 클릭하여 스캔(2단계)", "legfile": "파일", "legbar": "현재 폴더 크기에서의 비율", "helpnav": "탐색: 폴더를 클릭하여 열고, 상단 경로로 위로 이동합니다.", "helpgrey": "회색 폴더: 클릭하여 스캔(2단계 더). 스캔 중에는 대기(파랑)되어 마지막에 처리됩니다.", "helpfiles": "파일: 폴더를 열면 표시, 크기순 정렬, 최대 1000개.", "helpfilter": "표시 필터(새 스캔 버튼): 스캔은 바꾸지 않고 작은 항목을 숨겨 가독성을 높입니다.", "helpone": "한 번에 하나의 스캔: 두 개의 병렬 디스크 스캔은 서로 느려집니다.", "noanalysis": "분석 없음", "theme": "밝은 / 어두운 테마", "ready": "준비됨.", "launching": "스캔 시작 중...", "done": "스캔 완료. 자유롭게 탐색하세요.", "inprogress": "진행 중:", "scanningtag": "스캔 중", "parent": ".. (상위 폴더)", "scandots": "스캔 중...", "empty": "빈 폴더", "loadingfiles": "파일 불러오는 중...", "excluded": "제외", "junction": "정션", "unscanned": "미스캔", "protected": "보호됨", "error": "오류", "scanshort": "스캔...", "queued": "대기 중", "levels": "단계", "hide1": "1 MB 미만 숨기기", "hide100": "100 MB 미만 숨기기", "hide1g": "1 GB 미만 숨기기", "genby": "생성:", "invalidpath": "잘못되었거나 접근할 수 없는 경로.", "windowopen": "창이 열렸습니다...", "interrupted": "스캔이 중단되었습니다.", "enterpath": "경로를 입력하세요.", "serverstopped": "서버가 중지되었습니다. 이 탭을 닫아도 됩니다.", "you": "(나)", "network": "(네트워크)", "drivefree": "여유", "removeone": "제거", "clearhist": "기록 지우기", "exclusions": "제외", "scandepthfull": "깊이: 무제한", "scandepthn": "깊이: {n}", "browsetitle": "찾아보기", "upfolder": "상위 폴더", "clickscan": "클릭하여 이 폴더 스캔(2단계)", "depthhigh": "상세: 많은 단계를 미리 로드, 표시가 무거움.", "depthlow": "가볍고 빠른 표시: 첫 단계만 로드. 폴더를 클릭해 내려가세요.", "depthmid": "적절한 균형: 여러 단계를 한 번에 표시, 여전히 부드러움.", "depthtail": "크기는 항상 정확합니다. 이 깊이 너머는 회색 폴더를 클릭해 탐색하세요.", "depthunl": "전체 트리를 탐색하고 표시합니다. 큰 디스크에서는 매우 오래 걸리고 무거울 수 있습니다.", "exclcustom": "추가 제외(한 줄에 경로 하나):", "exclexp": "이 시스템 폴더는 스캔 시 항상 건너뜁니다.", "filesafter": "파일은 현재 스캔이 끝나면 표시됩니다.", "filescap": "이 폴더의 가장 큰 1000개 파일로 표시가 제한됩니다.", "filterexp": "가독성을 위해 표시만 필터링합니다. 스캔 속도는 빨라지지 않습니다: 크기는 항상 완전히 계산됩니다. 언제든 변경 가능.", "hiddenmsg": "임계값 미만 {n}개 항목이 표시 필터로 숨겨졌습니다.", "modalhint": "분석할 폴더를 선택한 다음 깊이와 표시를 설정하세요.", "pickerfail": "이 컴퓨터에서 기본 폴더 선택기를 열 수 없습니다. 경로를 직접 입력하세요.", "queuedmsg": "폴더가 대기열에 추가됨({n}개 대기). 현재 스캔 후 스캔됩니다."},it:{"sub": "Analizzatore spazio su disco", "newscan": "Nuova scansione", "total": "Totale:", "newanalysis": "Nuova analisi", "folderlabel": "Cartella da analizzare", "browse": "Sfoglia...", "quickaccess": "Accesso rapido", "recent": "Recenti", "drives": "Unità", "depthlabel": "Profondità di esplorazione:", "unlimited": "Illimitata (analizza tutto, può essere molto lunga)", "displaylabel": "Visualizzazione: nascondi elementi piccoli", "showall": "Mostra tutto", "analyze": "Analizza", "quit": "Chiudi il server", "legend": "Legenda", "help": "Aiuto", "stop": "Interrompi", "choose": "Scegli questa cartella", "cancel": "Annulla", "sortby": "Ordina", "byname": "nome", "bysize": "dimensione", "scangrey": "Analizza le cartelle grigie", "leggreen": "Cartella analizzata, dimensione nota", "legorange": "Cartella prevista nella scansione corrente", "legblue": "Cartella in coda (analizzata dopo la scansione corrente)", "legspin": "Cartella in analisi", "leggrey": "Non analizzata, esclusa, giunzione o protetta. Clicca per analizzarla (2 livelli)", "legfile": "File", "legbar": "Quota sulla dimensione della cartella corrente", "helpnav": "Naviga: clicca una cartella per entrare, i breadcrumb in alto per salire.", "helpgrey": "Cartelle grigie: clicca per analizzarle (2 livelli in piu). Durante una scansione vanno in coda (blu) e si analizzano alla fine.", "helpfiles": "File: mostrati quando apri una cartella, ordinati per dimensione, limitati ai 1000 piu grandi.", "helpfilter": "Filtro di visualizzazione (pulsante Nuova scansione): nasconde gli elementi piccoli per leggibilita, senza cambiare la scansione.", "helpone": "Una scansione alla volta: due scansioni disco in parallelo si rallenterebbero.", "noanalysis": "Nessuna analisi", "theme": "Tema chiaro / scuro", "ready": "Pronto.", "launching": "Avvio della scansione...", "done": "Scansione completata. Naviga liberamente.", "inprogress": "In corso:", "scanningtag": "in scansione", "parent": ".. (cartella superiore)", "scandots": "Scansione in corso...", "empty": "Cartella vuota", "loadingfiles": "Caricamento file...", "excluded": "escluso", "junction": "giunzione", "unscanned": "non analizzato", "protected": "protetto", "error": "errore", "scanshort": "scansione...", "queued": "in coda", "levels": "livello/i", "hide1": "Nascondi sotto 1 MB", "hide100": "Nascondi sotto 100 MB", "hide1g": "Nascondi sotto 1 GB", "genby": "Generato da:", "invalidpath": "Percorso non valido o inaccessibile.", "windowopen": "Finestra aperta...", "interrupted": "Scansione interrotta.", "enterpath": "Inserisci un percorso.", "serverstopped": "Server arrestato. Puoi chiudere questa scheda.", "you": "(tu)", "network": "(rete)", "drivefree": "liberi", "removeone": "Rimuovi", "clearhist": "Cancella cronologia", "exclusions": "Esclusioni", "scandepthfull": "Profondita: illimitata", "scandepthn": "Profondita: {n}", "browsetitle": "Sfoglia", "upfolder": "Cartella superiore", "clickscan": "Clicca per analizzare questa cartella (2 livelli)", "depthhigh": "Dettagliato: molti livelli precaricati, piu pesante da visualizzare.", "depthlow": "Leggero e veloce da visualizzare: solo i primi livelli sono caricati. Scendi cliccando una cartella.", "depthmid": "Buon compromesso: piu livelli visibili subito, visualizzazione fluida.", "depthtail": "Le dimensioni sono sempre esatte; oltre questa profondita, clicca una cartella grigia per esplorarla.", "depthunl": "Percorre e mostra l'intero albero. Puo essere molto lungo e pesante su un disco grande.", "exclcustom": "Escludi anche (un percorso per riga):", "exclexp": "Queste cartelle di sistema vengono sempre saltate durante la scansione.", "filesafter": "I file appariranno alla fine della scansione in corso.", "filescap": "Visualizzazione limitata ai 1000 file piu grandi di questa cartella.", "filterexp": "Filtra solo la visualizzazione, per leggibilita. Non accelera la scansione: le dimensioni sono sempre calcolate per intero. Modificabile in qualsiasi momento.", "hiddenmsg": "{n} elemento/i sotto la soglia nascosto/i dal filtro di visualizzazione.", "modalhint": "Scegli la cartella da analizzare, poi imposta profondita e visualizzazione.", "pickerfail": "Il selettore di cartelle nativo non si e aperto su questa macchina. Digita il percorso direttamente.", "queuedmsg": "Cartella in coda ({n} in attesa). Analizzata alla fine della scansione in corso."},tr:{"sub": "Disk Kullanım Analizörü", "newscan": "Yeni tarama", "total": "Toplam:", "newanalysis": "Yeni analiz", "folderlabel": "Analiz edilecek klasör", "browse": "Gözat...", "quickaccess": "Hızlı erişim", "recent": "Son kullanılanlar", "drives": "Sürücüler", "depthlabel": "Keşif derinliği:", "unlimited": "Sınırsız (her şeyi tarar, çok uzun sürebilir)", "displaylabel": "Görünüm: küçük ögeleri gizle", "showall": "Tümünü göster", "analyze": "Analiz et", "quit": "Sunucuyu kapat", "legend": "Gösterge", "help": "Yardım", "stop": "Durdur", "choose": "Bu klasörü seç", "cancel": "İptal", "sortby": "Sırala", "byname": "ad", "bysize": "boyut", "scangrey": "Gri klasörleri tara", "leggreen": "Taranmis klasor, boyut biliniyor", "legorange": "Gecerli taramada planlanan klasor", "legblue": "Sirada bekleyen klasor (gecerli taramadan sonra taranir)", "legspin": "Taranan klasor", "leggrey": "Taranmadi, haric, baglanti veya korumali. Taramak icin tiklayin (2 duzey)", "legfile": "Dosya", "legbar": "Gecerli klasor boyutundaki payi", "helpnav": "Gezinme: girmek icin bir klasore tiklayin, yukari cikmak icin ustteki yol izini kullanin.", "helpgrey": "Gri klasorler: taramak icin tiklayin (2 duzey daha). Tarama sirasinda siraya girer (mavi) ve sonda taranir.", "helpfiles": "Dosyalar: bir klasoru actiginizda gosterilir, boyuta gore siralanir, en buyuk 1000 ile sinirlidir.", "helpfilter": "Goruntu filtresi (Yeni tarama dugmesi): taramayi degistirmeden kucuk ogeleri gizler.", "helpone": "Ayni anda tek tarama: iki paralel disk taramasi birbirini yavaslatir.", "noanalysis": "Analiz yok", "theme": "Acik / koyu tema", "ready": "Hazir.", "launching": "Tarama baslatiliyor...", "done": "Tarama tamamlandi. Serbestce gezinin.", "inprogress": "Devam ediyor:", "scanningtag": "taraniyor", "parent": ".. (ust klasor)", "scandots": "Taraniyor...", "empty": "Bos klasor", "loadingfiles": "Dosyalar yukleniyor...", "excluded": "haric", "junction": "baglanti", "unscanned": "taranmadi", "protected": "korumali", "error": "hata", "scanshort": "tarama...", "queued": "sirada", "levels": "duzey", "hide1": "1 MB altini gizle", "hide100": "100 MB altini gizle", "hide1g": "1 GB altini gizle", "genby": "Olusturan:", "invalidpath": "Gecersiz veya erisilemeyen yol.", "windowopen": "Pencere acildi...", "interrupted": "Tarama kesildi.", "enterpath": "Lutfen bir yol girin.", "serverstopped": "Sunucu durduruldu. Bu sekmeyi kapatabilirsiniz.", "you": "(siz)", "network": "(ag)", "drivefree": "bos", "removeone": "Kaldir", "clearhist": "Gecmisi temizle", "exclusions": "Haric tutulanlar", "scandepthfull": "Derinlik: sinirsiz", "scandepthn": "Derinlik: {n}", "browsetitle": "Gozat", "upfolder": "Ust klasor", "clickscan": "Bu klasoru taramak icin tiklayin (2 duzey)", "depthhigh": "Ayrintili: bircok duzey onceden yuklenir, gosterim daha agir.", "depthlow": "Hafif ve hizli gosterim: yalnizca ilk duzeyler yuklenir. Bir klasore tiklayarak inin.", "depthmid": "Iyi denge: birkac duzey ayni anda gorunur, hala akici.", "depthtail": "Boyutlar her zaman kesindir; bu derinligin otesinde kesfetmek icin gri bir klasore tiklayin.", "depthunl": "Tum agaci gezer ve gosterir. Buyuk bir diskte cok uzun ve agir olabilir.", "exclcustom": "Ayrica haric tut (her satira bir yol):", "exclexp": "Bu sistem klasorleri tarama sirasinda her zaman atlanir.", "filesafter": "Dosyalar mevcut tarama bitince gosterilecek.", "filescap": "Gosterim bu klasordeki en buyuk 1000 dosya ile sinirli.", "filterexp": "Yalnizca gosterimi filtreler, okunabilirlik icin. Taramayi hizlandirmaz: boyutlar her zaman tam hesaplanir. Her zaman degistirilebilir.", "hiddenmsg": "Esigin altindaki {n} oge gosterim filtresiyle gizlendi.", "modalhint": "Analiz edilecek klasoru secin, ardindan derinligi ve gosterimi ayarlayin.", "pickerfail": "Yerel klasor secici bu makinede acilamadi. Yolu dogrudan yazin.", "queuedmsg": "Klasor siraya alindi ({n} bekliyor). Mevcut taramanin sonunda taranacak."},vi:{"sub": "Trình phân tích dung lượng đĩa", "newscan": "Quét mới", "total": "Tổng:", "newanalysis": "Phân tích mới", "folderlabel": "Thư mục cần phân tích", "browse": "Duyệt...", "quickaccess": "Truy cập nhanh", "recent": "Gần đây", "drives": "Ổ đĩa", "depthlabel": "Độ sâu khám phá:", "unlimited": "Không giới hạn (quét toàn bộ, có thể rất lâu)", "displaylabel": "Hiển thị: ẩn mục nhỏ", "showall": "Hiện tất cả", "analyze": "Phân tích", "quit": "Thoát máy chủ", "legend": "Chú giải", "help": "Trợ giúp", "stop": "Dừng", "choose": "Chọn thư mục này", "cancel": "Hủy", "sortby": "Sắp xếp", "byname": "tên", "bysize": "kích thước", "scangrey": "Quét thư mục xám", "leggreen": "Thu muc da quet, biet kich thuoc", "legorange": "Thu muc du kien trong lan quet hien tai", "legblue": "Thu muc trong hang doi (quet sau lan quet hien tai)", "legspin": "Thu muc dang duoc quet", "leggrey": "Chua quet, loai tru, junction hoac duoc bao ve. Nhap de quet (2 cap)", "legfile": "Tep", "legbar": "Ty le trong kich thuoc thu muc hien tai", "helpnav": "Dieu huong: nhap vao thu muc de vao, duong dan tren cung de len.", "helpgrey": "Thu muc xam: nhap de quet (them 2 cap). Trong khi quet chung vao hang doi (xanh) va duoc quet o cuoi.", "helpfiles": "Tep: hien khi ban mo thu muc, sap theo kich thuoc, gioi han 1000 tep lon nhat.", "helpfilter": "Bo loc hien thi (nut Quet moi): an cac muc nho de de doc, khong thay doi lan quet.", "helpone": "Mot lan quet moi luc: hai lan quet dia song song se lam cham nhau.", "noanalysis": "Khong co phan tich", "theme": "Giao dien sang / toi", "ready": "San sang.", "launching": "Dang bat dau quet...", "done": "Quet xong. Di chuyen tu do.", "inprogress": "Dang xu ly:", "scanningtag": "dang quet", "parent": ".. (thu muc cha)", "scandots": "Dang quet...", "empty": "Thu muc trong", "loadingfiles": "Dang tai tep...", "excluded": "loai tru", "junction": "junction", "unscanned": "chua quet", "protected": "duoc bao ve", "error": "loi", "scanshort": "quet...", "queued": "trong hang doi", "levels": "cap", "hide1": "An duoi 1 MB", "hide100": "An duoi 100 MB", "hide1g": "An duoi 1 GB", "genby": "Tao boi:", "invalidpath": "Duong dan khong hop le hoac khong truy cap duoc.", "windowopen": "Da mo cua so...", "interrupted": "Da dung quet.", "enterpath": "Vui long nhap duong dan.", "serverstopped": "May chu da dung. Ban co the dong tab nay.", "you": "(ban)", "network": "(mang)", "drivefree": "trong", "removeone": "Xoa", "clearhist": "Xoa lich su", "exclusions": "Loai tru", "scandepthfull": "Do sau: khong gioi han", "scandepthn": "Do sau: {n}", "browsetitle": "Duyet", "upfolder": "Thu muc cha", "clickscan": "Nhap de quet thu muc nay (2 cap)", "depthhigh": "Chi tiet: nhieu cap duoc tai truoc, hien thi nang hon.", "depthlow": "Nhe va hien thi nhanh: chi tai cac cap dau. Di sau bang cach nhap vao thu muc.", "depthmid": "Can bang tot: nhieu cap hien cung luc, van muot.", "depthtail": "Kich thuoc luon chinh xac; sau do sau nay, nhap thu muc xam de kham pha.", "depthunl": "Duyet va hien thi toan bo cay. Co the rat lau va nang tren dia lon.", "exclcustom": "Loai tru them (moi dong mot duong dan):", "exclexp": "Cac thu muc he thong nay luon bi bo qua khi quet.", "filesafter": "Tep se hien khi lan quet hien tai ket thuc.", "filescap": "Hien thi gioi han 1000 tep lon nhat trong thu muc nay.", "filterexp": "Chi loc hien thi, de de doc. Khong lam quet nhanh hon: kich thuoc luon duoc tinh day du. Co the thay doi bat cu luc nao.", "hiddenmsg": "{n} muc duoi nguong bi an boi bo loc hien thi.", "modalhint": "Chon thu muc can phan tich, roi dat do sau va hien thi.", "pickerfail": "Khong mo duoc bo chon thu muc goc tren may nay. Hay nhap duong dan truc tiep.", "queuedmsg": "Thu muc da vao hang doi ({n} dang cho). Se quet o cuoi lan quet hien tai."},pl:{"sub": "Analizator zajętości dysku", "newscan": "Nowe skanowanie", "total": "Razem:", "newanalysis": "Nowa analiza", "folderlabel": "Folder do analizy", "browse": "Przeglądaj...", "quickaccess": "Szybki dostęp", "recent": "Ostatnie", "drives": "Dyski", "depthlabel": "Głębokość eksploracji:", "unlimited": "Bez ograniczeń (skanuje wszystko, może trwać bardzo długo)", "displaylabel": "Widok: ukryj małe elementy", "showall": "Pokaż wszystko", "analyze": "Analizuj", "quit": "Zamknij serwer", "legend": "Legenda", "help": "Pomoc", "stop": "Przerwij", "choose": "Wybierz ten folder", "cancel": "Anuluj", "sortby": "Sortuj", "byname": "nazwa", "bysize": "rozmiar", "scangrey": "Skanuj szare foldery", "leggreen": "Przeskanowany folder, rozmiar znany", "legorange": "Folder zaplanowany w biezacym skanowaniu", "legblue": "Folder w kolejce (skanowany po biezacym skanowaniu)", "legspin": "Folder w trakcie skanowania", "leggrey": "Nieskanowany, wykluczony, polaczenie lub chroniony. Kliknij, aby przeskanowac (2 poziomy)", "legfile": "Plik", "legbar": "Udzial w rozmiarze biezacego folderu", "helpnav": "Nawigacja: kliknij folder, aby wejsc, sciezka u gory, aby wroci wyzej.", "helpgrey": "Szare foldery: kliknij, aby przeskanowac (2 kolejne poziomy). Podczas skanowania trafiaja do kolejki (niebieski) i sa skanowane na koncu.", "helpfiles": "Pliki: pokazywane po otwarciu folderu, sortowane wg rozmiaru, ograniczone do 1000 najwiekszych.", "helpfilter": "Filtr wyswietlania (przycisk Nowe skanowanie): ukrywa male elementy dla czytelnosci, bez zmiany skanowania.", "helpone": "Jedno skanowanie naraz: dwa rownolegle skanowania dysku spowalnialyby sie nawzajem.", "noanalysis": "Brak analizy", "theme": "Motyw jasny / ciemny", "ready": "Gotowe.", "launching": "Rozpoczynanie skanowania...", "done": "Skanowanie zakonczone. Nawiguj swobodnie.", "inprogress": "W toku:", "scanningtag": "skanowanie", "parent": ".. (folder nadrzedny)", "scandots": "Skanowanie...", "empty": "Pusty folder", "loadingfiles": "Ladowanie plikow...", "excluded": "wykluczony", "junction": "polaczenie", "unscanned": "nieskanowany", "protected": "chroniony", "error": "blad", "scanshort": "skan...", "queued": "w kolejce", "levels": "poziom(y)", "hide1": "Ukryj ponizej 1 MB", "hide100": "Ukryj ponizej 100 MB", "hide1g": "Ukryj ponizej 1 GB", "genby": "Wygenerowano przez:", "invalidpath": "Nieprawidlowa lub niedostepna sciezka.", "windowopen": "Okno otwarte...", "interrupted": "Skanowanie przerwane.", "enterpath": "Podaj sciezke.", "serverstopped": "Serwer zatrzymany. Mozesz zamknac te karte.", "you": "(ty)", "network": "(siec)", "drivefree": "wolne", "removeone": "Usun", "clearhist": "Wyczysc historie", "exclusions": "Wykluczenia", "scandepthfull": "Glebokosc: bez limitu", "scandepthn": "Glebokosc: {n}", "browsetitle": "Przegladaj", "upfolder": "Folder nadrzedny", "clickscan": "Kliknij, aby przeskanowac ten folder (2 poziomy)", "depthhigh": "Szczegolowo: wiele poziomow wczytanych z gory, ciezsze wyswietlanie.", "depthlow": "Lekko i szybko: wczytane tylko pierwsze poziomy. Schodz nizej klikajac folder.", "depthmid": "Dobry kompromis: kilka poziomow widocznych od razu, nadal plynnie.", "depthtail": "Rozmiary sa zawsze dokladne; ponizej tej glebokosci kliknij szary folder, aby go zbadac.", "depthunl": "Przechodzi i wyswietla cale drzewo. Moze byc bardzo dlugie i ciezkie na duzym dysku.", "exclcustom": "Wyklucz takze (jedna sciezka na wiersz):", "exclexp": "Te foldery systemowe sa zawsze pomijane podczas skanowania.", "filesafter": "Pliki pojawia sie po zakonczeniu biezacego skanowania.", "filescap": "Wyswietlanie ograniczone do 1000 najwiekszych plikow w tym folderze.", "filterexp": "Filtruje tylko wyswietlanie, dla czytelnosci. Nie przyspiesza skanowania: rozmiary sa zawsze liczone w calosci. Mozna zmienic w kazdej chwili.", "hiddenmsg": "{n} element(ow) ponizej progu ukryto filtrem wyswietlania.", "modalhint": "Wybierz folder do analizy, potem ustaw glebokosc i wyswietlanie.", "pickerfail": "Natywny wybor folderu nie otworzyl sie na tym komputerze. Wpisz sciezke recznie.", "queuedmsg": "Folder w kolejce ({n} czeka). Zostanie przeskanowany po biezacym skanowaniu."},nl:{"sub": "Schijfruimte-analyser", "newscan": "Nieuwe scan", "total": "Totaal:", "newanalysis": "Nieuwe analyse", "folderlabel": "Te analyseren map", "browse": "Bladeren...", "quickaccess": "Snelle toegang", "recent": "Recent", "drives": "Stations", "depthlabel": "Verkenningsdiepte:", "unlimited": "Onbeperkt (scant alles, kan zeer lang duren)", "displaylabel": "Weergave: kleine items verbergen", "showall": "Alles tonen", "analyze": "Analyseren", "quit": "Server afsluiten", "legend": "Legenda", "help": "Help", "stop": "Stoppen", "choose": "Deze map kiezen", "cancel": "Annuleren", "sortby": "Sorteren", "byname": "naam", "bysize": "grootte", "scangrey": "Grijze mappen scannen", "leggreen": "Gescande map, grootte bekend", "legorange": "Map gepland in huidige scan", "legblue": "Map in wachtrij (gescand na huidige scan)", "legspin": "Map wordt gescand", "leggrey": "Niet gescand, uitgesloten, junction of beveiligd. Klik om te scannen (2 niveaus)", "legfile": "Bestand", "legbar": "Aandeel in grootte van huidige map", "helpnav": "Navigeren: klik op een map om te openen, kruimelpad bovenaan om omhoog te gaan.", "helpgrey": "Grijze mappen: klik om te scannen (2 extra niveaus). Tijdens een scan komen ze in de wachtrij (blauw) en worden aan het einde gescand.", "helpfiles": "Bestanden: getoond bij het openen van een map, gesorteerd op grootte, beperkt tot de 1000 grootste.", "helpfilter": "Weergavefilter (knop Nieuwe scan): verbergt kleine items voor leesbaarheid, zonder de scan te wijzigen.", "helpone": "Een scan tegelijk: twee parallelle schijfscans zouden elkaar vertragen.", "noanalysis": "Geen analyse", "theme": "Licht / donker thema", "ready": "Klaar.", "launching": "Scan wordt gestart...", "done": "Scan voltooid. Navigeer vrij.", "inprogress": "Bezig:", "scanningtag": "scannen", "parent": ".. (bovenliggende map)", "scandots": "Scannen...", "empty": "Lege map", "loadingfiles": "Bestanden laden...", "excluded": "uitgesloten", "junction": "junction", "unscanned": "niet gescand", "protected": "beveiligd", "error": "fout", "scanshort": "scan...", "queued": "in wachtrij", "levels": "niveau(s)", "hide1": "Verberg onder 1 MB", "hide100": "Verberg onder 100 MB", "hide1g": "Verberg onder 1 GB", "genby": "Gegenereerd door:", "invalidpath": "Ongeldig of ontoegankelijk pad.", "windowopen": "Venster geopend...", "interrupted": "Scan onderbroken.", "enterpath": "Voer een pad in.", "serverstopped": "Server gestopt. U kunt dit tabblad sluiten.", "you": "(u)", "network": "(netwerk)", "drivefree": "vrij", "removeone": "Verwijderen", "clearhist": "Geschiedenis wissen", "exclusions": "Uitsluitingen", "scandepthfull": "Diepte: onbeperkt", "scandepthn": "Diepte: {n}", "browsetitle": "Bladeren", "upfolder": "Bovenliggende map", "clickscan": "Klik om deze map te scannen (2 niveaus)", "depthhigh": "Gedetailleerd: veel niveaus vooraf geladen, zwaarder om weer te geven.", "depthlow": "Licht en snel: alleen de eerste niveaus worden geladen. Ga dieper door op een map te klikken.", "depthmid": "Goede balans: meerdere niveaus tegelijk zichtbaar, nog steeds vloeiend.", "depthtail": "Groottes zijn altijd exact; voorbij deze diepte klikt u op een grijze map om die te verkennen.", "depthunl": "Doorloopt en toont de hele boom. Kan zeer lang en zwaar zijn op een grote schijf.", "exclcustom": "Ook uitsluiten (een pad per regel):", "exclexp": "Deze systeemmappen worden tijdens de scan altijd overgeslagen.", "filesafter": "Bestanden verschijnen aan het einde van de huidige scan.", "filescap": "Weergave beperkt tot de 1000 grootste bestanden in deze map.", "filterexp": "Filtert alleen de weergave, voor leesbaarheid. Versnelt de scan niet: groottes worden altijd volledig berekend. Op elk moment aanpasbaar.", "hiddenmsg": "{n} item(s) onder de drempel verborgen door het weergavefilter.", "modalhint": "Kies de te analyseren map en stel dan diepte en weergave in.", "pickerfail": "De native mapkiezer kon niet worden geopend op deze machine. Typ het pad direct.", "queuedmsg": "Map in wachtrij ({n} wachtend). Wordt gescand aan het einde van de huidige scan."},uk:{"sub": "Аналізатор використання диска", "newscan": "Нове сканування", "total": "Разом:", "newanalysis": "Новий аналіз", "folderlabel": "Папка для аналізу", "browse": "Огляд...", "quickaccess": "Швидкий доступ", "recent": "Нещодавні", "drives": "Диски", "depthlabel": "Глибина дослідження:", "unlimited": "Без обмежень (сканує все, може бути дуже довго)", "displaylabel": "Показ: приховати малі елементи", "showall": "Показати все", "analyze": "Аналізувати", "quit": "Зупинити сервер", "legend": "Легенда", "help": "Довідка", "stop": "Перервати", "choose": "Вибрати цю папку", "cancel": "Скасувати", "sortby": "Сортувати", "byname": "назва", "bysize": "розмір", "scangrey": "Сканувати сірі папки", "leggreen": "Проскановану папку, розмір відомий", "legorange": "Папка запланована в поточному скануванні", "legblue": "Папка в черзі (сканується після поточного сканування)", "legspin": "Папка сканується", "leggrey": "Не проскановано, виключено, сполучення або захищено. Натисніть, щоб просканувати (2 рівні)", "legfile": "Файл", "legbar": "Частка в розмірі поточної папки", "helpnav": "Навігація: натисніть папку, щоб увійти, ланцюжок угорі, щоб піднятися.", "helpgrey": "Сірі папки: натисніть, щоб просканувати (ще 2 рівні). Під час сканування вони стають у чергу (синій) і скануються в кінці.", "helpfiles": "Файли: показуються при відкритті папки, сортуються за розміром, обмежено 1000 найбільшими.", "helpfilter": "Фільтр показу (кнопка Нове сканування): приховує малі елементи для читабельності, не змінюючи сканування.", "helpone": "Одне сканування за раз: два паралельні сканування диска сповільнювали б одне одного.", "noanalysis": "Немає аналізу", "theme": "Світла / темна тема", "ready": "Готово.", "launching": "Запуск сканування...", "done": "Сканування завершено. Вільна навігація.", "inprogress": "Виконується:", "scanningtag": "сканування", "parent": ".. (батьківська папка)", "scandots": "Сканування...", "empty": "Порожня папка", "loadingfiles": "Завантаження файлів...", "excluded": "виключено", "junction": "сполучення", "unscanned": "не скановано", "protected": "захищено", "error": "помилка", "scanshort": "скан...", "queued": "у черзі", "levels": "рівень(і)", "hide1": "Приховати менше 1 МБ", "hide100": "Приховати менше 100 МБ", "hide1g": "Приховати менше 1 ГБ", "genby": "Створено:", "invalidpath": "Недійсний або недоступний шлях.", "windowopen": "Вікно відкрито...", "interrupted": "Сканування перервано.", "enterpath": "Вкажіть шлях.", "serverstopped": "Сервер зупинено. Можна закрити цю вкладку.", "you": "(ви)", "network": "(мережа)", "drivefree": "вільно", "removeone": "Видалити", "clearhist": "Очистити історію", "exclusions": "Виключення", "scandepthfull": "Глибина: без обмежень", "scandepthn": "Глибина: {n}", "browsetitle": "Огляд", "upfolder": "Батьківська папка", "clickscan": "Натисніть, щоб просканувати цю папку (2 рівні)", "depthhigh": "Детально: багато рівнів завантажено наперед, важче відображати.", "depthlow": "Легко і швидко: завантажуються лише перші рівні. Заглиблюйтесь, натискаючи папку.", "depthmid": "Добрий баланс: кілька рівнів видно одразу, все ще плавно.", "depthtail": "Розміри завжди точні; глибше за цей рівень натисніть сіру папку, щоб її дослідити.", "depthunl": "Проходить і показує все дерево. Може бути дуже довго і важко на великому диску.", "exclcustom": "Також виключити (один шлях на рядок):", "exclexp": "Ці системні папки завжди пропускаються під час сканування.", "filesafter": "Файли з'являться після завершення поточного сканування.", "filescap": "Показ обмежено 1000 найбільшими файлами цієї папки.", "filterexp": "Фільтрує лише показ, для читабельності. Не пришвидшує сканування: розміри завжди обчислюються повністю. Можна змінити будь-коли.", "hiddenmsg": "{n} елемент(ів) нижче порога приховано фільтром показу.", "modalhint": "Оберіть папку для аналізу, потім задайте глибину і показ.", "pickerfail": "Не вдалося відкрити системний вибір папки на цьому комп'ютері. Введіть шлях вручну.", "queuedmsg": "Папку поставлено в чергу ({n} очікує). Буде проскановано після поточного сканування."},ro:{"sub": "Analizor de spațiu pe disc", "newscan": "Scanare nouă", "total": "Total:", "newanalysis": "Analiză nouă", "folderlabel": "Dosarul de analizat", "browse": "Răsfoiește...", "quickaccess": "Acces rapid", "recent": "Recente", "drives": "Unități", "depthlabel": "Adâncimea explorării:", "unlimited": "Nelimitat (scanează tot, poate dura foarte mult)", "displaylabel": "Afișare: ascunde elementele mici", "showall": "Arată tot", "analyze": "Analizează", "quit": "Închide serverul", "legend": "Legendă", "help": "Ajutor", "stop": "Oprește", "choose": "Alege acest dosar", "cancel": "Anulează", "sortby": "Sortează", "byname": "nume", "bysize": "mărime", "scangrey": "Scanează dosarele gri", "leggreen": "Dosar scanat, dimensiune cunoscuta", "legorange": "Dosar planificat in scanarea curenta", "legblue": "Dosar in coada (scanat dupa scanarea curenta)", "legspin": "Dosar in curs de scanare", "leggrey": "Nescanat, exclus, jonctiune sau protejat. Clic pentru a-l scana (2 niveluri)", "legfile": "Fisier", "legbar": "Ponderea in dimensiunea dosarului curent", "helpnav": "Navigare: clic pe un dosar pentru a intra, firimiturile de sus pentru a urca.", "helpgrey": "Dosare gri: clic pentru a le scana (inca 2 niveluri). In timpul unei scanari intra in coada (albastru) si sunt scanate la final.", "helpfiles": "Fisiere: afisate cand deschizi un dosar, sortate dupa dimensiune, limitate la cele mai mari 1000.", "helpfilter": "Filtru de afisare (buton Scanare noua): ascunde elementele mici pentru lizibilitate, fara a schimba scanarea.", "helpone": "O scanare pe rand: doua scanari de disc in paralel s-ar incetini reciproc.", "noanalysis": "Nicio analiza", "theme": "Tema deschisa / inchisa", "ready": "Gata.", "launching": "Se porneste scanarea...", "done": "Scanare finalizata. Navigheaza liber.", "inprogress": "In curs:", "scanningtag": "se scaneaza", "parent": ".. (dosar parinte)", "scandots": "Se scaneaza...", "empty": "Dosar gol", "loadingfiles": "Se incarca fisierele...", "excluded": "exclus", "junction": "jonctiune", "unscanned": "nescanat", "protected": "protejat", "error": "eroare", "scanshort": "scanare...", "queued": "in coada", "levels": "nivel(uri)", "hide1": "Ascunde sub 1 MB", "hide100": "Ascunde sub 100 MB", "hide1g": "Ascunde sub 1 GB", "genby": "Generat de:", "invalidpath": "Cale invalida sau inaccesibila.", "windowopen": "Fereastra deschisa...", "interrupted": "Scanare intrerupta.", "enterpath": "Introduceti o cale.", "serverstopped": "Server oprit. Puteti inchide aceasta fila.", "you": "(dvs.)", "network": "(retea)", "drivefree": "liberi", "removeone": "Elimina", "clearhist": "Sterge istoricul", "exclusions": "Excluderi", "scandepthfull": "Adancime: nelimitata", "scandepthn": "Adancime: {n}", "browsetitle": "Rasfoieste", "upfolder": "Dosar parinte", "clickscan": "Clic pentru a scana acest dosar (2 niveluri)", "depthhigh": "Detaliat: multe niveluri preincarcate, mai greu de afisat.", "depthlow": "Usor si rapid de afisat: doar primele niveluri sunt incarcate. Coborati dand clic pe un dosar.", "depthmid": "Echilibru bun: mai multe niveluri vizibile deodata, tot fluid.", "depthtail": "Dimensiunile sunt mereu exacte; dincolo de aceasta adancime, dati clic pe un dosar gri pentru a-l explora.", "depthunl": "Parcurge si afiseaza intregul arbore. Poate fi foarte lung si greu pe un disc mare.", "exclcustom": "Exclude si (o cale pe linie):", "exclexp": "Aceste dosare de sistem sunt mereu sarite in timpul scanarii.", "filesafter": "Fisierele vor aparea la finalul scanarii curente.", "filescap": "Afisare limitata la cele mai mari 1000 de fisiere din acest dosar.", "filterexp": "Filtreaza doar afisarea, pentru lizibilitate. Nu accelereaza scanarea: dimensiunile sunt mereu calculate complet. Se poate schimba oricand.", "hiddenmsg": "{n} element(e) sub prag ascuns(e) de filtrul de afisare.", "modalhint": "Alegeti dosarul de analizat, apoi setati adancimea si afisarea.", "pickerfail": "Selectorul nativ de dosare nu s-a putut deschide pe aceasta masina. Tastati calea direct.", "queuedmsg": "Dosar pus in coada ({n} in asteptare). Scanat la finalul scanarii curente."},cs:{"sub": "Analyzátor využití disku", "newscan": "Nové skenování", "total": "Celkem:", "newanalysis": "Nová analýza", "folderlabel": "Složka k analýze", "browse": "Procházet...", "quickaccess": "Rychlý přístup", "recent": "Nedávné", "drives": "Disky", "depthlabel": "Hloubka průzkumu:", "unlimited": "Neomezeně (skenuje vše, může trvat velmi dlouho)", "displaylabel": "Zobrazení: skrýt malé položky", "showall": "Zobrazit vše", "analyze": "Analyzovat", "quit": "Ukončit server", "legend": "Legenda", "help": "Nápověda", "stop": "Přerušit", "choose": "Vybrat tuto složku", "cancel": "Zrušit", "sortby": "Řadit", "byname": "název", "bysize": "velikost", "scangrey": "Skenovat šedé složky", "leggreen": "Naskenovana slozka, velikost znama", "legorange": "Slozka naplanovana v aktualnim skenovani", "legblue": "Slozka ve fronte (skenovana po aktualnim skenovani)", "legspin": "Slozka se skenuje", "leggrey": "Neskenovano, vylouceno, spojeni nebo chraneno. Kliknutim naskenujete (2 urovne)", "legfile": "Soubor", "legbar": "Podil na velikosti aktualni slozky", "helpnav": "Navigace: kliknutim na slozku vstoupite, drobenkova navigace nahore pro navrat vyse.", "helpgrey": "Sede slozky: kliknutim je naskenujete (dalsi 2 urovne). Behem skenovani se zaradi do fronty (modra) a skenuji se na konci.", "helpfiles": "Soubory: zobrazeny pri otevreni slozky, serazeny podle velikosti, omezeny na 1000 nejvetsich.", "helpfilter": "Filtr zobrazeni (tlacitko Nove skenovani): skryje male polozky pro citelnost, bez zmeny skenovani.", "helpone": "Jedno skenovani najednou: dve soubezna skenovani disku by se navzajem zpomalovala.", "noanalysis": "Zadna analyza", "theme": "Svetly / tmavy motiv", "ready": "Pripraveno.", "launching": "Spousteni skenovani...", "done": "Skenovani dokonceno. Volne prochazejte.", "inprogress": "Probiha:", "scanningtag": "skenuje se", "parent": ".. (nadrazena slozka)", "scandots": "Skenovani...", "empty": "Prazdna slozka", "loadingfiles": "Nacitani souboru...", "excluded": "vylouceno", "junction": "spojeni", "unscanned": "neskenovano", "protected": "chraneno", "error": "chyba", "scanshort": "sken...", "queued": "ve fronte", "levels": "uroven/urovni", "hide1": "Skryt pod 1 MB", "hide100": "Skryt pod 100 MB", "hide1g": "Skryt pod 1 GB", "genby": "Vygeneroval:", "invalidpath": "Neplatna nebo nepristupna cesta.", "windowopen": "Okno otevreno...", "interrupted": "Skenovani preruseno.", "enterpath": "Zadejte cestu.", "serverstopped": "Server zastaven. Tuto kartu muzete zavrit.", "you": "(vy)", "network": "(sit)", "drivefree": "volne", "removeone": "Odebrat", "clearhist": "Vymazat historii", "exclusions": "Vylouceni", "scandepthfull": "Hloubka: neomezena", "scandepthn": "Hloubka: {n}", "browsetitle": "Prochazet", "upfolder": "Nadrazena slozka", "clickscan": "Kliknutim naskenujete tuto slozku (2 urovne)", "depthhigh": "Podrobne: mnoho urovni prednacteno, narocnejsi zobrazeni.", "depthlow": "Lehke a rychle zobrazeni: nacteny jen prvni urovne. Sestupujte kliknutim na slozku.", "depthmid": "Dobry kompromis: nekolik urovni videt najednou, stale plynule.", "depthtail": "Velikosti jsou vzdy presne; za touto hloubkou kliknete na sedou slozku pro prozkoumani.", "depthunl": "Projde a zobrazi cely strom. Na velkem disku muze byt velmi dlouhe a narocne.", "exclcustom": "Vyloucit take (jedna cesta na radek):", "exclexp": "Tyto systemove slozky jsou pri skenovani vzdy preskoceny.", "filesafter": "Soubory se zobrazi po dokonceni aktualniho skenovani.", "filescap": "Zobrazeni omezeno na 1000 nejvetsich souboru v teto slozce.", "filterexp": "Filtruje pouze zobrazeni, pro citelnost. Nezrychluje skenovani: velikosti se vzdy pocitaji cele. Lze zmenit kdykoli.", "hiddenmsg": "{n} polozek pod prahem skryto filtrem zobrazeni.", "modalhint": "Vyberte slozku k analyze, pote nastavte hloubku a zobrazeni.", "pickerfail": "Nativni vyber slozky se na tomto pocitaci nepodarilo otevrit. Zadejte cestu rucne.", "queuedmsg": "Slozka zarazena do fronty ({n} ceka). Naskenuje se po aktualnim skenovani."},el:{"sub": "Αναλυτής χώρου δίσκου", "newscan": "Νέα σάρωση", "total": "Σύνολο:", "newanalysis": "Νέα ανάλυση", "folderlabel": "Φάκελος προς ανάλυση", "browse": "Αναζήτηση...", "quickaccess": "Γρήγορη πρόσβαση", "recent": "Πρόσφατα", "drives": "Μονάδες", "depthlabel": "Βάθος εξερεύνησης:", "unlimited": "Απεριόριστο (σαρώνει τα πάντα, μπορεί να διαρκέσει πολύ)", "displaylabel": "Εμφάνιση: απόκρυψη μικρών στοιχείων", "showall": "Εμφάνιση όλων", "analyze": "Ανάλυση", "quit": "Τερματισμός διακομιστή", "legend": "Υπόμνημα", "help": "Βοήθεια", "stop": "Διακοπή", "choose": "Επιλογή αυτού του φακέλου", "cancel": "Άκυρο", "sortby": "Ταξινόμηση", "byname": "όνομα", "bysize": "μέγεθος", "scangrey": "Σάρωση γκρι φακέλων", "leggreen": "Σαρωμένος φάκελος, γνωστό μέγεθος", "legorange": "Φάκελος προγραμματισμένος στην τρέχουσα σάρωση", "legblue": "Φάκελος σε ουρά (σαρώνεται μετά την τρέχουσα σάρωση)", "legspin": "Φάκελος υπό σάρωση", "leggrey": "Μη σαρωμένος, εξαιρεμένος, σύνδεσμος ή προστατευμένος. Κάντε κλικ για σάρωση (2 επίπεδα)", "legfile": "Αρχείο", "legbar": "Μερίδιο στο μέγεθος του τρέχοντος φακέλου", "helpnav": "Πλοήγηση: κάντε κλικ σε φάκελο για είσοδο, διαδρομή στην κορυφή για άνοδο.", "helpgrey": "Γκρι φάκελοι: κάντε κλικ για σάρωση (2 ακόμη επίπεδα). Κατά τη σάρωση μπαίνουν σε ουρά (μπλε) και σαρώνονται στο τέλος.", "helpfiles": "Αρχεία: εμφανίζονται όταν ανοίγετε φάκελο, ταξινομημένα κατά μέγεθος, έως τα 1000 μεγαλύτερα.", "helpfilter": "Φίλτρο εμφάνισης (κουμπί Νέα σάρωση): κρύβει μικρά στοιχεία για αναγνωσιμότητα, χωρίς αλλαγή της σάρωσης.", "helpone": "Μία σάρωση τη φορά: δύο παράλληλες σαρώσεις δίσκου θα καθυστερούσαν η μία την άλλη.", "noanalysis": "Καμία ανάλυση", "theme": "Φωτεινό / σκούρο θέμα", "ready": "Έτοιμο.", "launching": "Έναρξη σάρωσης...", "done": "Η σάρωση ολοκληρώθηκε. Περιηγηθείτε ελεύθερα.", "inprogress": "Σε εξέλιξη:", "scanningtag": "σάρωση", "parent": ".. (γονικός φάκελος)", "scandots": "Σάρωση...", "empty": "Κενός φάκελος", "loadingfiles": "Φόρτωση αρχείων...", "excluded": "εξαιρέθηκε", "junction": "σύνδεσμος", "unscanned": "μη σαρωμένο", "protected": "προστατευμένο", "error": "σφάλμα", "scanshort": "σάρωση...", "queued": "σε ουρά", "levels": "επίπεδο/α", "hide1": "Απόκρυψη κάτω από 1 MB", "hide100": "Απόκρυψη κάτω από 100 MB", "hide1g": "Απόκρυψη κάτω από 1 GB", "genby": "Δημιουργήθηκε από:", "invalidpath": "Μη έγκυρη ή μη προσβάσιμη διαδρομή.", "windowopen": "Το παράθυρο άνοιξε...", "interrupted": "Η σάρωση διακόπηκε.", "enterpath": "Εισαγάγετε μια διαδρομή.", "serverstopped": "Ο διακομιστής σταμάτησε. Μπορείτε να κλείσετε αυτήν την καρτέλα.", "you": "(εσείς)", "network": "(δίκτυο)", "drivefree": "ελεύθερα", "removeone": "Αφαίρεση", "clearhist": "Εκκαθάριση ιστορικού", "exclusions": "Εξαιρέσεις", "scandepthfull": "Βάθος: απεριόριστο", "scandepthn": "Βάθος: {n}", "browsetitle": "Αναζήτηση", "upfolder": "Γονικός φάκελος", "clickscan": "Κάντε κλικ για σάρωση αυτού του φακέλου (2 επίπεδα)", "depthhigh": "Λεπτομερές: πολλά επίπεδα προφορτωμένα, βαρύτερη εμφάνιση.", "depthlow": "Ελαφρύ και γρήγορο: φορτώνονται μόνο τα πρώτα επίπεδα. Κατεβείτε κάνοντας κλικ σε φάκελο.", "depthmid": "Καλή ισορροπία: αρκετά επίπεδα ορατά ταυτόχρονα, παραμένει ομαλό.", "depthtail": "Τα μεγέθη είναι πάντα ακριβή· πέρα από αυτό το βάθος, κάντε κλικ σε γκρι φάκελο για εξερεύνηση.", "depthunl": "Διατρέχει και εμφανίζει όλο το δέντρο. Μπορεί να είναι πολύ αργό και βαρύ σε μεγάλο δίσκο.", "exclcustom": "Εξαίρεση επίσης (μία διαδρομή ανά γραμμή):", "exclexp": "Αυτοί οι φάκελοι συστήματος παραλείπονται πάντα κατά τη σάρωση.", "filesafter": "Τα αρχεία θα εμφανιστούν στο τέλος της τρέχουσας σάρωσης.", "filescap": "Εμφάνιση περιορισμένη στα 1000 μεγαλύτερα αρχεία αυτού του φακέλου.", "filterexp": "Φιλτράρει μόνο την εμφάνιση, για αναγνωσιμότητα. Δεν επιταχύνει τη σάρωση: τα μεγέθη υπολογίζονται πάντα πλήρως. Αλλάζει οποτεδήποτε.", "hiddenmsg": "{n} στοιχείο(α) κάτω από το όριο κρύφτηκαν από το φίλτρο εμφάνισης.", "modalhint": "Επιλέξτε τον φάκελο προς ανάλυση, μετά ορίστε βάθος και εμφάνιση.", "pickerfail": "Ο εγγενής επιλογέας φακέλων δεν άνοιξε σε αυτό το μηχάνημα. Πληκτρολογήστε τη διαδρομή απευθείας.", "queuedmsg": "Φάκελος σε ουρά ({n} σε αναμονή). Σαρώνεται στο τέλος της τρέχουσας σάρωσης."},sv:{"sub": "Diskutrymmesanalys", "newscan": "Ny skanning", "total": "Totalt:", "newanalysis": "Ny analys", "folderlabel": "Mapp att analysera", "browse": "Bläddra...", "quickaccess": "Snabbåtkomst", "recent": "Senaste", "drives": "Enheter", "depthlabel": "Utforskningsdjup:", "unlimited": "Obegränsat (skannar allt, kan ta mycket lång tid)", "displaylabel": "Visning: dölj små objekt", "showall": "Visa allt", "analyze": "Analysera", "quit": "Avsluta servern", "legend": "Teckenförklaring", "help": "Hjälp", "stop": "Avbryt", "choose": "Välj denna mapp", "cancel": "Avbryt", "sortby": "Sortera", "byname": "namn", "bysize": "storlek", "scangrey": "Skanna grå mappar", "leggreen": "Skannad mapp, storlek kand", "legorange": "Mapp planerad i aktuell skanning", "legblue": "Mapp i ko (skannas efter aktuell skanning)", "legspin": "Mapp skannas", "leggrey": "Inte skannad, utesluten, lankning eller skyddad. Klicka for att skanna (2 nivaer)", "legfile": "Fil", "legbar": "Andel av aktuell mapps storlek", "helpnav": "Navigera: klicka pa en mapp for att oppna, sokvagen hogst upp for att ga upp.", "helpgrey": "Graa mappar: klicka for att skanna (2 nivaer till). Under en skanning koas de (bla) och skannas i slutet.", "helpfiles": "Filer: visas nar du oppnar en mapp, sorterade efter storlek, begransade till de 1000 storsta.", "helpfilter": "Visningsfilter (knappen Ny skanning): doljer sma objekt for lasbarhet, utan att andra skanningen.", "helpone": "En skanning i taget: tva parallella diskskanningar skulle bromsa varandra.", "noanalysis": "Ingen analys", "theme": "Ljust / morkt tema", "ready": "Klar.", "launching": "Startar skanning...", "done": "Skanning klar. Navigera fritt.", "inprogress": "Pagar:", "scanningtag": "skannar", "parent": ".. (overordnad mapp)", "scandots": "Skannar...", "empty": "Tom mapp", "loadingfiles": "Laddar filer...", "excluded": "utesluten", "junction": "lankning", "unscanned": "inte skannad", "protected": "skyddad", "error": "fel", "scanshort": "skanning...", "queued": "i ko", "levels": "niva(er)", "hide1": "Dolj under 1 MB", "hide100": "Dolj under 100 MB", "hide1g": "Dolj under 1 GB", "genby": "Genererad av:", "invalidpath": "Ogiltig eller otillganglig sokvag.", "windowopen": "Fonster oppnat...", "interrupted": "Skanning avbruten.", "enterpath": "Ange en sokvag.", "serverstopped": "Servern stoppad. Du kan stanga den har fliken.", "you": "(du)", "network": "(natverk)", "drivefree": "ledigt", "removeone": "Ta bort", "clearhist": "Rensa historik", "exclusions": "Undantag", "scandepthfull": "Djup: obegransat", "scandepthn": "Djup: {n}", "browsetitle": "Bladdra", "upfolder": "Overordnad mapp", "clickscan": "Klicka for att skanna denna mapp (2 nivaer)", "depthhigh": "Detaljerat: manga nivaer forladdade, tyngre att visa.", "depthlow": "Latt och snabbt: bara de forsta nivaerna laddas. Ga djupare genom att klicka pa en mapp.", "depthmid": "Bra balans: flera nivaer synliga samtidigt, fortfarande smidigt.", "depthtail": "Storlekar ar alltid exakta; bortom detta djup, klicka pa en gra mapp for att utforska den.", "depthunl": "Gar igenom och visar hela tradet. Kan ta mycket lang tid och bli tungt pa en stor disk.", "exclcustom": "Uteslut aven (en sokvag per rad):", "exclexp": "Dessa systemmappar hoppas alltid over under skanningen.", "filesafter": "Filer visas nar den aktuella skanningen ar klar.", "filescap": "Visning begransad till de 1000 storsta filerna i denna mapp.", "filterexp": "Filtrerar bara visningen, for lasbarhet. Snabbar inte upp skanningen: storlekar beraknas alltid fullt ut. Kan andras nar som helst.", "hiddenmsg": "{n} objekt under troskeln dolda av visningsfiltret.", "modalhint": "Valj mappen att analysera, stall sedan in djup och visning.", "pickerfail": "Den inbyggda mappvaljaren kunde inte oppnas pa denna dator. Skriv sokvagen direkt.", "queuedmsg": "Mapp i ko ({n} vantar). Skannas i slutet av aktuell skanning."},hu:{"sub": "Lemezhasználat-elemző", "newscan": "Új vizsgálat", "total": "Összesen:", "newanalysis": "Új elemzés", "folderlabel": "Elemzendő mappa", "browse": "Tallózás...", "quickaccess": "Gyors elérés", "recent": "Legutóbbi", "drives": "Meghajtók", "depthlabel": "Feltárási mélység:", "unlimited": "Korlátlan (mindent átvizsgál, nagyon hosszú lehet)", "displaylabel": "Megjelenítés: kis elemek elrejtése", "showall": "Összes megjelenítése", "analyze": "Elemzés", "quit": "Szerver leállítása", "legend": "Jelmagyarázat", "help": "Súgó", "stop": "Megszakítás", "choose": "Mappa kiválasztása", "cancel": "Mégse", "sortby": "Rendezés", "byname": "név", "bysize": "méret", "scangrey": "Szürke mappák vizsgálata", "leggreen": "Vizsgalt mappa, meret ismert", "legorange": "A jelenlegi vizsgalatban tervezett mappa", "legblue": "Sorban allo mappa (a jelenlegi vizsgalat utan)", "legspin": "Vizsgalat alatt allo mappa", "leggrey": "Nem vizsgalt, kizart, csatolas vagy vedett. Kattintson a vizsgalathoz (2 szint)", "legfile": "Fajl", "legbar": "Reszesedes a jelenlegi mappa meretebol", "helpnav": "Navigacio: kattintson egy mappara a belepeshez, felul a morzsameny a feljebb lepeshez.", "helpgrey": "Szurke mappak: kattintson a vizsgalathoz (2 tovabbi szint). Vizsgalat kozben sorba allnak (kek), es a vegen vizsgaljak oket.", "helpfiles": "Fajlok: mappa megnyitasakor jelennek meg, meret szerint rendezve, a legnagyobb 1000-re korlatozva.", "helpfilter": "Megjelenitesi szuro (Uj vizsgalat gomb): elrejti a kis elemeket az olvashatosagert, a vizsgalat valtoztatasa nelkul.", "helpone": "Egyszerre egy vizsgalat: ket parhuzamos lemezvizsgalat lassitana egymast.", "noanalysis": "Nincs elemzes", "theme": "Vilagos / sotet tema", "ready": "Kesz.", "launching": "Vizsgalat inditasa...", "done": "Vizsgalat kesz. Navigaljon szabadon.", "inprogress": "Folyamatban:", "scanningtag": "vizsgalat", "parent": ".. (szulomappa)", "scandots": "Vizsgalat...", "empty": "Ures mappa", "loadingfiles": "Fajlok betoltese...", "excluded": "kizarva", "junction": "csatolas", "unscanned": "nem vizsgalt", "protected": "vedett", "error": "hiba", "scanshort": "vizsgalat...", "queued": "sorban", "levels": "szint", "hide1": "1 MB alattiak elrejtese", "hide100": "100 MB alattiak elrejtese", "hide1g": "1 GB alattiak elrejtese", "genby": "Keszitette:", "invalidpath": "Ervenytelen vagy elerhetetlen utvonal.", "windowopen": "Ablak megnyitva...", "interrupted": "Vizsgalat megszakitva.", "enterpath": "Adjon meg egy utvonalat.", "serverstopped": "Szerver leallitva. Bezarhatja ezt a lapot.", "you": "(on)", "network": "(halozat)", "drivefree": "szabad", "removeone": "Eltavolitas", "clearhist": "Elozmenyek torlese", "exclusions": "Kizarasok", "scandepthfull": "Melyseg: korlatlan", "scandepthn": "Melyseg: {n}", "browsetitle": "Tallozas", "upfolder": "Szulomappa", "clickscan": "Kattintson a mappa vizsgalatahoz (2 szint)", "depthhigh": "Reszletes: sok szint elore betoltve, nehezebb megjeleniteni.", "depthlow": "Konnyu es gyors: csak az elso szintek toltodnek be. Mappara kattintva lepjen lejjebb.", "depthmid": "Jo egyensuly: tobb szint lathato egyszerre, tovabbra is gordulekeny.", "depthtail": "A meretek mindig pontosak; e melysegen tul kattintson szurke mappara a felfedezeshez.", "depthunl": "Bejarja es megjeleniti a teljes fat. Nagy lemezen nagyon hosszu es nehez lehet.", "exclcustom": "Kizaras meg (soronkent egy utvonal):", "exclexp": "Ezeket a rendszermappakat a vizsgalat mindig kihagyja.", "filesafter": "A fajlok az aktualis vizsgalat vegen jelennek meg.", "filescap": "A megjelenites e mappa 1000 legnagyobb fajljara korlatozva.", "filterexp": "Csak a megjelenitest szuri, az olvashatosagert. Nem gyorsitja a vizsgalatot: a meretek mindig teljesen kiszamitasra kerulnek. Barmikor modosithato.", "hiddenmsg": "{n} elem a kuszob alatt elrejtve a megjelenitesi szurovel.", "modalhint": "Valassza ki az elemzendo mappat, majd allitsa be a melyseget es a megjelenitest.", "pickerfail": "A nativ mappavalaszto nem nyithato meg ezen a gepen. Irja be az utvonalat kozvetlenul.", "queuedmsg": "Mappa sorba allitva ({n} var). Az aktualis vizsgalat vegen vizsgaljuk."},fa:{"sub": "تحلیلگر فضای دیسک", "newscan": "اسکن جدید", "total": "مجموع:", "newanalysis": "تحلیل جدید", "folderlabel": "پوشه برای تحلیل", "browse": "مرور...", "quickaccess": "دسترسی سریع", "recent": "اخیر", "drives": "درایوها", "depthlabel": "عمق کاوش:", "unlimited": "نامحدود (همه را اسکن می‌کند، ممکن است بسیار طولانی باشد)", "displaylabel": "نمایش: پنهان کردن موارد کوچک", "showall": "نمایش همه", "analyze": "تحلیل", "quit": "بستن سرور", "legend": "راهنما", "help": "کمک", "stop": "توقف", "choose": "انتخاب این پوشه", "cancel": "لغو", "sortby": "مرتب‌سازی", "byname": "نام", "bysize": "اندازه", "scangrey": "اسکن پوشه‌های خاکستری", "leggreen": "پوشه اسکن‌شده، اندازه مشخص", "legorange": "پوشه برنامه‌ریزی‌شده در اسکن جاری", "legblue": "پوشه در صف (پس از اسکن جاری اسکن می‌شود)", "legspin": "پوشه در حال اسکن", "leggrey": "اسکن‌نشده، مستثنی، پیوند یا محافظت‌شده. برای اسکن کلیک کنید (۲ سطح)", "legfile": "فایل", "legbar": "سهم از اندازه پوشه جاری", "helpnav": "پیمایش: برای ورود روی پوشه کلیک کنید، مسیر بالا برای بالا رفتن.", "helpgrey": "پوشه‌های خاکستری: برای اسکن کلیک کنید (۲ سطح بیشتر). هنگام اسکن در صف قرار می‌گیرند (آبی) و در پایان اسکن می‌شوند.", "helpfiles": "فایل‌ها: هنگام باز کردن پوشه نمایش داده می‌شوند، بر اساس اندازه مرتب، محدود به ۱۰۰۰ فایل بزرگ‌تر.", "helpfilter": "فیلتر نمایش (دکمه اسکن جدید): موارد کوچک را برای خوانایی پنهان می‌کند، بدون تغییر اسکن.", "helpone": "هر بار یک اسکن: دو اسکن هم‌زمان دیسک یکدیگر را کند می‌کنند.", "noanalysis": "بدون تحلیل", "theme": "پوسته روشن / تیره", "ready": "آماده.", "launching": "در حال شروع اسکن...", "done": "اسکن کامل شد. آزادانه پیمایش کنید.", "inprogress": "در حال انجام:", "scanningtag": "در حال اسکن", "parent": ".. (پوشه والد)", "scandots": "در حال اسکن...", "empty": "پوشه خالی", "loadingfiles": "در حال بارگذاری فایل‌ها...", "excluded": "مستثنی", "junction": "پیوند", "unscanned": "اسکن‌نشده", "protected": "محافظت‌شده", "error": "خطا", "scanshort": "اسکن...", "queued": "در صف", "levels": "سطح", "hide1": "پنهان کردن زیر ۱ مگابایت", "hide100": "پنهان کردن زیر ۱۰۰ مگابایت", "hide1g": "پنهان کردن زیر ۱ گیگابایت", "genby": "تولیدشده توسط:", "invalidpath": "مسیر نامعتبر یا غیرقابل‌دسترس.", "windowopen": "پنجره باز شد...", "interrupted": "اسکن متوقف شد.", "enterpath": "لطفاً یک مسیر وارد کنید.", "serverstopped": "سرور متوقف شد. می‌توانید این زبانه را ببندید.", "you": "(شما)", "network": "(شبکه)", "drivefree": "آزاد", "removeone": "حذف", "clearhist": "پاک کردن سابقه", "exclusions": "استثناها", "scandepthfull": "عمق: نامحدود", "scandepthn": "عمق: {n}", "browsetitle": "مرور", "upfolder": "پوشه والد", "clickscan": "برای اسکن این پوشه کلیک کنید (۲ سطح)", "depthhigh": "تفصیلی: سطوح زیاد از پیش بارگذاری شده، نمایش سنگین‌تر.", "depthlow": "سبک و سریع: فقط سطوح اول بارگذاری می‌شوند. با کلیک روی پوشه عمیق‌تر بروید.", "depthmid": "تعادل خوب: چند سطح هم‌زمان قابل مشاهده، همچنان روان.", "depthtail": "اندازه‌ها همیشه دقیق‌اند؛ فراتر از این عمق، روی پوشه خاکستری کلیک کنید تا کاوش شود.", "depthunl": "کل درخت را پیمایش و نمایش می‌دهد. روی دیسک بزرگ می‌تواند بسیار طولانی و سنگین باشد.", "exclcustom": "همچنین مستثنی کنید (هر خط یک مسیر):", "exclexp": "این پوشه‌های سیستمی همیشه در اسکن نادیده گرفته می‌شوند.", "filesafter": "فایل‌ها در پایان اسکن جاری نمایش داده می‌شوند.", "filescap": "نمایش محدود به ۱۰۰۰ فایل بزرگ‌تر این پوشه.", "filterexp": "فقط نمایش را فیلتر می‌کند، برای خوانایی. اسکن را سریع‌تر نمی‌کند: اندازه‌ها همیشه کامل محاسبه می‌شوند. هر زمان قابل تغییر.", "hiddenmsg": "{n} مورد زیر آستانه توسط فیلتر نمایش پنهان شد.", "modalhint": "پوشه مورد تحلیل را انتخاب کنید، سپس عمق و نمایش را تنظیم کنید.", "pickerfail": "انتخابگر پوشه بومی روی این دستگاه باز نشد. مسیر را مستقیم وارد کنید.", "queuedmsg": "پوشه در صف قرار گرفت ({n} در انتظار). در پایان اسکن جاری اسکن می‌شود."},th:{"sub": "เครื่องวิเคราะห์การใช้ดิสก์", "newscan": "สแกนใหม่", "total": "รวม:", "newanalysis": "การวิเคราะห์ใหม่", "folderlabel": "โฟลเดอร์ที่จะวิเคราะห์", "browse": "เรียกดู...", "quickaccess": "เข้าถึงด่วน", "recent": "ล่าสุด", "drives": "ไดรฟ์", "depthlabel": "ความลึกในการสำรวจ:", "unlimited": "ไม่จำกัด (สแกนทั้งหมด อาจใช้เวลานานมาก)", "displaylabel": "แสดง: ซ่อนรายการเล็ก", "showall": "แสดงทั้งหมด", "analyze": "วิเคราะห์", "quit": "ปิดเซิร์ฟเวอร์", "legend": "คำอธิบาย", "help": "ช่วยเหลือ", "stop": "หยุด", "choose": "เลือกโฟลเดอร์นี้", "cancel": "ยกเลิก", "sortby": "เรียง", "byname": "ชื่อ", "bysize": "ขนาด", "scangrey": "สแกนโฟลเดอร์สีเทา", "leggreen": "โฟลเดอร์ที่สแกนแล้ว ทราบขนาด", "legorange": "โฟลเดอร์ที่วางแผนไว้ในการสแกนปัจจุบัน", "legblue": "โฟลเดอร์ในคิว (สแกนหลังการสแกนปัจจุบัน)", "legspin": "โฟลเดอร์กำลังถูกสแกน", "leggrey": "ยังไม่สแกน ยกเว้น จุดเชื่อม หรือได้รับการป้องกัน คลิกเพื่อสแกน (2 ระดับ)", "legfile": "ไฟล์", "legbar": "สัดส่วนของขนาดโฟลเดอร์ปัจจุบัน", "helpnav": "นำทาง: คลิกโฟลเดอร์เพื่อเข้า เส้นทางด้านบนเพื่อขึ้น", "helpgrey": "โฟลเดอร์สีเทา: คลิกเพื่อสแกน (อีก 2 ระดับ) ระหว่างการสแกนจะเข้าคิว (สีน้ำเงิน) และสแกนตอนท้าย", "helpfiles": "ไฟล์: แสดงเมื่อเปิดโฟลเดอร์ เรียงตามขนาด จำกัด 1000 ไฟล์ใหญ่สุด", "helpfilter": "ตัวกรองการแสดง (ปุ่มสแกนใหม่): ซ่อนรายการเล็กเพื่อให้อ่านง่าย โดยไม่เปลี่ยนการสแกน", "helpone": "สแกนครั้งละหนึ่ง: การสแกนดิสก์สองรายการพร้อมกันจะทำให้ช้าลง", "noanalysis": "ไม่มีการวิเคราะห์", "theme": "ธีมสว่าง / มืด", "ready": "พร้อม", "launching": "กำลังเริ่มสแกน...", "done": "สแกนเสร็จ นำทางได้อย่างอิสระ", "inprogress": "กำลังดำเนินการ:", "scanningtag": "กำลังสแกน", "parent": ".. (โฟลเดอร์แม่)", "scandots": "กำลังสแกน...", "empty": "โฟลเดอร์ว่าง", "loadingfiles": "กำลังโหลดไฟล์...", "excluded": "ยกเว้น", "junction": "จุดเชื่อม", "unscanned": "ยังไม่สแกน", "protected": "ได้รับการป้องกัน", "error": "ข้อผิดพลาด", "scanshort": "สแกน...", "queued": "อยู่ในคิว", "levels": "ระดับ", "hide1": "ซ่อนต่ำกว่า 1 MB", "hide100": "ซ่อนต่ำกว่า 100 MB", "hide1g": "ซ่อนต่ำกว่า 1 GB", "genby": "สร้างโดย:", "invalidpath": "เส้นทางไม่ถูกต้องหรือเข้าถึงไม่ได้", "windowopen": "เปิดหน้าต่างแล้ว...", "interrupted": "สแกนถูกขัดจังหวะ", "enterpath": "โปรดป้อนเส้นทาง", "serverstopped": "เซิร์ฟเวอร์หยุดแล้ว คุณสามารถปิดแท็บนี้ได้", "you": "(คุณ)", "network": "(เครือข่าย)", "drivefree": "ว่าง", "removeone": "ลบ", "clearhist": "ล้างประวัติ", "exclusions": "การยกเว้น", "scandepthfull": "ความลึก: ไม่จำกัด", "scandepthn": "ความลึก: {n}", "browsetitle": "เรียกดู", "upfolder": "โฟลเดอร์แม่", "clickscan": "คลิกเพื่อสแกนโฟลเดอร์นี้ (2 ระดับ)", "depthhigh": "ละเอียด: โหลดหลายระดับล่วงหน้า แสดงผลหนักขึ้น", "depthlow": "เบาและแสดงเร็ว: โหลดเฉพาะระดับแรก ลงลึกโดยคลิกโฟลเดอร์", "depthmid": "สมดุลดี: เห็นหลายระดับพร้อมกัน ยังคงลื่นไหล", "depthtail": "ขนาดแม่นยำเสมอ เกินความลึกนี้ให้คลิกโฟลเดอร์สีเทาเพื่อสำรวจ", "depthunl": "เดินและแสดงทั้งต้นไม้ อาจนานและหนักมากบนดิสก์ใหญ่", "exclcustom": "ยกเว้นเพิ่ม (หนึ่งเส้นทางต่อบรรทัด):", "exclexp": "โฟลเดอร์ระบบเหล่านี้จะถูกข้ามเสมอระหว่างสแกน", "filesafter": "ไฟล์จะแสดงเมื่อการสแกนปัจจุบันเสร็จ", "filescap": "แสดงจำกัด 1000 ไฟล์ใหญ่สุดในโฟลเดอร์นี้", "filterexp": "กรองเฉพาะการแสดงผลเพื่อให้อ่านง่าย ไม่ทำให้สแกนเร็วขึ้น: ขนาดคำนวณเต็มเสมอ เปลี่ยนได้ทุกเมื่อ", "hiddenmsg": "{n} รายการต่ำกว่าเกณฑ์ถูกซ่อนโดยตัวกรองการแสดงผล", "modalhint": "เลือกโฟลเดอร์ที่จะวิเคราะห์ แล้วตั้งความลึกและการแสดงผล", "pickerfail": "ไม่สามารถเปิดตัวเลือกโฟลเดอร์ของระบบบนเครื่องนี้ พิมพ์เส้นทางโดยตรง", "queuedmsg": "โฟลเดอร์เข้าคิวแล้ว ({n} รอ) จะสแกนเมื่อการสแกนปัจจุบันเสร็จ"},ms:{"sub": "Penganalisis Penggunaan Cakera", "newscan": "Imbasan baharu", "total": "Jumlah:", "newanalysis": "Analisis baharu", "folderlabel": "Folder untuk dianalisis", "browse": "Semak imbas...", "quickaccess": "Akses pantas", "recent": "Terkini", "drives": "Pemacu", "depthlabel": "Kedalaman penerokaan:", "unlimited": "Tanpa had (imbas semua, mungkin sangat lama)", "displaylabel": "Paparan: sembunyikan item kecil", "showall": "Papar semua", "analyze": "Analisis", "quit": "Tutup pelayan", "legend": "Petunjuk", "help": "Bantuan", "stop": "Henti", "choose": "Pilih folder ini", "cancel": "Batal", "sortby": "Isih", "byname": "nama", "bysize": "saiz", "scangrey": "Imbas folder kelabu", "leggreen": "Folder diimbas, saiz diketahui", "legorange": "Folder dirancang dalam imbasan semasa", "legblue": "Folder dalam baris gilir (diimbas selepas imbasan semasa)", "legspin": "Folder sedang diimbas", "leggrey": "Belum diimbas, dikecualikan, junction atau dilindungi. Klik untuk mengimbas (2 aras)", "legfile": "Fail", "legbar": "Bahagian daripada saiz folder semasa", "helpnav": "Navigasi: klik folder untuk masuk, jejak di atas untuk naik.", "helpgrey": "Folder kelabu: klik untuk mengimbas (2 aras lagi). Semasa imbasan ia masuk baris gilir (biru) dan diimbas di akhir.", "helpfiles": "Fail: dipaparkan apabila anda buka folder, diisih ikut saiz, terhad kepada 1000 terbesar.", "helpfilter": "Penapis paparan (butang Imbasan baharu): menyembunyikan item kecil untuk kebolehbacaan, tanpa mengubah imbasan.", "helpone": "Satu imbasan pada satu masa: dua imbasan cakera selari akan memperlahankan satu sama lain.", "noanalysis": "Tiada analisis", "theme": "Tema cerah / gelap", "ready": "Sedia.", "launching": "Memulakan imbasan...", "done": "Imbasan selesai. Navigasi bebas.", "inprogress": "Sedang berjalan:", "scanningtag": "mengimbas", "parent": ".. (folder induk)", "scandots": "Mengimbas...", "empty": "Folder kosong", "loadingfiles": "Memuatkan fail...", "excluded": "dikecualikan", "junction": "junction", "unscanned": "belum diimbas", "protected": "dilindungi", "error": "ralat", "scanshort": "imbas...", "queued": "dalam baris gilir", "levels": "aras", "hide1": "Sembunyi bawah 1 MB", "hide100": "Sembunyi bawah 100 MB", "hide1g": "Sembunyi bawah 1 GB", "genby": "Dijana oleh:", "invalidpath": "Laluan tidak sah atau tidak boleh diakses.", "windowopen": "Tetingkap dibuka...", "interrupted": "Imbasan terhenti.", "enterpath": "Sila masukkan laluan.", "serverstopped": "Pelayan dihentikan. Anda boleh tutup tab ini.", "you": "(anda)", "network": "(rangkaian)", "drivefree": "bebas", "removeone": "Buang", "clearhist": "Kosongkan sejarah", "exclusions": "Pengecualian", "scandepthfull": "Kedalaman: tanpa had", "scandepthn": "Kedalaman: {n}", "browsetitle": "Semak imbas", "upfolder": "Folder induk", "clickscan": "Klik untuk mengimbas folder ini (2 aras)", "depthhigh": "Terperinci: banyak aras dipramuat, lebih berat untuk dipaparkan.", "depthlow": "Ringan dan pantas: hanya aras pertama dimuatkan. Turun dengan klik folder.", "depthmid": "Keseimbangan baik: beberapa aras kelihatan serentak, masih lancar.", "depthtail": "Saiz sentiasa tepat; melepasi kedalaman ini, klik folder kelabu untuk meneroka.", "depthunl": "Melalui dan memaparkan seluruh pokok. Boleh sangat lama dan berat pada cakera besar.", "exclcustom": "Kecualikan juga (satu laluan setiap baris):", "exclexp": "Folder sistem ini sentiasa dilangkau semasa imbasan.", "filesafter": "Fail akan dipaparkan di akhir imbasan semasa.", "filescap": "Paparan terhad kepada 1000 fail terbesar dalam folder ini.", "filterexp": "Menapis paparan sahaja, untuk kebolehbacaan. Tidak mempercepat imbasan: saiz sentiasa dikira sepenuhnya. Boleh diubah bila-bila masa.", "hiddenmsg": "{n} item di bawah ambang disembunyikan oleh penapis paparan.", "modalhint": "Pilih folder untuk dianalisis, kemudian tetapkan kedalaman dan paparan.", "pickerfail": "Pemilih folder asli tidak dapat dibuka pada mesin ini. Taip laluan secara terus.", "queuedmsg": "Folder dalam baris gilir ({n} menunggu). Diimbas di akhir imbasan semasa."},fil:{"sub": "Analyzer ng Paggamit ng Disk", "newscan": "Bagong scan", "total": "Kabuuan:", "newanalysis": "Bagong pagsusuri", "folderlabel": "Folder na susuriin", "browse": "Mag-browse...", "quickaccess": "Mabilis na access", "recent": "Kamakailan", "drives": "Mga drive", "depthlabel": "Lalim ng paggalugad:", "unlimited": "Walang limitasyon (ini-scan lahat, maaaring napakatagal)", "displaylabel": "Display: itago ang maliliit na item", "showall": "Ipakita lahat", "analyze": "Suriin", "quit": "Isara ang server", "legend": "Alamat", "help": "Tulong", "stop": "Ihinto", "choose": "Piliin ang folder na ito", "cancel": "Kanselahin", "sortby": "Ayusin", "byname": "pangalan", "bysize": "laki", "scangrey": "I-scan ang kulay-abong folder", "leggreen": "Na-scan na folder, alam ang laki", "legorange": "Folder na nakaplano sa kasalukuyang scan", "legblue": "Folder sa pila (isi-scan pagkatapos ng kasalukuyang scan)", "legspin": "Folder na sina-scan", "leggrey": "Hindi na-scan, hindi kasama, junction o protektado. I-click para i-scan (2 antas)", "legfile": "File", "legbar": "Bahagi sa laki ng kasalukuyang folder", "helpnav": "Mag-navigate: i-click ang folder para pumasok, breadcrumb sa itaas para umakyat.", "helpgrey": "Kulay-abong folder: i-click para i-scan (2 pang antas). Habang nag-i-scan, pumipila sila (asul) at sina-scan sa huli.", "helpfiles": "Mga file: ipinapakita kapag binuksan mo ang folder, inaayos ayon sa laki, limitado sa 1000 pinakamalaki.", "helpfilter": "Filter ng display (button na Bagong scan): itinatago ang maliliit na item para sa kalinawan, nang hindi binabago ang scan.", "helpone": "Isang scan sa isang pagkakataon: dalawang magkatabing disk scan ay magpapabagal sa isa't isa.", "noanalysis": "Walang pagsusuri", "theme": "Maliwanag / madilim na tema", "ready": "Handa na.", "launching": "Sinisimulan ang scan...", "done": "Tapos na ang scan. Malayang mag-navigate.", "inprogress": "Kasalukuyang ginagawa:", "scanningtag": "nag-i-scan", "parent": ".. (folder na magulang)", "scandots": "Nag-i-scan...", "empty": "Walang laman na folder", "loadingfiles": "Nilo-load ang mga file...", "excluded": "hindi kasama", "junction": "junction", "unscanned": "hindi na-scan", "protected": "protektado", "error": "error", "scanshort": "scan...", "queued": "nasa pila", "levels": "antas", "hide1": "Itago ang wala pang 1 MB", "hide100": "Itago ang wala pang 100 MB", "hide1g": "Itago ang wala pang 1 GB", "genby": "Ginawa ng:", "invalidpath": "Di-wasto o di-maabot na path.", "windowopen": "Nabuksan ang window...", "interrupted": "Naantala ang scan.", "enterpath": "Maglagay ng path.", "serverstopped": "Huminto ang server. Puwede mong isara ang tab na ito.", "you": "(ikaw)", "network": "(network)", "drivefree": "libre", "removeone": "Alisin", "clearhist": "I-clear ang history", "exclusions": "Mga eksklusyon", "scandepthfull": "Lalim: walang limitasyon", "scandepthn": "Lalim: {n}", "browsetitle": "Mag-browse", "upfolder": "Folder na magulang", "clickscan": "I-click para i-scan ang folder na ito (2 antas)", "depthhigh": "Detalyado: maraming antas ang naka-preload, mas mabigat ipakita.", "depthlow": "Magaan at mabilis: unang antas lang ang nilo-load. Bumaba sa pag-click ng folder.", "depthmid": "Magandang balanse: ilang antas ang nakikita nang sabay, maayos pa rin.", "depthtail": "Palaging eksakto ang mga laki; lampas sa lalim na ito, i-click ang kulay-abong folder para tuklasin.", "depthunl": "Dinadaanan at ipinapakita ang buong tree. Maaaring napakatagal at mabigat sa malaking disk.", "exclcustom": "Ibukod din (isang path bawat linya):", "exclexp": "Palaging nilalaktawan ang mga system folder na ito habang nag-i-scan.", "filesafter": "Lalabas ang mga file sa dulo ng kasalukuyang scan.", "filescap": "Limitado ang display sa 1000 pinakamalaking file sa folder na ito.", "filterexp": "Sinasala lang ang display, para sa kalinawan. Hindi nito pinapabilis ang scan: palaging buong kinukwenta ang laki. Mababago anumang oras.", "hiddenmsg": "{n} item sa ilalim ng threshold ang itinago ng display filter.", "modalhint": "Piliin ang folder na susuriin, pagkatapos itakda ang lalim at display.", "pickerfail": "Hindi mabuksan ang native folder picker sa makinang ito. I-type nang direkta ang path.", "queuedmsg": "Nakapila ang folder ({n} naghihintay). Isi-scan sa dulo ng kasalukuyang scan."},sw:{"sub": "Kichanganuzi cha Matumizi ya Diski", "newscan": "Uchanganuzi mpya", "total": "Jumla:", "newanalysis": "Uchambuzi mpya", "folderlabel": "Folda ya kuchambua", "browse": "Vinjari...", "quickaccess": "Ufikiaji wa haraka", "recent": "Za hivi karibuni", "drives": "Diski", "depthlabel": "Kina cha uchunguzi:", "unlimited": "Bila kikomo (huchanganua yote, kunaweza kuchukua muda mrefu sana)", "displaylabel": "Onyesho: ficha vipengele vidogo", "showall": "Onyesha vyote", "analyze": "Changanua", "quit": "Funga seva", "legend": "Ufafanuzi", "help": "Msaada", "stop": "Simamisha", "choose": "Chagua folda hii", "cancel": "Ghairi", "sortby": "Panga", "byname": "jina", "bysize": "ukubwa", "scangrey": "Changanua folda za kijivu", "leggreen": "Folda iliyochanganuliwa, ukubwa unajulikana", "legorange": "Folda iliyopangwa katika uchanganuzi wa sasa", "legblue": "Folda kwenye foleni (huchanganuliwa baada ya uchanganuzi wa sasa)", "legspin": "Folda inachanganuliwa", "leggrey": "Haijachanganuliwa, imetengwa, kiungo au imelindwa. Bofya kuichanganua (viwango 2)", "legfile": "Faili", "legbar": "Sehemu ya ukubwa wa folda ya sasa", "helpnav": "Sogeza: bofya folda kuingia, njia ya juu kupanda.", "helpgrey": "Folda za kijivu: bofya kuzichanganua (viwango 2 zaidi). Wakati wa uchanganuzi zinaingia foleni (buluu) na huchanganuliwa mwishoni.", "helpfiles": "Faili: huonyeshwa unapofungua folda, hupangwa kwa ukubwa, hadi 1000 kubwa zaidi.", "helpfilter": "Kichujio cha maonyesho (kitufe cha Uchanganuzi mpya): huficha vipengele vidogo kwa usomaji, bila kubadilisha uchanganuzi.", "helpone": "Uchanganuzi mmoja kwa wakati: uchanganuzi wa diski mbili sambamba ungepunguzana kasi.", "noanalysis": "Hakuna uchambuzi", "theme": "Mandhari angavu / meusi", "ready": "Tayari.", "launching": "Inaanza uchanganuzi...", "done": "Uchanganuzi umekamilika. Nenda kwa uhuru.", "inprogress": "Inaendelea:", "scanningtag": "inachanganua", "parent": ".. (folda mzazi)", "scandots": "Inachanganua...", "empty": "Folda tupu", "loadingfiles": "Inapakia faili...", "excluded": "imetengwa", "junction": "kiungo", "unscanned": "haijachanganuliwa", "protected": "imelindwa", "error": "hitilafu", "scanshort": "changanua...", "queued": "kwenye foleni", "levels": "ngazi", "hide1": "Ficha chini ya 1 MB", "hide100": "Ficha chini ya 100 MB", "hide1g": "Ficha chini ya 1 GB", "genby": "Imetengenezwa na:", "invalidpath": "Njia batili au isiyofikika.", "windowopen": "Dirisha limefunguliwa...", "interrupted": "Uchanganuzi umesimamishwa.", "enterpath": "Tafadhali weka njia.", "serverstopped": "Seva imesimamishwa. Unaweza kufunga kichupo hiki.", "you": "(wewe)", "network": "(mtandao)", "drivefree": "bure", "removeone": "Ondoa", "clearhist": "Futa historia", "exclusions": "Vitengo", "scandepthfull": "Kina: bila kikomo", "scandepthn": "Kina: {n}", "browsetitle": "Vinjari", "upfolder": "Folda mzazi", "clickscan": "Bofya kuchanganua folda hii (viwango 2)", "depthhigh": "Kwa kina: viwango vingi vimepakiwa mapema, vizito kuonyesha.", "depthlow": "Nyepesi na haraka: viwango vya kwanza tu hupakiwa. Shuka kwa kubofya folda.", "depthmid": "Usawa mzuri: viwango kadhaa vinaonekana kwa pamoja, bado laini.", "depthtail": "Ukubwa daima ni sahihi; zaidi ya kina hiki, bofya folda ya kijivu kuichunguza.", "depthunl": "Hupitia na kuonyesha mti mzima. Inaweza kuwa ndefu na nzito sana kwenye diski kubwa.", "exclcustom": "Tenga pia (njia moja kwa mstari):", "exclexp": "Folda hizi za mfumo huruka daima wakati wa uchanganuzi.", "filesafter": "Faili zitaonyeshwa mwishoni mwa uchanganuzi wa sasa.", "filescap": "Onyesho limepunguzwa kwa faili 1000 kubwa zaidi kwenye folda hii.", "filterexp": "Huchuja onyesho tu, kwa usomaji. Haiharakishi uchanganuzi: ukubwa daima huhesabiwa kikamilifu. Inaweza kubadilishwa wakati wowote.", "hiddenmsg": "Vipengele {n} chini ya kizingiti vimefichwa na kichujio cha onyesho.", "modalhint": "Chagua folda ya kuchambua, kisha weka kina na onyesho.", "pickerfail": "Kichaguzi cha folda cha asili hakikuweza kufunguka kwenye mashine hii. Andika njia moja kwa moja.", "queuedmsg": "Folda imewekwa kwenye foleni ({n} zinasubiri). Itachanganuliwa mwishoni mwa uchanganuzi wa sasa."},ta:{"sub": "வட்டு பயன்பாட்டு பகுப்பாய்வி", "newscan": "புதிய ஸ்கேன்", "total": "மொத்தம்:", "newanalysis": "புதிய பகுப்பாய்வு", "folderlabel": "பகுப்பாய்வு செய்யும் கோப்புறை", "browse": "உலாவு...", "quickaccess": "விரைவு அணுகல்", "recent": "சமீபத்தியவை", "drives": "இயக்கிகள்", "depthlabel": "ஆய்வு ஆழம்:", "unlimited": "வரம்பற்றது (அனைத்தையும் ஸ்கேன் செய்யும், மிக நீளமாக இருக்கலாம்)", "displaylabel": "காட்சி: சிறிய உருப்படிகளை மறை", "showall": "அனைத்தையும் காட்டு", "analyze": "பகுப்பாய்வு", "quit": "சேவையகத்தை மூடு", "legend": "விளக்கம்", "help": "உதவி", "stop": "நிறுத்து", "choose": "இந்தக் கோப்புறையைத் தேர்வுசெய்", "cancel": "ரத்து", "sortby": "வரிசைப்படுத்து", "byname": "பெயர்", "bysize": "அளவு", "scangrey": "சாம்பல் கோப்புறைகளை ஸ்கேன் செய்", "leggreen": "ஸ்கேன் செய்யப்பட்ட கோப்புறை, அளவு தெரியும்", "legorange": "தற்போதைய ஸ்கேனில் திட்டமிடப்பட்ட கோப்புறை", "legblue": "வரிசையில் உள்ள கோப்புறை (தற்போதைய ஸ்கேனுக்குப் பிறகு)", "legspin": "ஸ்கேன் செய்யப்படும் கோப்புறை", "leggrey": "ஸ்கேன் செய்யப்படவில்லை, விலக்கப்பட்டது, இணைப்பு அல்லது பாதுகாக்கப்பட்டது. ஸ்கேன் செய்ய கிளிக் செய்யவும் (2 நிலைகள்)", "legfile": "கோப்பு", "legbar": "தற்போதைய கோப்புறை அளவில் பங்கு", "helpnav": "வழிசெலுத்தல்: நுழைய கோப்புறையைக் கிளிக் செய்யவும், மேலே உள்ள பாதையால் மேலே செல்லவும்.", "helpgrey": "சாம்பல் கோப்புறைகள்: ஸ்கேன் செய்ய கிளிக் செய்யவும் (மேலும் 2 நிலைகள்). ஸ்கேன் போது வரிசையில் (நீலம்) சேர்ந்து இறுதியில் ஸ்கேன் செய்யப்படும்.", "helpfiles": "கோப்புகள்: கோப்புறையைத் திறக்கும்போது காட்டப்படும், அளவின்படி வரிசைப்படுத்தப்படும், பெரிய 1000 வரை.", "helpfilter": "காட்சி வடிகட்டி (புதிய ஸ்கேன் பொத்தான்): ஸ்கேனை மாற்றாமல் சிறிய உருப்படிகளை மறைக்கும்.", "helpone": "ஒரு நேரத்தில் ஒரு ஸ்கேன்: இணையான இரண்டு வட்டு ஸ்கேன்கள் ஒன்றையொன்று மெதுவாக்கும்.", "noanalysis": "பகுப்பாய்வு இல்லை", "theme": "வெளிர் / இருள் தீம்", "ready": "தயார்.", "launching": "ஸ்கேன் தொடங்குகிறது...", "done": "ஸ்கேன் முடிந்தது. சுதந்திரமாக உலாவவும்.", "inprogress": "நடைபெறுகிறது:", "scanningtag": "ஸ்கேன் செய்கிறது", "parent": ".. (பெற்றோர் கோப்புறை)", "scandots": "ஸ்கேன் செய்கிறது...", "empty": "காலி கோப்புறை", "loadingfiles": "கோப்புகள் ஏற்றுகிறது...", "excluded": "விலக்கப்பட்டது", "junction": "இணைப்பு", "unscanned": "ஸ்கேன் செய்யப்படவில்லை", "protected": "பாதுகாக்கப்பட்டது", "error": "பிழை", "scanshort": "ஸ்கேன்...", "queued": "வரிசையில்", "levels": "நிலைகள்", "hide1": "1 MB கீழ் மறை", "hide100": "100 MB கீழ் மறை", "hide1g": "1 GB கீழ் மறை", "genby": "உருவாக்கியவர்:", "invalidpath": "தவறான அல்லது அணுக முடியாத பாதை.", "windowopen": "சாளரம் திறக்கப்பட்டது...", "interrupted": "ஸ்கேன் இடைநிறுத்தப்பட்டது.", "enterpath": "பாதையை உள்ளிடவும்.", "serverstopped": "சேவையகம் நிறுத்தப்பட்டது. இந்த தாவலை மூடலாம்.", "you": "(நீங்கள்)", "network": "(வலையமைப்பு)", "drivefree": "காலி", "removeone": "அகற்று", "clearhist": "வரலாற்றை அழி", "exclusions": "விலக்குகள்", "scandepthfull": "ஆழம்: வரம்பற்றது", "scandepthn": "ஆழம்: {n}", "browsetitle": "உலாவு", "upfolder": "பெற்றோர் கோப்புறை", "clickscan": "இந்தக் கோப்புறையை ஸ்கேன் செய்ய கிளிக் செய்யவும் (2 நிலைகள்)", "depthhigh": "விரிவானது: பல நிலைகள் முன்னேற்றப்பட்டன, காட்சி கனமானது.", "depthlow": "இலகு மற்றும் வேகம்: முதல் நிலைகள் மட்டும் ஏற்றப்படும். கோப்புறையைக் கிளிக் செய்து கீழே செல்லவும்.", "depthmid": "நல்ல சமநிலை: பல நிலைகள் ஒரே நேரத்தில் தெரியும், இன்னும் சீராக.", "depthtail": "அளவுகள் எப்போதும் துல்லியம்; இந்த ஆழத்திற்கு அப்பால், சாம்பல் கோப்புறையைக் கிளிக் செய்து ஆராயவும்.", "depthunl": "முழு மரத்தையும் கடந்து காட்டுகிறது. பெரிய வட்டில் மிக நீளமும் கனமும் ஆகலாம்.", "exclcustom": "மேலும் விலக்கு (ஒரு வரிக்கு ஒரு பாதை):", "exclexp": "இந்த சிஸ்டம் கோப்புறைகள் ஸ்கேனில் எப்போதும் தவிர்க்கப்படும்.", "filesafter": "தற்போதைய ஸ்கேன் முடிந்ததும் கோப்புகள் காட்டப்படும்.", "filescap": "இந்தக் கோப்புறையின் பெரிய 1000 கோப்புகளுக்கு காட்சி வரம்பு.", "filterexp": "காட்சியை மட்டும் வடிகட்டும், படிக்க எளிதாக. ஸ்கேனை வேகப்படுத்தாது: அளவுகள் எப்போதும் முழுமையாக கணக்கிடப்படும். எப்போதும் மாற்றலாம்.", "hiddenmsg": "வரம்புக்கு கீழ் {n} உருப்படி(கள்) காட்சி வடிகட்டியால் மறைக்கப்பட்டன.", "modalhint": "பகுப்பாய்வு செய்யும் கோப்புறையைத் தேர்வுசெய்து, ஆழத்தையும் காட்சியையும் அமைக்கவும்.", "pickerfail": "இந்த இயந்திரத்தில் சொந்த கோப்புறை தேர்வி திறக்க முடியவில்லை. பாதையை நேரடியாக தட்டச்சு செய்யவும்.", "queuedmsg": "கோப்புறை வரிசையில் ({n} காத்திருப்பு). தற்போதைய ஸ்கேன் முடிவில் ஸ்கேன் செய்யப்படும்."},te:{"sub": "డిస్క్ వినియోగ విశ్లేషకి", "newscan": "కొత్త స్కాన్", "total": "మొత్తం:", "newanalysis": "కొత్త విశ్లేషణ", "folderlabel": "విశ్లేషించాల్సిన ఫోల్డర్", "browse": "బ్రౌజ్...", "quickaccess": "త్వరిత ప్రాప్యత", "recent": "ఇటీవలివి", "drives": "డ్రైవ్‌లు", "depthlabel": "అన్వేషణ లోతు:", "unlimited": "అపరిమితం (అన్నింటినీ స్కాన్ చేస్తుంది, చాలా సమయం పట్టవచ్చు)", "displaylabel": "ప్రదర్శన: చిన్న అంశాలను దాచు", "showall": "అన్నీ చూపు", "analyze": "విశ్లేషించు", "quit": "సర్వర్‌ను మూసివేయి", "legend": "సూచిక", "help": "సహాయం", "stop": "ఆపు", "choose": "ఈ ఫోల్డర్‌ను ఎంచుకో", "cancel": "రద్దు", "sortby": "క్రమబద్ధీకరించు", "byname": "పేరు", "bysize": "పరిమాణం", "scangrey": "బూడిద ఫోల్డర్‌లను స్కాన్ చేయి", "leggreen": "స్కాన్ చేసిన ఫోల్డర్, పరిమాణం తెలుసు", "legorange": "ప్రస్తుత స్కాన్‌లో ప్రణాళిక చేసిన ఫోల్డర్", "legblue": "క్యూలో ఉన్న ఫోల్డర్ (ప్రస్తుత స్కాన్ తర్వాత)", "legspin": "స్కాన్ అవుతున్న ఫోల్డర్", "leggrey": "స్కాన్ కాలేదు, మినహాయించబడింది, జంక్షన్ లేదా రక్షితం. స్కాన్ చేయడానికి క్లిక్ చేయండి (2 స్థాయిలు)", "legfile": "ఫైల్", "legbar": "ప్రస్తుత ఫోల్డర్ పరిమాణంలో వాటా", "helpnav": "నావిగేషన్: ప్రవేశించడానికి ఫోల్డర్‌పై క్లిక్ చేయండి, పైకి వెళ్లడానికి పైన ఉన్న మార్గం.", "helpgrey": "బూడిద ఫోల్డర్‌లు: స్కాన్ చేయడానికి క్లిక్ చేయండి (మరో 2 స్థాయిలు). స్కాన్ సమయంలో అవి క్యూలో (నీలం) చేరి చివర్లో స్కాన్ అవుతాయి.", "helpfiles": "ఫైల్‌లు: ఫోల్డర్ తెరిచినప్పుడు చూపబడతాయి, పరిమాణం ప్రకారం క్రమబద్ధీకరించబడతాయి, పెద్ద 1000 వరకు.", "helpfilter": "ప్రదర్శన వడపోత (కొత్త స్కాన్ బటన్): స్కాన్‌ను మార్చకుండా చిన్న అంశాలను దాచుతుంది.", "helpone": "ఒకేసారి ఒక స్కాన్: రెండు సమాంతర డిస్క్ స్కాన్‌లు ఒకదానికొకటి నెమ్మదిస్తాయి.", "noanalysis": "విశ్లేషణ లేదు", "theme": "లేత / ముదురు థీమ్", "ready": "సిద్ధం.", "launching": "స్కాన్ ప్రారంభమవుతోంది...", "done": "స్కాన్ పూర్తయింది. స్వేచ్ఛగా విహరించండి.", "inprogress": "జరుగుతోంది:", "scanningtag": "స్కాన్ అవుతోంది", "parent": ".. (మూల ఫోల్డర్)", "scandots": "స్కాన్ అవుతోంది...", "empty": "ఖాళీ ఫోల్డర్", "loadingfiles": "ఫైల్‌లు లోడ్ అవుతున్నాయి...", "excluded": "మినహాయించబడింది", "junction": "జంక్షన్", "unscanned": "స్కాన్ కాలేదు", "protected": "రక్షితం", "error": "లోపం", "scanshort": "స్కాన్...", "queued": "క్యూలో", "levels": "స్థాయిలు", "hide1": "1 MB కంటే తక్కువ దాచు", "hide100": "100 MB కంటే తక్కువ దాచు", "hide1g": "1 GB కంటే తక్కువ దాచు", "genby": "రూపొందించినది:", "invalidpath": "చెల్లని లేదా అందుబాటులో లేని మార్గం.", "windowopen": "విండో తెరవబడింది...", "interrupted": "స్కాన్ ఆగింది.", "enterpath": "మార్గాన్ని నమోదు చేయండి.", "serverstopped": "సర్వర్ ఆగింది. మీరు ఈ ట్యాబ్‌ను మూసివేయవచ్చు.", "you": "(మీరు)", "network": "(నెట్‌వర్క్)", "drivefree": "ఖాళీ", "removeone": "తీసివేయి", "clearhist": "చరిత్రను తుడిచివేయి", "exclusions": "మినహాయింపులు", "scandepthfull": "లోతు: అపరిమితం", "scandepthn": "లోతు: {n}", "browsetitle": "బ్రౌజ్", "upfolder": "మూల ఫోల్డర్", "clickscan": "ఈ ఫోల్డర్‌ను స్కాన్ చేయడానికి క్లిక్ చేయండి (2 స్థాయిలు)", "depthhigh": "వివరణాత్మకం: చాలా స్థాయిలు ముందే లోడ్, ప్రదర్శన బరువు.", "depthlow": "తేలిక మరియు వేగం: మొదటి స్థాయిలు మాత్రమే లోడ్. ఫోల్డర్‌పై క్లిక్ చేసి లోతుకు వెళ్లండి.", "depthmid": "మంచి సమతుల్యత: అనేక స్థాయిలు ఒకేసారి కనిపిస్తాయి, ఇంకా సాఫీగా.", "depthtail": "పరిమాణాలు ఎప్పుడూ ఖచ్చితం; ఈ లోతు దాటి, బూడిద ఫోల్డర్‌పై క్లిక్ చేసి అన్వేషించండి.", "depthunl": "మొత్తం చెట్టును నడిచి చూపిస్తుంది. పెద్ద డిస్క్‌లో చాలా సమయం మరియు బరువు కావచ్చు.", "exclcustom": "అదనంగా మినహాయించు (ప్రతి పంక్తికి ఒక మార్గం):", "exclexp": "ఈ సిస్టమ్ ఫోల్డర్‌లు స్కాన్‌లో ఎప్పుడూ దాటవేయబడతాయి.", "filesafter": "ప్రస్తుత స్కాన్ ముగిసిన తర్వాత ఫైల్‌లు చూపబడతాయి.", "filescap": "ఈ ఫోల్డర్‌లోని పెద్ద 1000 ఫైల్‌లకు ప్రదర్శన పరిమితం.", "filterexp": "ప్రదర్శనను మాత్రమే వడపోస్తుంది, చదవడానికి. స్కాన్‌ను వేగవంతం చేయదు: పరిమాణాలు ఎప్పుడూ పూర్తిగా లెక్కించబడతాయి. ఎప్పుడైనా మార్చవచ్చు.", "hiddenmsg": "పరిమితి కంటే తక్కువ {n} అంశం(లు) ప్రదర్శన వడపోత ద్వారా దాచబడ్డాయి.", "modalhint": "విశ్లేషించాల్సిన ఫోల్డర్‌ను ఎంచుకుని, లోతు మరియు ప్రదర్శనను సెట్ చేయండి.", "pickerfail": "ఈ యంత్రంలో స్థానిక ఫోల్డర్ ఎంపిక తెరవలేకపోయింది. మార్గాన్ని నేరుగా టైప్ చేయండి.", "queuedmsg": "ఫోల్డర్ క్యూలో ({n} వేచి ఉన్నాయి). ప్రస్తుత స్కాన్ చివర్లో స్కాన్ అవుతుంది."},mr:{"sub": "डिस्क वापर विश्लेषक", "newscan": "नवीन स्कॅन", "total": "एकूण:", "newanalysis": "नवीन विश्लेषण", "folderlabel": "विश्लेषण करायचे फोल्डर", "browse": "ब्राउझ करा...", "quickaccess": "जलद प्रवेश", "recent": "अलीकडील", "drives": "ड्राइव्ह", "depthlabel": "अन्वेषण खोली:", "unlimited": "अमर्यादित (सर्व स्कॅन करते, खूप वेळ लागू शकतो)", "displaylabel": "प्रदर्शन: लहान घटक लपवा", "showall": "सर्व दाखवा", "analyze": "विश्लेषण करा", "quit": "सर्व्हर बंद करा", "legend": "सूची", "help": "मदत", "stop": "थांबवा", "choose": "हे फोल्डर निवडा", "cancel": "रद्द करा", "sortby": "क्रमवारी", "byname": "नाव", "bysize": "आकार", "scangrey": "राखाडी फोल्डर स्कॅन करा", "leggreen": "स्कॅन केलेले फोल्डर, आकार माहीत", "legorange": "सध्याच्या स्कॅनमध्ये नियोजित फोल्डर", "legblue": "रांगेतील फोल्डर (सध्याच्या स्कॅननंतर)", "legspin": "स्कॅन होत असलेले फोल्डर", "leggrey": "स्कॅन नाही, वगळलेले, जंक्शन किंवा संरक्षित. स्कॅन करण्यासाठी क्लिक करा (2 स्तर)", "legfile": "फाइल", "legbar": "सध्याच्या फोल्डर आकारातील वाटा", "helpnav": "नेव्हिगेशन: आत जाण्यासाठी फोल्डरवर क्लिक करा, वर जाण्यासाठी वरील मार्ग.", "helpgrey": "राखाडी फोल्डर: स्कॅन करण्यासाठी क्लिक करा (आणखी 2 स्तर). स्कॅन दरम्यान ती रांगेत (निळे) जातात आणि शेवटी स्कॅन होतात.", "helpfiles": "फाइल्स: फोल्डर उघडल्यावर दाखवल्या जातात, आकारानुसार क्रमवारी, सर्वात मोठ्या 1000 पर्यंत.", "helpfilter": "प्रदर्शन फिल्टर (नवीन स्कॅन बटण): स्कॅन न बदलता लहान घटक लपवते.", "helpone": "एका वेळी एक स्कॅन: दोन समांतर डिस्क स्कॅन एकमेकांना मंद करतील.", "noanalysis": "विश्लेषण नाही", "theme": "उजळ / गडद थीम", "ready": "तयार.", "launching": "स्कॅन सुरू होत आहे...", "done": "स्कॅन पूर्ण. मुक्तपणे नेव्हिगेट करा.", "inprogress": "सुरू आहे:", "scanningtag": "स्कॅन होत आहे", "parent": ".. (मूळ फोल्डर)", "scandots": "स्कॅन होत आहे...", "empty": "रिकामे फोल्डर", "loadingfiles": "फाइल्स लोड होत आहेत...", "excluded": "वगळले", "junction": "जंक्शन", "unscanned": "स्कॅन नाही", "protected": "संरक्षित", "error": "त्रुटी", "scanshort": "स्कॅन...", "queued": "रांगेत", "levels": "स्तर", "hide1": "1 MB खाली लपवा", "hide100": "100 MB खाली लपवा", "hide1g": "1 GB खाली लपवा", "genby": "तयार केले:", "invalidpath": "अवैध किंवा अगम्य मार्ग.", "windowopen": "विंडो उघडली...", "interrupted": "स्कॅन थांबवले.", "enterpath": "कृपया मार्ग प्रविष्ट करा.", "serverstopped": "सर्व्हर थांबवला. तुम्ही हा टॅब बंद करू शकता.", "you": "(तुम्ही)", "network": "(नेटवर्क)", "drivefree": "मोकळे", "removeone": "काढा", "clearhist": "इतिहास साफ करा", "exclusions": "वगळणे", "scandepthfull": "खोली: अमर्यादित", "scandepthn": "खोली: {n}", "browsetitle": "ब्राउझ", "upfolder": "मूळ फोल्डर", "clickscan": "हे फोल्डर स्कॅन करण्यासाठी क्लिक करा (2 स्तर)", "depthhigh": "तपशीलवार: अनेक स्तर आधीच लोड, दाखवणे जड.", "depthlow": "हलके आणि जलद: फक्त पहिले स्तर लोड होतात. फोल्डरवर क्लिक करून खाली जा.", "depthmid": "चांगला समतोल: अनेक स्तर एकाच वेळी दिसतात, तरीही सुरळीत.", "depthtail": "आकार नेहमी अचूक; या खोलीच्या पलीकडे, राखाडी फोल्डरवर क्लिक करून शोधा.", "depthunl": "संपूर्ण वृक्ष फिरते आणि दाखवते. मोठ्या डिस्कवर खूप लांब आणि जड असू शकते.", "exclcustom": "आणखी वगळा (प्रति ओळ एक मार्ग):", "exclexp": "हे सिस्टम फोल्डर स्कॅन दरम्यान नेहमी वगळले जातात.", "filesafter": "सध्याचा स्कॅन संपल्यावर फाइल्स दाखवल्या जातील.", "filescap": "या फोल्डरमधील सर्वात मोठ्या 1000 फाइल्सपर्यंत प्रदर्शन मर्यादित.", "filterexp": "फक्त प्रदर्शन फिल्टर करते, वाचनीयतेसाठी. स्कॅन वेगवान करत नाही: आकार नेहमी पूर्ण मोजले जातात. कधीही बदलता येते.", "hiddenmsg": "मर्यादेखालील {n} घटक प्रदर्शन फिल्टरने लपवले.", "modalhint": "विश्लेषण करायचे फोल्डर निवडा, नंतर खोली आणि प्रदर्शन सेट करा.", "pickerfail": "या मशीनवर मूळ फोल्डर निवडक उघडू शकला नाही. मार्ग थेट टाइप करा.", "queuedmsg": "फोल्डर रांगेत ({n} प्रतीक्षेत). सध्याच्या स्कॅनच्या शेवटी स्कॅन होईल."},gu:{"sub": "ડિસ્ક વપરાશ વિશ્લેષક", "newscan": "નવું સ્કૅન", "total": "કુલ:", "newanalysis": "નવું વિશ્લેષણ", "folderlabel": "વિશ્લેષણ કરવાનું ફોલ્ડર", "browse": "બ્રાઉઝ કરો...", "quickaccess": "ઝડપી ઍક્સેસ", "recent": "તાજેતરના", "drives": "ડ્રાઇવ્સ", "depthlabel": "અન્વેષણ ઊંડાઈ:", "unlimited": "અમર્યાદિત (બધું સ્કૅન કરે છે, ખૂબ લાંબું હોઈ શકે)", "displaylabel": "પ્રદર્શન: નાની આઇટમ છુપાવો", "showall": "બધું બતાવો", "analyze": "વિશ્લેષણ કરો", "quit": "સર્વર બંધ કરો", "legend": "સંકેત", "help": "મદદ", "stop": "અટકાવો", "choose": "આ ફોલ્ડર પસંદ કરો", "cancel": "રદ કરો", "sortby": "ક્રમમાં", "byname": "નામ", "bysize": "કદ", "scangrey": "ભૂખરા ફોલ્ડર સ્કૅન કરો", "leggreen": "સ્કૅન કરેલ ફોલ્ડર, કદ જાણીતું", "legorange": "વર્તમાન સ્કૅનમાં આયોજિત ફોલ્ડર", "legblue": "કતારમાં ફોલ્ડર (વર્તમાન સ્કૅન પછી)", "legspin": "સ્કૅન થઈ રહેલ ફોલ્ડર", "leggrey": "સ્કૅન નથી, બાકાત, જંક્શન અથવા સુરક્ષિત. સ્કૅન કરવા ક્લિક કરો (2 સ્તર)", "legfile": "ફાઇલ", "legbar": "વર્તમાન ફોલ્ડર કદમાં હિસ્સો", "helpnav": "નેવિગેશન: પ્રવેશવા ફોલ્ડર પર ક્લિક કરો, ઉપર જવા ટોચનો માર્ગ.", "helpgrey": "ભૂખરા ફોલ્ડર: સ્કૅન કરવા ક્લિક કરો (વધુ 2 સ્તર). સ્કૅન દરમિયાન તે કતારમાં (વાદળી) જાય છે અને અંતે સ્કૅન થાય છે.", "helpfiles": "ફાઇલો: ફોલ્ડર ખોલો ત્યારે બતાવાય, કદ મુજબ ગોઠવાય, સૌથી મોટી 1000 સુધી.", "helpfilter": "પ્રદર્શન ફિલ્ટર (નવું સ્કૅન બટન): સ્કૅન બદલ્યા વિના નાની આઇટમ છુપાવે.", "helpone": "એક સમયે એક સ્કૅન: બે સમાંતર ડિસ્ક સ્કૅન એકબીજાને ધીમા કરશે.", "noanalysis": "કોઈ વિશ્લેષણ નથી", "theme": "ઉજળી / ઘેરી થીમ", "ready": "તૈયાર.", "launching": "સ્કૅન શરૂ થઈ રહ્યું છે...", "done": "સ્કૅન પૂર્ણ. મુક્તપણે નેવિગેટ કરો.", "inprogress": "ચાલુ છે:", "scanningtag": "સ્કૅન થઈ રહ્યું છે", "parent": ".. (મૂળ ફોલ્ડર)", "scandots": "સ્કૅન થઈ રહ્યું છે...", "empty": "ખાલી ફોલ્ડર", "loadingfiles": "ફાઇલો લોડ થઈ રહી છે...", "excluded": "બાકાત", "junction": "જંક્શન", "unscanned": "સ્કૅન નથી", "protected": "સુરક્ષિત", "error": "ભૂલ", "scanshort": "સ્કૅન...", "queued": "કતારમાં", "levels": "સ્તર", "hide1": "1 MB નીચે છુપાવો", "hide100": "100 MB નીચે છુપાવો", "hide1g": "1 GB નીચે છુપાવો", "genby": "બનાવનાર:", "invalidpath": "અમાન્ય અથવા અગમ્ય પાથ.", "windowopen": "વિન્ડો ખૂલી...", "interrupted": "સ્કૅન અટકાવ્યું.", "enterpath": "કૃપા કરી પાથ દાખલ કરો.", "serverstopped": "સર્વર બંધ થયું. તમે આ ટૅબ બંધ કરી શકો છો.", "you": "(તમે)", "network": "(નેટવર્ક)", "drivefree": "મુક્ત", "removeone": "દૂર કરો", "clearhist": "ઇતિહાસ સાફ કરો", "exclusions": "બાકાત", "scandepthfull": "ઊંડાઈ: અમર્યાદિત", "scandepthn": "ઊંડાઈ: {n}", "browsetitle": "બ્રાઉઝ", "upfolder": "મૂળ ફોલ્ડર", "clickscan": "આ ફોલ્ડર સ્કૅન કરવા ક્લિક કરો (2 સ્તર)", "depthhigh": "વિગતવાર: ઘણા સ્તર પહેલેથી લોડ, બતાવવું ભારે.", "depthlow": "હલકું અને ઝડપી: માત્ર પહેલા સ્તર લોડ થાય. ફોલ્ડર પર ક્લિક કરી નીચે જાઓ.", "depthmid": "સારું સંતુલન: ઘણા સ્તર એકસાથે દેખાય, હજી સરળ.", "depthtail": "કદ હંમેશાં ચોક્કસ; આ ઊંડાઈથી આગળ, ભૂખરા ફોલ્ડર પર ક્લિક કરી શોધો.", "depthunl": "આખું વૃક્ષ ફરે અને બતાવે. મોટી ડિસ્ક પર ખૂબ લાંબું અને ભારે હોઈ શકે.", "exclcustom": "વધુ બાકાત (દરેક લીટીએ એક પાથ):", "exclexp": "આ સિસ્ટમ ફોલ્ડર સ્કૅન દરમિયાન હંમેશાં છોડી દેવાય છે.", "filesafter": "વર્તમાન સ્કૅન પૂરું થાય ત્યારે ફાઇલો બતાવાશે.", "filescap": "આ ફોલ્ડરની સૌથી મોટી 1000 ફાઇલો સુધી પ્રદર્શન મર્યાદિત.", "filterexp": "માત્ર પ્રદર્શન ફિલ્ટર કરે, વાંચવા માટે. સ્કૅન ઝડપી કરતું નથી: કદ હંમેશાં પૂરું ગણાય. ગમે ત્યારે બદલી શકાય.", "hiddenmsg": "મર્યાદા નીચે {n} આઇટમ પ્રદર્શન ફિલ્ટરે છુપાવી.", "modalhint": "વિશ્લેષણ કરવાનું ફોલ્ડર પસંદ કરો, પછી ઊંડાઈ અને પ્રદર્શન સેટ કરો.", "pickerfail": "આ મશીન પર મૂળ ફોલ્ડર પસંદકર્તા ખૂલી શક્યો નહીં. પાથ સીધો ટાઇપ કરો.", "queuedmsg": "ફોલ્ડર કતારમાં ({n} રાહમાં). વર્તમાન સ્કૅનના અંતે સ્કૅન થશે."},kn:{"sub": "ಡಿಸ್ಕ್ ಬಳಕೆ ವಿಶ್ಲೇಷಕ", "newscan": "ಹೊಸ ಸ್ಕ್ಯಾನ್", "total": "ಒಟ್ಟು:", "newanalysis": "ಹೊಸ ವಿಶ್ಲೇಷಣೆ", "folderlabel": "ವಿಶ್ಲೇಷಿಸಬೇಕಾದ ಫೋಲ್ಡರ್", "browse": "ಬ್ರೌಸ್...", "quickaccess": "ತ್ವರಿತ ಪ್ರವೇಶ", "recent": "ಇತ್ತೀಚಿನವು", "drives": "ಡ್ರೈವ್‌ಗಳು", "depthlabel": "ಅನ್ವೇಷಣೆ ಆಳ:", "unlimited": "ಅಪರಿಮಿತ (ಎಲ್ಲವನ್ನೂ ಸ್ಕ್ಯಾನ್ ಮಾಡುತ್ತದೆ, ಬಹಳ ಸಮಯ ತೆಗೆದುಕೊಳ್ಳಬಹುದು)", "displaylabel": "ಪ್ರದರ್ಶನ: ಸಣ್ಣ ಅಂಶಗಳನ್ನು ಮರೆಮಾಡಿ", "showall": "ಎಲ್ಲವನ್ನೂ ತೋರಿಸಿ", "analyze": "ವಿಶ್ಲೇಷಿಸಿ", "quit": "ಸರ್ವರ್ ಮುಚ್ಚಿ", "legend": "ಸೂಚಿ", "help": "ಸಹಾಯ", "stop": "ನಿಲ್ಲಿಸಿ", "choose": "ಈ ಫೋಲ್ಡರ್ ಆಯ್ಕೆಮಾಡಿ", "cancel": "ರದ್ದುಮಾಡಿ", "sortby": "ವಿಂಗಡಿಸಿ", "byname": "ಹೆಸರು", "bysize": "ಗಾತ್ರ", "scangrey": "ಬೂದು ಫೋಲ್ಡರ್‌ಗಳನ್ನು ಸ್ಕ್ಯಾನ್ ಮಾಡಿ", "leggreen": "ಸ್ಕ್ಯಾನ್ ಮಾಡಿದ ಫೋಲ್ಡರ್, ಗಾತ್ರ ತಿಳಿದಿದೆ", "legorange": "ಪ್ರಸ್ತುತ ಸ್ಕ್ಯಾನ್‌ನಲ್ಲಿ ಯೋಜಿಸಿದ ಫೋಲ್ಡರ್", "legblue": "ಸರತಿಯಲ್ಲಿರುವ ಫೋಲ್ಡರ್ (ಪ್ರಸ್ತುತ ಸ್ಕ್ಯಾನ್ ನಂತರ)", "legspin": "ಸ್ಕ್ಯಾನ್ ಆಗುತ್ತಿರುವ ಫೋಲ್ಡರ್", "leggrey": "ಸ್ಕ್ಯಾನ್ ಆಗಿಲ್ಲ, ಹೊರಗಿಡಲಾಗಿದೆ, ಜಂಕ್ಷನ್ ಅಥವಾ ರಕ್ಷಿತ. ಸ್ಕ್ಯಾನ್ ಮಾಡಲು ಕ್ಲಿಕ್ ಮಾಡಿ (2 ಹಂತ)", "legfile": "ಫೈಲ್", "legbar": "ಪ್ರಸ್ತುತ ಫೋಲ್ಡರ್ ಗಾತ್ರದಲ್ಲಿ ಪಾಲು", "helpnav": "ನ್ಯಾವಿಗೇಶನ್: ಪ್ರವೇಶಿಸಲು ಫೋಲ್ಡರ್ ಕ್ಲಿಕ್ ಮಾಡಿ, ಮೇಲಕ್ಕೆ ಹೋಗಲು ಮೇಲಿನ ಮಾರ್ಗ.", "helpgrey": "ಬೂದು ಫೋಲ್ಡರ್‌ಗಳು: ಸ್ಕ್ಯಾನ್ ಮಾಡಲು ಕ್ಲಿಕ್ ಮಾಡಿ (ಇನ್ನೂ 2 ಹಂತ). ಸ್ಕ್ಯಾನ್ ಸಮಯದಲ್ಲಿ ಅವು ಸರತಿಗೆ (ನೀಲಿ) ಸೇರಿ ಕೊನೆಯಲ್ಲಿ ಸ್ಕ್ಯಾನ್ ಆಗುತ್ತವೆ.", "helpfiles": "ಫೈಲ್‌ಗಳು: ಫೋಲ್ಡರ್ ತೆರೆದಾಗ ತೋರಿಸಲಾಗುತ್ತದೆ, ಗಾತ್ರದ ಪ್ರಕಾರ ವಿಂಗಡಿಸಲಾಗುತ್ತದೆ, ದೊಡ್ಡ 1000 ವರೆಗೆ.", "helpfilter": "ಪ್ರದರ್ಶನ ಫಿಲ್ಟರ್ (ಹೊಸ ಸ್ಕ್ಯಾನ್ ಬಟನ್): ಸ್ಕ್ಯಾನ್ ಬದಲಾಯಿಸದೆ ಸಣ್ಣ ಅಂಶಗಳನ್ನು ಮರೆಮಾಡುತ್ತದೆ.", "helpone": "ಒಂದೇ ಸಮಯದಲ್ಲಿ ಒಂದು ಸ್ಕ್ಯಾನ್: ಎರಡು ಸಮಾನಾಂತರ ಡಿಸ್ಕ್ ಸ್ಕ್ಯಾನ್‌ಗಳು ಪರಸ್ಪರ ನಿಧಾನಗೊಳಿಸುತ್ತವೆ.", "noanalysis": "ವಿಶ್ಲೇಷಣೆ ಇಲ್ಲ", "theme": "ತಿಳಿ / ಗಾಢ ಥೀಮ್", "ready": "ಸಿದ್ಧ.", "launching": "ಸ್ಕ್ಯಾನ್ ಪ್ರಾರಂಭವಾಗುತ್ತಿದೆ...", "done": "ಸ್ಕ್ಯಾನ್ ಮುಗಿದಿದೆ. ಮುಕ್ತವಾಗಿ ನ್ಯಾವಿಗೇಟ್ ಮಾಡಿ.", "inprogress": "ಪ್ರಗತಿಯಲ್ಲಿದೆ:", "scanningtag": "ಸ್ಕ್ಯಾನ್ ಆಗುತ್ತಿದೆ", "parent": ".. (ಮೂಲ ಫೋಲ್ಡರ್)", "scandots": "ಸ್ಕ್ಯಾನ್ ಆಗುತ್ತಿದೆ...", "empty": "ಖಾಲಿ ಫೋಲ್ಡರ್", "loadingfiles": "ಫೈಲ್‌ಗಳು ಲೋಡ್ ಆಗುತ್ತಿವೆ...", "excluded": "ಹೊರಗಿಡಲಾಗಿದೆ", "junction": "ಜಂಕ್ಷನ್", "unscanned": "ಸ್ಕ್ಯಾನ್ ಆಗಿಲ್ಲ", "protected": "ರಕ್ಷಿತ", "error": "ದೋಷ", "scanshort": "ಸ್ಕ್ಯಾನ್...", "queued": "ಸರತಿಯಲ್ಲಿ", "levels": "ಹಂತಗಳು", "hide1": "1 MB ಕೆಳಗೆ ಮರೆಮಾಡಿ", "hide100": "100 MB ಕೆಳಗೆ ಮರೆಮಾಡಿ", "hide1g": "1 GB ಕೆಳಗೆ ಮರೆಮಾಡಿ", "genby": "ರಚಿಸಿದವರು:", "invalidpath": "ಅಮಾನ್ಯ ಅಥವಾ ಪ್ರವೇಶಿಸಲಾಗದ ಮಾರ್ಗ.", "windowopen": "ವಿಂಡೋ ತೆರೆಯಲಾಗಿದೆ...", "interrupted": "ಸ್ಕ್ಯಾನ್ ಅಡ್ಡಿಪಡಿಸಲಾಗಿದೆ.", "enterpath": "ದಯವಿಟ್ಟು ಮಾರ್ಗ ನಮೂದಿಸಿ.", "serverstopped": "ಸರ್ವರ್ ನಿಲ್ಲಿಸಲಾಗಿದೆ. ಈ ಟ್ಯಾಬ್ ಮುಚ್ಚಬಹುದು.", "you": "(ನೀವು)", "network": "(ನೆಟ್‌ವರ್ಕ್)", "drivefree": "ಖಾಲಿ", "removeone": "ತೆಗೆದುಹಾಕಿ", "clearhist": "ಇತಿಹಾಸ ಅಳಿಸಿ", "exclusions": "ಹೊರಗಿಡುವಿಕೆಗಳು", "scandepthfull": "ಆಳ: ಅಪರಿಮಿತ", "scandepthn": "ಆಳ: {n}", "browsetitle": "ಬ್ರೌಸ್", "upfolder": "ಮೂಲ ಫೋಲ್ಡರ್", "clickscan": "ಈ ಫೋಲ್ಡರ್ ಸ್ಕ್ಯಾನ್ ಮಾಡಲು ಕ್ಲಿಕ್ ಮಾಡಿ (2 ಹಂತ)", "depthhigh": "ವಿವರವಾದ: ಹಲವು ಹಂತಗಳು ಮೊದಲೇ ಲೋಡ್, ತೋರಿಸಲು ಭಾರ.", "depthlow": "ಹಗುರ ಮತ್ತು ವೇಗ: ಮೊದಲ ಹಂತಗಳು ಮಾತ್ರ ಲೋಡ್. ಫೋಲ್ಡರ್ ಕ್ಲಿಕ್ ಮಾಡಿ ಕೆಳಗೆ ಹೋಗಿ.", "depthmid": "ಉತ್ತಮ ಸಮತೋಲನ: ಹಲವು ಹಂತಗಳು ಒಟ್ಟಿಗೆ ಕಾಣುತ್ತವೆ, ಇನ್ನೂ ಸುಗಮ.", "depthtail": "ಗಾತ್ರಗಳು ಯಾವಾಗಲೂ ನಿಖರ; ಈ ಆಳದ ಆಚೆ, ಬೂದು ಫೋಲ್ಡರ್ ಕ್ಲಿಕ್ ಮಾಡಿ ಅನ್ವೇಷಿಸಿ.", "depthunl": "ಇಡೀ ಮರವನ್ನು ನಡೆದು ತೋರಿಸುತ್ತದೆ. ದೊಡ್ಡ ಡಿಸ್ಕ್‌ನಲ್ಲಿ ಬಹಳ ಸಮಯ ಮತ್ತು ಭಾರವಾಗಬಹುದು.", "exclcustom": "ಇನ್ನಷ್ಟು ಹೊರಗಿಡಿ (ಪ್ರತಿ ಸಾಲಿಗೆ ಒಂದು ಮಾರ್ಗ):", "exclexp": "ಈ ಸಿಸ್ಟಂ ಫೋಲ್ಡರ್‌ಗಳನ್ನು ಸ್ಕ್ಯಾನ್‌ನಲ್ಲಿ ಯಾವಾಗಲೂ ಬಿಡಲಾಗುತ್ತದೆ.", "filesafter": "ಪ್ರಸ್ತುತ ಸ್ಕ್ಯಾನ್ ಮುಗಿದ ಮೇಲೆ ಫೈಲ್‌ಗಳು ತೋರಿಸಲ್ಪಡುತ್ತವೆ.", "filescap": "ಈ ಫೋಲ್ಡರ್‌ನ ದೊಡ್ಡ 1000 ಫೈಲ್‌ಗಳಿಗೆ ಪ್ರದರ್ಶನ ಸೀಮಿತ.", "filterexp": "ಪ್ರದರ್ಶನವನ್ನು ಮಾತ್ರ ಫಿಲ್ಟರ್ ಮಾಡುತ್ತದೆ, ಓದಲು. ಸ್ಕ್ಯಾನ್ ವೇಗಗೊಳಿಸುವುದಿಲ್ಲ: ಗಾತ್ರಗಳು ಯಾವಾಗಲೂ ಪೂರ್ಣವಾಗಿ ಲೆಕ್ಕ. ಯಾವಾಗ ಬೇಕಾದರೂ ಬದಲಾಯಿಸಬಹುದು.", "hiddenmsg": "ಮಿತಿಗಿಂತ ಕೆಳಗಿನ {n} ಅಂಶ(ಗಳು) ಪ್ರದರ್ಶನ ಫಿಲ್ಟರ್‌ನಿಂದ ಮರೆ.", "modalhint": "ವಿಶ್ಲೇಷಿಸಬೇಕಾದ ಫೋಲ್ಡರ್ ಆಯ್ಕೆಮಾಡಿ, ನಂತರ ಆಳ ಮತ್ತು ಪ್ರದರ್ಶನ ಹೊಂದಿಸಿ.", "pickerfail": "ಈ ಯಂತ್ರದಲ್ಲಿ ಸ್ಥಳೀಯ ಫೋಲ್ಡರ್ ಆಯ್ಕೆ ತೆರೆಯಲಾಗಲಿಲ್ಲ. ಮಾರ್ಗವನ್ನು ನೇರವಾಗಿ ಟೈಪ್ ಮಾಡಿ.", "queuedmsg": "ಫೋಲ್ಡರ್ ಸರತಿಯಲ್ಲಿ ({n} ಕಾಯುತ್ತಿವೆ). ಪ್ರಸ್ತುತ ಸ್ಕ್ಯಾನ್ ಕೊನೆಯಲ್ಲಿ ಸ್ಕ್ಯಾನ್."},ml:{"sub": "ഡിസ്ക് ഉപയോഗ വിശകലനി", "newscan": "പുതിയ സ്കാൻ", "total": "ആകെ:", "newanalysis": "പുതിയ വിശകലനം", "folderlabel": "വിശകലനം ചെയ്യേണ്ട ഫോൾഡർ", "browse": "ബ്രൗസ്...", "quickaccess": "വേഗത്തിലുള്ള ആക്സസ്", "recent": "സമീപകാലം", "drives": "ഡ്രൈവുകൾ", "depthlabel": "പര്യവേക്ഷണ ആഴം:", "unlimited": "പരിധിയില്ലാത്തത് (എല്ലാം സ്കാൻ ചെയ്യുന്നു, വളരെ സമയമെടുത്തേക്കാം)", "displaylabel": "പ്രദർശനം: ചെറിയ ഇനങ്ങൾ മറയ്ക്കുക", "showall": "എല്ലാം കാണിക്കുക", "analyze": "വിശകലനം", "quit": "സെർവർ അടയ്ക്കുക", "legend": "സൂചിക", "help": "സഹായം", "stop": "നിർത്തുക", "choose": "ഈ ഫോൾഡർ തിരഞ്ഞെടുക്കുക", "cancel": "റദ്ദാക്കുക", "sortby": "അടുക്കുക", "byname": "പേര്", "bysize": "വലുപ്പം", "scangrey": "ചാര ഫോൾഡറുകൾ സ്കാൻ ചെയ്യുക", "leggreen": "സ്കാൻ ചെയ്ത ഫോൾഡർ, വലുപ്പം അറിയാം", "legorange": "നിലവിലെ സ്കാനിൽ ആസൂത്രണം ചെയ്ത ഫോൾഡർ", "legblue": "ക്യൂവിലുള്ള ഫോൾഡർ (നിലവിലെ സ്കാന് ശേഷം)", "legspin": "സ്കാൻ ചെയ്യുന്ന ഫോൾഡർ", "leggrey": "സ്കാൻ ചെയ്തിട്ടില്ല, ഒഴിവാക്കി, ജംഗ്ഷൻ അല്ലെങ്കിൽ പരിരക്ഷിതം. സ്കാൻ ചെയ്യാൻ ക്ലിക്കുചെയ്യുക (2 തലം)", "legfile": "ഫയൽ", "legbar": "നിലവിലെ ഫോൾഡർ വലുപ്പത്തിലെ പങ്ക്", "helpnav": "നാവിഗേഷൻ: പ്രവേശിക്കാൻ ഫോൾഡറിൽ ക്ലിക്കുചെയ്യുക, മുകളിലേക്ക് പോകാൻ മുകളിലെ പാത.", "helpgrey": "ചാര ഫോൾഡറുകൾ: സ്കാൻ ചെയ്യാൻ ക്ലിക്കുചെയ്യുക (2 തലം കൂടി). സ്കാൻ സമയത്ത് അവ ക്യൂവിൽ (നീല) ചേർന്ന് അവസാനം സ്കാൻ ചെയ്യുന്നു.", "helpfiles": "ഫയലുകൾ: ഫോൾഡർ തുറക്കുമ്പോൾ കാണിക്കുന്നു, വലുപ്പം അനുസരിച്ച് അടുക്കുന്നു, വലിയ 1000 വരെ.", "helpfilter": "പ്രദർശന ഫിൽട്ടർ (പുതിയ സ്കാൻ ബട്ടൺ): സ്കാൻ മാറ്റാതെ ചെറിയ ഇനങ്ങൾ മറയ്ക്കുന്നു.", "helpone": "ഒരു സമയത്ത് ഒരു സ്കാൻ: രണ്ട് സമാന്തര ഡിസ്ക് സ്കാനുകൾ പരസ്പരം മന്ദഗതിയിലാക്കും.", "noanalysis": "വിശകലനം ഇല്ല", "theme": "ഇളം / ഇരുണ്ട തീം", "ready": "തയ്യാർ.", "launching": "സ്കാൻ ആരംഭിക്കുന്നു...", "done": "സ്കാൻ പൂർത്തിയായി. സ്വതന്ത്രമായി നാവിഗേറ്റ് ചെയ്യുക.", "inprogress": "പുരോഗമിക്കുന്നു:", "scanningtag": "സ്കാൻ ചെയ്യുന്നു", "parent": ".. (മാതൃ ഫോൾഡർ)", "scandots": "സ്കാൻ ചെയ്യുന്നു...", "empty": "ശൂന്യമായ ഫോൾഡർ", "loadingfiles": "ഫയലുകൾ ലോഡ് ചെയ്യുന്നു...", "excluded": "ഒഴിവാക്കി", "junction": "ജംഗ്ഷൻ", "unscanned": "സ്കാൻ ചെയ്തിട്ടില്ല", "protected": "പരിരക്ഷിതം", "error": "പിശക്", "scanshort": "സ്കാൻ...", "queued": "ക്യൂവിൽ", "levels": "തലങ്ങൾ", "hide1": "1 MB താഴെ മറയ്ക്കുക", "hide100": "100 MB താഴെ മറയ്ക്കുക", "hide1g": "1 GB താഴെ മറയ്ക്കുക", "genby": "സൃഷ്ടിച്ചത്:", "invalidpath": "അസാധുവായ അല്ലെങ്കിൽ പ്രവേശിക്കാനാകാത്ത പാത.", "windowopen": "വിൻഡോ തുറന്നു...", "interrupted": "സ്കാൻ തടസ്സപ്പെട്ടു.", "enterpath": "ദയവായി ഒരു പാത നൽകുക.", "serverstopped": "സെർവർ നിർത്തി. ഈ ടാബ് അടയ്ക്കാം.", "you": "(നിങ്ങൾ)", "network": "(നെറ്റ്‌വർക്ക്)", "drivefree": "സൗജന്യം", "removeone": "നീക്കം ചെയ്യുക", "clearhist": "ചരിത്രം മായ്ക്കുക", "exclusions": "ഒഴിവാക്കലുകൾ", "scandepthfull": "ആഴം: പരിധിയില്ലാത്തത്", "scandepthn": "ആഴം: {n}", "browsetitle": "ബ്രൗസ്", "upfolder": "മാതൃ ഫോൾഡർ", "clickscan": "ഈ ഫോൾഡർ സ്കാൻ ചെയ്യാൻ ക്ലിക്കുചെയ്യുക (2 തലം)", "depthhigh": "വിശദം: പല തലങ്ങൾ മുൻകൂട്ടി ലോഡ്, കാണിക്കാൻ ഭാരം.", "depthlow": "ലഘുവും വേഗവും: ആദ്യ തലങ്ങൾ മാത്രം ലോഡ്. ഫോൾഡറിൽ ക്ലിക്കുചെയ്ത് താഴേക്ക് പോകുക.", "depthmid": "നല്ല സന്തുലനം: പല തലങ്ങൾ ഒരുമിച്ച് കാണാം, ഇപ്പോഴും സുഗമം.", "depthtail": "വലുപ്പങ്ങൾ എപ്പോഴും കൃത്യം; ഈ ആഴത്തിനപ്പുറം, ചാര ഫോൾഡറിൽ ക്ലിക്കുചെയ്ത് പര്യവേക്ഷണം ചെയ്യുക.", "depthunl": "മുഴുവൻ വൃക്ഷവും നടന്ന് കാണിക്കുന്നു. വലിയ ഡിസ്കിൽ വളരെ ദീർഘവും ഭാരവുമാകാം.", "exclcustom": "കൂടുതൽ ഒഴിവാക്കുക (ഓരോ വരിയിലും ഒരു പാത):", "exclexp": "ഈ സിസ്റ്റം ഫോൾഡറുകൾ സ്കാനിൽ എപ്പോഴും ഒഴിവാക്കുന്നു.", "filesafter": "നിലവിലെ സ്കാൻ കഴിയുമ്പോൾ ഫയലുകൾ കാണിക്കും.", "filescap": "ഈ ഫോൾഡറിലെ വലിയ 1000 ഫയലുകളിലേക്ക് പ്രദർശനം പരിമിതം.", "filterexp": "പ്രദർശനം മാത്രം ഫിൽട്ടർ ചെയ്യുന്നു, വായനയ്ക്ക്. സ്കാൻ വേഗത്തിലാക്കുന്നില്ല: വലുപ്പങ്ങൾ എപ്പോഴും പൂർണമായി കണക്കാക്കുന്നു. എപ്പോൾ വേണമെങ്കിലും മാറ്റാം.", "hiddenmsg": "പരിധിക്ക് താഴെ {n} ഇനം(ങ്ങൾ) പ്രദർശന ഫിൽട്ടർ മറച്ചു.", "modalhint": "വിശകലനം ചെയ്യേണ്ട ഫോൾഡർ തിരഞ്ഞെടുത്ത്, ആഴവും പ്രദർശനവും സജ്ജമാക്കുക.", "pickerfail": "ഈ യന്ത്രത്തിൽ നേറ്റീവ് ഫോൾഡർ പിക്കർ തുറക്കാനായില്ല. പാത നേരിട്ട് ടൈപ്പ് ചെയ്യുക.", "queuedmsg": "ഫോൾഡർ ക്യൂവിൽ ({n} കാത്തിരിക്കുന്നു). നിലവിലെ സ്കാനിന്റെ അവസാനം സ്കാൻ ചെയ്യും."},pa:{"sub": "ਡਿਸਕ ਵਰਤੋਂ ਵਿਸ਼ਲੇਸ਼ਕ", "newscan": "ਨਵਾਂ ਸਕੈਨ", "total": "ਕੁੱਲ:", "newanalysis": "ਨਵਾਂ ਵਿਸ਼ਲੇਸ਼ਣ", "folderlabel": "ਵਿਸ਼ਲੇਸ਼ਣ ਕਰਨ ਵਾਲਾ ਫੋਲਡਰ", "browse": "ਬ੍ਰਾਊਜ਼...", "quickaccess": "ਤੇਜ਼ ਪਹੁੰਚ", "recent": "ਹਾਲੀਆ", "drives": "ਡ੍ਰਾਈਵਾਂ", "depthlabel": "ਖੋਜ ਡੂੰਘਾਈ:", "unlimited": "ਅਸੀਮਤ (ਸਭ ਸਕੈਨ ਕਰਦਾ ਹੈ, ਬਹੁਤ ਲੰਬਾ ਹੋ ਸਕਦਾ ਹੈ)", "displaylabel": "ਦਿਖਾਓ: ਛੋਟੀਆਂ ਆਈਟਮਾਂ ਲੁਕਾਓ", "showall": "ਸਭ ਦਿਖਾਓ", "analyze": "ਵਿਸ਼ਲੇਸ਼ਣ", "quit": "ਸਰਵਰ ਬੰਦ ਕਰੋ", "legend": "ਸੰਕੇਤ", "help": "ਮਦਦ", "stop": "ਰੋਕੋ", "choose": "ਇਹ ਫੋਲਡਰ ਚੁਣੋ", "cancel": "ਰੱਦ ਕਰੋ", "sortby": "ਕ੍ਰਮਬੱਧ", "byname": "ਨਾਮ", "bysize": "ਆਕਾਰ", "scangrey": "ਸਲੇਟੀ ਫੋਲਡਰ ਸਕੈਨ ਕਰੋ", "leggreen": "ਸਕੈਨ ਕੀਤਾ ਫੋਲਡਰ, ਆਕਾਰ ਪਤਾ", "legorange": "ਮੌਜੂਦਾ ਸਕੈਨ ਵਿੱਚ ਯੋਜਨਾਬੱਧ ਫੋਲਡਰ", "legblue": "ਕਤਾਰ ਵਿੱਚ ਫੋਲਡਰ (ਮੌਜੂਦਾ ਸਕੈਨ ਤੋਂ ਬਾਅਦ)", "legspin": "ਸਕੈਨ ਹੋ ਰਿਹਾ ਫੋਲਡਰ", "leggrey": "ਸਕੈਨ ਨਹੀਂ, ਬਾਹਰ, ਜੰਕਸ਼ਨ ਜਾਂ ਸੁਰੱਖਿਅਤ. ਸਕੈਨ ਕਰਨ ਲਈ ਕਲਿੱਕ ਕਰੋ (2 ਪੱਧਰ)", "legfile": "ਫਾਈਲ", "legbar": "ਮੌਜੂਦਾ ਫੋਲਡਰ ਆਕਾਰ ਵਿੱਚ ਹਿੱਸਾ", "helpnav": "ਨੈਵੀਗੇਸ਼ਨ: ਦਾਖਲ ਹੋਣ ਲਈ ਫੋਲਡਰ 'ਤੇ ਕਲਿੱਕ ਕਰੋ, ਉੱਪਰ ਜਾਣ ਲਈ ਉੱਪਰਲਾ ਰਾਹ.", "helpgrey": "ਸਲੇਟੀ ਫੋਲਡਰ: ਸਕੈਨ ਕਰਨ ਲਈ ਕਲਿੱਕ ਕਰੋ (2 ਹੋਰ ਪੱਧਰ). ਸਕੈਨ ਦੌਰਾਨ ਇਹ ਕਤਾਰ (ਨੀਲਾ) ਵਿੱਚ ਜਾਂਦੇ ਹਨ ਅਤੇ ਅੰਤ ਵਿੱਚ ਸਕੈਨ ਹੁੰਦੇ ਹਨ.", "helpfiles": "ਫਾਈਲਾਂ: ਫੋਲਡਰ ਖੋਲ੍ਹਣ 'ਤੇ ਦਿਖਾਈਆਂ ਜਾਂਦੀਆਂ, ਆਕਾਰ ਅਨੁਸਾਰ ਕ੍ਰਮਬੱਧ, ਸਭ ਤੋਂ ਵੱਡੀਆਂ 1000 ਤੱਕ.", "helpfilter": "ਡਿਸਪਲੇ ਫਿਲਟਰ (ਨਵਾਂ ਸਕੈਨ ਬਟਨ): ਸਕੈਨ ਬਦਲੇ ਬਿਨਾਂ ਛੋਟੀਆਂ ਆਈਟਮਾਂ ਲੁਕਾਉਂਦਾ ਹੈ.", "helpone": "ਇੱਕ ਵਾਰ ਇੱਕ ਸਕੈਨ: ਦੋ ਸਮਾਨਾਂਤਰ ਡਿਸਕ ਸਕੈਨ ਇੱਕ-ਦੂਜੇ ਨੂੰ ਹੌਲੀ ਕਰਨਗੇ.", "noanalysis": "ਕੋਈ ਵਿਸ਼ਲੇਸ਼ਣ ਨਹੀਂ", "theme": "ਹਲਕਾ / ਗੂੜ੍ਹਾ ਥੀਮ", "ready": "ਤਿਆਰ.", "launching": "ਸਕੈਨ ਸ਼ੁਰੂ ਹੋ ਰਿਹਾ ਹੈ...", "done": "ਸਕੈਨ ਪੂਰਾ. ਖੁੱਲ੍ਹ ਕੇ ਨੈਵੀਗੇਟ ਕਰੋ.", "inprogress": "ਜਾਰੀ ਹੈ:", "scanningtag": "ਸਕੈਨ ਹੋ ਰਿਹਾ", "parent": ".. (ਮੂਲ ਫੋਲਡਰ)", "scandots": "ਸਕੈਨ ਹੋ ਰਿਹਾ ਹੈ...", "empty": "ਖਾਲੀ ਫੋਲਡਰ", "loadingfiles": "ਫਾਈਲਾਂ ਲੋਡ ਹੋ ਰਹੀਆਂ...", "excluded": "ਬਾਹਰ", "junction": "ਜੰਕਸ਼ਨ", "unscanned": "ਸਕੈਨ ਨਹੀਂ", "protected": "ਸੁਰੱਖਿਅਤ", "error": "ਗਲਤੀ", "scanshort": "ਸਕੈਨ...", "queued": "ਕਤਾਰ ਵਿੱਚ", "levels": "ਪੱਧਰ", "hide1": "1 MB ਹੇਠ ਲੁਕਾਓ", "hide100": "100 MB ਹੇਠ ਲੁਕਾਓ", "hide1g": "1 GB ਹੇਠ ਲੁਕਾਓ", "genby": "ਬਣਾਇਆ:", "invalidpath": "ਅਵੈਧ ਜਾਂ ਪਹੁੰਚ ਤੋਂ ਬਾਹਰ ਰਾਹ.", "windowopen": "ਵਿੰਡੋ ਖੁੱਲ੍ਹੀ...", "interrupted": "ਸਕੈਨ ਰੋਕਿਆ.", "enterpath": "ਕਿਰਪਾ ਕਰਕੇ ਰਾਹ ਦਾਖਲ ਕਰੋ.", "serverstopped": "ਸਰਵਰ ਰੁਕਿਆ. ਤੁਸੀਂ ਇਹ ਟੈਬ ਬੰਦ ਕਰ ਸਕਦੇ ਹੋ.", "you": "(ਤੁਸੀਂ)", "network": "(ਨੈੱਟਵਰਕ)", "drivefree": "ਖਾਲੀ", "removeone": "ਹਟਾਓ", "clearhist": "ਇਤਿਹਾਸ ਸਾਫ਼ ਕਰੋ", "exclusions": "ਬਾਹਰ ਰੱਖੇ", "scandepthfull": "ਡੂੰਘਾਈ: ਅਸੀਮਤ", "scandepthn": "ਡੂੰਘਾਈ: {n}", "browsetitle": "ਬ੍ਰਾਊਜ਼", "upfolder": "ਮੂਲ ਫੋਲਡਰ", "clickscan": "ਇਹ ਫੋਲਡਰ ਸਕੈਨ ਕਰਨ ਲਈ ਕਲਿੱਕ ਕਰੋ (2 ਪੱਧਰ)", "depthhigh": "ਵਿਸਤ੍ਰਿਤ: ਕਈ ਪੱਧਰ ਪਹਿਲਾਂ ਲੋਡ, ਦਿਖਾਉਣਾ ਭਾਰੀ.", "depthlow": "ਹਲਕਾ ਅਤੇ ਤੇਜ਼: ਸਿਰਫ਼ ਪਹਿਲੇ ਪੱਧਰ ਲੋਡ ਹੁੰਦੇ ਹਨ. ਫੋਲਡਰ 'ਤੇ ਕਲਿੱਕ ਕਰਕੇ ਹੇਠਾਂ ਜਾਓ.", "depthmid": "ਵਧੀਆ ਸੰਤੁਲਨ: ਕਈ ਪੱਧਰ ਇੱਕੋ ਵੇਲੇ ਦਿਖਦੇ ਹਨ, ਫਿਰ ਵੀ ਸੁਚਾਰੂ.", "depthtail": "ਆਕਾਰ ਹਮੇਸ਼ਾ ਸਹੀ; ਇਸ ਡੂੰਘਾਈ ਤੋਂ ਅੱਗੇ, ਸਲੇਟੀ ਫੋਲਡਰ 'ਤੇ ਕਲਿੱਕ ਕਰਕੇ ਖੋਜੋ.", "depthunl": "ਪੂਰੇ ਰੁੱਖ ਵਿੱਚੋਂ ਲੰਘਦਾ ਅਤੇ ਦਿਖਾਉਂਦਾ ਹੈ. ਵੱਡੀ ਡਿਸਕ 'ਤੇ ਬਹੁਤ ਲੰਬਾ ਅਤੇ ਭਾਰੀ ਹੋ ਸਕਦਾ ਹੈ.", "exclcustom": "ਹੋਰ ਬਾਹਰ ਰੱਖੋ (ਪ੍ਰਤੀ ਲਾਈਨ ਇੱਕ ਰਾਹ):", "exclexp": "ਇਹ ਸਿਸਟਮ ਫੋਲਡਰ ਸਕੈਨ ਦੌਰਾਨ ਹਮੇਸ਼ਾ ਛੱਡੇ ਜਾਂਦੇ ਹਨ.", "filesafter": "ਮੌਜੂਦਾ ਸਕੈਨ ਖ਼ਤਮ ਹੋਣ 'ਤੇ ਫਾਈਲਾਂ ਦਿਖਾਈਆਂ ਜਾਣਗੀਆਂ.", "filescap": "ਇਸ ਫੋਲਡਰ ਦੀਆਂ ਸਭ ਤੋਂ ਵੱਡੀਆਂ 1000 ਫਾਈਲਾਂ ਤੱਕ ਡਿਸਪਲੇ ਸੀਮਤ.", "filterexp": "ਸਿਰਫ਼ ਡਿਸਪਲੇ ਫਿਲਟਰ ਕਰਦਾ ਹੈ, ਪੜ੍ਹਨਯੋਗਤਾ ਲਈ. ਸਕੈਨ ਤੇਜ਼ ਨਹੀਂ ਕਰਦਾ: ਆਕਾਰ ਹਮੇਸ਼ਾ ਪੂਰੇ ਗਿਣੇ ਜਾਂਦੇ ਹਨ. ਕਦੇ ਵੀ ਬਦਲਿਆ ਜਾ ਸਕਦਾ ਹੈ.", "hiddenmsg": "ਸੀਮਾ ਤੋਂ ਹੇਠਾਂ {n} ਆਈਟਮ(ਾਂ) ਡਿਸਪਲੇ ਫਿਲਟਰ ਨੇ ਲੁਕਾਈਆਂ.", "modalhint": "ਵਿਸ਼ਲੇਸ਼ਣ ਲਈ ਫੋਲਡਰ ਚੁਣੋ, ਫਿਰ ਡੂੰਘਾਈ ਅਤੇ ਡਿਸਪਲੇ ਸੈੱਟ ਕਰੋ.", "pickerfail": "ਇਸ ਮਸ਼ੀਨ 'ਤੇ ਮੂਲ ਫੋਲਡਰ ਚੋਣਕਾਰ ਨਹੀਂ ਖੁੱਲ੍ਹ ਸਕਿਆ. ਰਾਹ ਸਿੱਧਾ ਟਾਈਪ ਕਰੋ.", "queuedmsg": "ਫੋਲਡਰ ਕਤਾਰ ਵਿੱਚ ({n} ਉਡੀਕ ਵਿੱਚ). ਮੌਜੂਦਾ ਸਕੈਨ ਦੇ ਅੰਤ ਵਿੱਚ ਸਕੈਨ ਹੋਵੇਗਾ."},he:{"sub": "מנתח שימוש בדיסק", "newscan": "סריקה חדשה", "total": "סה\"כ:", "newanalysis": "ניתוח חדש", "folderlabel": "תיקייה לניתוח", "browse": "עיון...", "quickaccess": "גישה מהירה", "recent": "אחרונים", "drives": "כוננים", "depthlabel": "עומק חקירה:", "unlimited": "ללא הגבלה (סורק הכול, עלול להימשך זמן רב)", "displaylabel": "תצוגה: הסתר פריטים קטנים", "showall": "הצג הכול", "analyze": "נתח", "quit": "סגור שרת", "legend": "מקרא", "help": "עזרה", "stop": "עצור", "choose": "בחר תיקייה זו", "cancel": "ביטול", "sortby": "מיין", "byname": "שם", "bysize": "גודל", "scangrey": "סרוק תיקיות אפורות", "leggreen": "תיקייה שנסרקה, הגודל ידוע", "legorange": "תיקייה מתוכננת בסריקה הנוכחית", "legblue": "תיקייה בתור (נסרקת אחרי הסריקה הנוכחית)", "legspin": "תיקייה בסריקה", "leggrey": "לא נסרקה, הוחרגה, צומת או מוגנת. לחץ כדי לסרוק (2 רמות)", "legfile": "קובץ", "legbar": "חלק מגודל התיקייה הנוכחית", "helpnav": "ניווט: לחץ על תיקייה כדי להיכנס, פירורי הלחם למעלה כדי לעלות.", "helpgrey": "תיקיות אפורות: לחץ כדי לסרוק (2 רמות נוספות). במהלך סריקה הן נכנסות לתור (כחול) ונסרקות בסוף.", "helpfiles": "קבצים: מוצגים כשפותחים תיקייה, ממוינים לפי גודל, מוגבלים ל-1000 הגדולים ביותר.", "helpfilter": "מסנן תצוגה (כפתור סריקה חדשה): מסתיר פריטים קטנים לקריאות, בלי לשנות את הסריקה.", "helpone": "סריקה אחת בכל פעם: שתי סריקות דיסק מקבילות יאטו זו את זו.", "noanalysis": "אין ניתוח", "theme": "ערכת נושא בהירה / כהה", "ready": "מוכן.", "launching": "מתחיל סריקה...", "done": "הסריקה הושלמה. נווט בחופשיות.", "inprogress": "בתהליך:", "scanningtag": "סורק", "parent": ".. (תיקיית אב)", "scandots": "סורק...", "empty": "תיקייה ריקה", "loadingfiles": "טוען קבצים...", "excluded": "הוחרג", "junction": "צומת", "unscanned": "לא נסרק", "protected": "מוגן", "error": "שגיאה", "scanshort": "סריקה...", "queued": "בתור", "levels": "רמות", "hide1": "הסתר מתחת ל-1 MB", "hide100": "הסתר מתחת ל-100 MB", "hide1g": "הסתר מתחת ל-1 GB", "genby": "נוצר על ידי:", "invalidpath": "נתיב לא חוקי או לא נגיש.", "windowopen": "החלון נפתח...", "interrupted": "הסריקה הופסקה.", "enterpath": "הזן נתיב.", "serverstopped": "השרת הופסק. אפשר לסגור כרטיסייה זו.", "you": "(אתה)", "network": "(רשת)", "drivefree": "פנוי", "removeone": "הסר", "clearhist": "נקה היסטוריה", "exclusions": "החרגות", "scandepthfull": "עומק: ללא הגבלה", "scandepthn": "עומק: {n}", "browsetitle": "עיון", "upfolder": "תיקיית אב", "clickscan": "לחץ כדי לסרוק תיקייה זו (2 רמות)", "depthhigh": "מפורט: רמות רבות נטענות מראש, כבד יותר להצגה.", "depthlow": "קל ומהיר: רק הרמות הראשונות נטענות. רד למטה בלחיצה על תיקייה.", "depthmid": "איזון טוב: כמה רמות נראות בבת אחת, עדיין חלק.", "depthtail": "הגדלים תמיד מדויקים; מעבר לעומק זה, לחץ על תיקייה אפורה כדי לחקור אותה.", "depthunl": "עובר ומציג את כל העץ. עלול להיות ארוך וכבד מאוד בדיסק גדול.", "exclcustom": "החרג גם (נתיב אחד בכל שורה):", "exclexp": "תיקיות מערכת אלו תמיד מדולגות במהלך הסריקה.", "filesafter": "קבצים יוצגו בסיום הסריקה הנוכחית.", "filescap": "התצוגה מוגבלת ל-1000 הקבצים הגדולים ביותר בתיקייה זו.", "filterexp": "מסנן את התצוגה בלבד, לקריאות. לא מאיץ את הסריקה: הגדלים תמיד מחושבים במלואם. ניתן לשנות בכל עת.", "hiddenmsg": "{n} פריטים מתחת לסף הוסתרו על ידי מסנן התצוגה.", "modalhint": "בחר את התיקייה לניתוח, ואז הגדר עומק ותצוגה.", "pickerfail": "בורר התיקיות המובנה לא נפתח במחשב זה. הקלד את הנתיב ישירות.", "queuedmsg": "התיקייה נכנסה לתור ({n} ממתינות). תיסרק בסיום הסריקה הנוכחית."},ha:{"sub": "Manazarcin Amfani da Faifai", "newscan": "Sabon bincike", "total": "Jimla:", "newanalysis": "Sabon bincike", "folderlabel": "Babbar fayil don bincike", "browse": "Bincika...", "quickaccess": "Sauri shiga", "recent": "Na baya-bayan nan", "drives": "Na'urori", "depthlabel": "Zurfin bincike:", "unlimited": "Mara iyaka (yana bincika komai, yana iya daukar lokaci mai tsawo)", "displaylabel": "Nuni: boye kananan abubuwa", "showall": "Nuna duka", "analyze": "Bincika", "quit": "Rufe uwar garke", "legend": "Bayani", "help": "Taimako", "stop": "Tsaya", "choose": "Zabi wannan babbar fayil", "cancel": "Soke", "sortby": "Tsara", "byname": "suna", "bysize": "girma", "scangrey": "Bincika manyan fayil masu launin toka", "leggreen": "Babbar fayil da aka bincika, an san girma", "legorange": "Babbar fayil da aka tsara a binciken yanzu", "legblue": "Babbar fayil a jerin gwano (za a bincika bayan binciken yanzu)", "legspin": "Ana bincika babbar fayil", "leggrey": "Ba a bincika ba, an cire, junction ko kariya. Danna don bincika (matakai 2)", "legfile": "Fayil", "legbar": "Rabo a girman babbar fayil na yanzu", "helpnav": "Kewayawa: danna babbar fayil don shiga, hanyar sama don hawa.", "helpgrey": "Manyan fayil masu launin toka: danna don bincika (karin matakai 2). Yayin bincike suna shiga jerin gwano (shudi) kuma ana bincika su a karshe.", "helpfiles": "Fayiloli: ana nunawa lokacin da ka bude babbar fayil, an tsara su bisa girma, iyaka 1000 mafi girma.", "helpfilter": "Matatar nuni (maballin Sabon bincike): yana boye kananan abubuwa don karantawa, ba tare da canza bincike ba.", "helpone": "Bincike daya a lokaci: bincike-bincike guda biyu na faifai a lokaci daya za su rage juna.", "noanalysis": "Babu bincike", "theme": "Jigo mai haske / duhu", "ready": "A shirye.", "launching": "Ana fara bincike...", "done": "An gama bincike. Yi kewayawa cikin yanci.", "inprogress": "Ana ci gaba:", "scanningtag": "ana bincika", "parent": ".. (babbar fayil na sama)", "scandots": "Ana bincika...", "empty": "Babbar fayil babu komai", "loadingfiles": "Ana loda fayiloli...", "excluded": "an cire", "junction": "junction", "unscanned": "ba a bincika ba", "protected": "kariya", "error": "kuskure", "scanshort": "bincike...", "queued": "a jerin gwano", "levels": "matakai", "hide1": "Boye kasa da 1 MB", "hide100": "Boye kasa da 100 MB", "hide1g": "Boye kasa da 1 GB", "genby": "An kirkira ta:", "invalidpath": "Hanya mara inganci ko wadda ba a iya kaiwa.", "windowopen": "An bude taga...", "interrupted": "An dakatar da bincike.", "enterpath": "Da fatan shigar da hanya.", "serverstopped": "An dakatar da uwar garke. Kana iya rufe wannan shafin.", "you": "(kai)", "network": "(hanyar sadarwa)", "drivefree": "kyauta", "removeone": "Cire", "clearhist": "Share tarihi", "exclusions": "Abubuwan da aka cire", "scandepthfull": "Zurfi: mara iyaka", "scandepthn": "Zurfi: {n}", "browsetitle": "Bincika", "upfolder": "Babbar fayil na sama", "clickscan": "Danna don bincika wannan babbar fayil (matakai 2)", "depthhigh": "Cikakke: matakai da yawa an loda tun farko, nauyi wajen nunawa.", "depthlow": "Sauki da sauri: matakan farko kawai ake lodawa. Sauka ta danna babbar fayil.", "depthmid": "Daidaito mai kyau: matakai da dama ana gani a lokaci daya, har yanzu a hankali.", "depthtail": "Girma koyaushe daidai ne; bayan wannan zurfi, danna babbar fayil mai launin toka don bincika.", "depthunl": "Yana ratsa ya nuna dukkan bishiya. Zai iya zama mai tsawo da nauyi sosai a babban faifai.", "exclcustom": "Cire kuma (hanya daya a kowane layi):", "exclexp": "Wadannan manyan fayil na tsarin ana tsallake su koyaushe yayin bincike.", "filesafter": "Fayiloli za su bayyana a karshen binciken yanzu.", "filescap": "Nuni ya takaita ga fayiloli 1000 mafi girma a wannan babbar fayil.", "filterexp": "Yana tace nuni kawai, don karantawa. Baya hanzarta bincike: girma koyaushe ana lissafa gaba daya. Ana iya canzawa a kowane lokaci.", "hiddenmsg": "Abubuwa {n} kasa da iyaka an boye ta matatar nuni.", "modalhint": "Zabi babbar fayil don bincike, sannan saita zurfi da nuni.", "pickerfail": "Ba a iya bude zabin babbar fayil na asali a wannan na'ura ba. Rubuta hanyar kai tsaye.", "queuedmsg": "Babbar fayil a jerin gwano ({n} suna jira). Za a bincika a karshen binciken yanzu."},my:{"sub": "ဒစ်ခ်အသုံးပြုမှု ခွဲခြမ်းစိတ်ဖြာစနစ်", "newscan": "စကင်ဖတ်အသစ်", "total": "စုစုပေါင်း-", "newanalysis": "ခွဲခြမ်းစိတ်ဖြာမှုအသစ်", "folderlabel": "ခွဲခြမ်းစိတ်ဖြာမည့် ဖိုင်တွဲ", "browse": "ရှာဖွေ...", "quickaccess": "အမြန်ဝင်ရောက်", "recent": "မကြာသေးမီ", "drives": "ဒရိုက်များ", "depthlabel": "စူးစမ်းမှုအနက်-", "unlimited": "အကန့်အသတ်မဲ့ (အားလုံးစကင်ဖတ်သည်၊ အလွန်ကြာနိုင်သည်)", "displaylabel": "ပြသမှု- အသေးအဖွဲများ ဖျောက်ရန်", "showall": "အားလုံးပြရန်", "analyze": "ခွဲခြမ်းစိတ်ဖြာ", "quit": "ဆာဗာပိတ်ရန်", "legend": "အညွှန်း", "help": "အကူအညီ", "stop": "ရပ်ရန်", "choose": "ဤဖိုင်တွဲကို ရွေးရန်", "cancel": "မလုပ်တော့ပါ", "sortby": "စဉ်ရန်", "byname": "အမည်", "bysize": "အရွယ်", "scangrey": "မီးခိုးရောင်ဖိုင်တွဲများ စကင်ဖတ်ရန်", "leggreen": "စကင်ဖတ်ပြီး ဖိုင်တွဲ၊ အရွယ်သိရှိ", "legorange": "လက်ရှိစကင်ဖတ်မှုတွင် စီစဉ်ထားသော ဖိုင်တွဲ", "legblue": "တန်းစီထားသော ဖိုင်တွဲ (လက်ရှိစကင်ဖတ်ပြီးမှ)", "legspin": "စကင်ဖတ်နေသော ဖိုင်တွဲ", "leggrey": "မစကင်ဖတ်ရသေး၊ ချန်လှပ်၊ junction သို့မဟုတ် ကာကွယ်ထား။ စကင်ဖတ်ရန် နှိပ်ပါ (အဆင့် 2)", "legfile": "ဖိုင်", "legbar": "လက်ရှိဖိုင်တွဲအရွယ်၏ အချိုး", "helpnav": "လမ်းညွှန်- ဝင်ရန် ဖိုင်တွဲကို နှိပ်ပါ၊ အပေါ်သွားရန် အထက်လမ်းကြောင်း။", "helpgrey": "မီးခိုးရောင်ဖိုင်တွဲများ- စကင်ဖတ်ရန် နှိပ်ပါ (နောက်ထပ်အဆင့် 2)။ စကင်ဖတ်နေစဉ် တန်းစီ (အပြာ) ဝင်ပြီး အဆုံးတွင် စကင်ဖတ်သည်။", "helpfiles": "ဖိုင်များ- ဖိုင်တွဲဖွင့်သောအခါ ပြသည်၊ အရွယ်အလိုက် စီ၊ အကြီးဆုံး 1000 အထိ။", "helpfilter": "ပြသမှုစစ်ထုတ်ကိရိယာ (စကင်ဖတ်အသစ်ခလုတ်)- စကင်ဖတ်မှုမပြောင်းဘဲ အသေးအဖွဲများကို ဖျောက်သည်။", "helpone": "တစ်ကြိမ်လျှင် စကင်ဖတ်တစ်ခု- ပြိုင်တူဒစ်ခ်စကင်ဖတ်နှစ်ခုသည် အချင်းချင်း နှေးကွေးစေသည်။", "noanalysis": "ခွဲခြမ်းစိတ်ဖြာမှုမရှိ", "theme": "အလင်း / အမှောင် အပြင်အဆင်", "ready": "အသင့်။", "launching": "စကင်ဖတ်စတင်နေသည်...", "done": "စကင်ဖတ်ပြီး။ လွတ်လပ်စွာ လှည့်လည်နိုင်သည်။", "inprogress": "ဆောင်ရွက်ဆဲ-", "scanningtag": "စကင်ဖတ်နေသည်", "parent": ".. (မိဘဖိုင်တွဲ)", "scandots": "စကင်ဖတ်နေသည်...", "empty": "ဗလာဖိုင်တွဲ", "loadingfiles": "ဖိုင်များ ဖွင့်နေသည်...", "excluded": "ချန်လှပ်", "junction": "junction", "unscanned": "မစကင်ဖတ်ရသေး", "protected": "ကာကွယ်ထား", "error": "အမှား", "scanshort": "စကင်...", "queued": "တန်းစီထား", "levels": "အဆင့်", "hide1": "1 MB အောက် ဖျောက်ရန်", "hide100": "100 MB အောက် ဖျောက်ရန်", "hide1g": "1 GB အောက် ဖျောက်ရန်", "genby": "ဖန်တီးသူ-", "invalidpath": "မမှန်ကန်သော သို့မဟုတ် ဝင်ရောက်၍မရသော လမ်းကြောင်း။", "windowopen": "ဝင်းဒိုးဖွင့်ပြီး...", "interrupted": "စကင်ဖတ်မှု ရပ်တန့်ခဲ့သည်။", "enterpath": "လမ်းကြောင်းတစ်ခု ထည့်ပါ။", "serverstopped": "ဆာဗာရပ်ပြီး။ ဤtabကို ပိတ်နိုင်သည်။", "you": "(သင်)", "network": "(ကွန်ရက်)", "drivefree": "အားလပ်", "removeone": "ဖယ်ရှား", "clearhist": "မှတ်တမ်းရှင်း", "exclusions": "ချန်လှပ်ချက်များ", "scandepthfull": "အနက်- အကန့်အသတ်မဲ့", "scandepthn": "အနက်- {n}", "browsetitle": "ရှာဖွေ", "upfolder": "မိဘဖိုင်တွဲ", "clickscan": "ဤဖိုင်တွဲကို စကင်ဖတ်ရန် နှိပ်ပါ (အဆင့် 2)", "depthhigh": "အသေးစိတ်- အဆင့်များစွာ ကြိုတင်ဖွင့်ထား၊ ပြသရန် လေးလံသည်။", "depthlow": "ပေါ့ပါးမြန်ဆန်- ပထမအဆင့်များသာ ဖွင့်သည်။ ဖိုင်တွဲကို နှိပ်၍ အောက်သွားပါ။", "depthmid": "ဟန်ချက်ကောင်း- အဆင့်များစွာ တစ်ပြိုင်နက်မြင်ရ၊ ချောမွေ့ဆဲ။", "depthtail": "အရွယ်များ အမြဲတိကျသည်။ ဤအနက်ကျော်လွန်ပါက မီးခိုးရောင်ဖိုင်တွဲကို နှိပ်၍ စူးစမ်းပါ။", "depthunl": "သစ်ပင်တစ်ခုလုံးကို ဖြတ်သန်းပြသသည်။ ဒစ်ခ်ကြီးတွင် အလွန်ကြာပြီး လေးလံနိုင်သည်။", "exclcustom": "ထပ်ချန်လှပ်ရန် (တစ်ကြောင်းလျှင် လမ်းကြောင်းတစ်ခု)-", "exclexp": "ဤစနစ်ဖိုင်တွဲများကို စကင်ဖတ်စဉ် အမြဲကျော်သည်။", "filesafter": "လက်ရှိစကင်ဖတ်မှု ပြီးမှ ဖိုင်များ ပြမည်။", "filescap": "ဤဖိုင်တွဲ၏ အကြီးဆုံးဖိုင် 1000 အထိသာ ပြသည်။", "filterexp": "ပြသမှုကိုသာ စစ်ထုတ်သည်၊ ဖတ်ရလွယ်ရန်။ စကင်ဖတ်မှု မမြန်စေပါ- အရွယ်များ အမြဲအပြည့်တွက်သည်။ အချိန်မရွေး ပြောင်းနိုင်သည်။", "hiddenmsg": "သတ်မှတ်ချက်အောက် {n} ခုကို ပြသမှုစစ်ထုတ်ကိရိယာက ဖျောက်ထားသည်။", "modalhint": "ခွဲခြမ်းစိတ်ဖြာမည့် ဖိုင်တွဲကို ရွေးပြီး အနက်နှင့် ပြသမှုကို သတ်မှတ်ပါ။", "pickerfail": "ဤစက်တွင် မူလဖိုင်တွဲရွေးချယ်ကိရိယာ ဖွင့်၍မရပါ။ လမ်းကြောင်းကို တိုက်ရိုက်ရိုက်ထည့်ပါ။", "queuedmsg": "ဖိုင်တွဲ တန်းစီပြီး ({n} ခု စောင့်ဆဲ)။ လက်ရှိစကင်ဖတ်မှုအဆုံးတွင် စကင်ဖတ်မည်။"},am:{"sub": "የዲስክ አጠቃቀም ተንታኝ", "newscan": "አዲስ ቅኝት", "total": "ጠቅላላ:", "newanalysis": "አዲስ ትንተና", "folderlabel": "የሚተነተን አቃፊ", "browse": "አስስ...", "quickaccess": "ፈጣን መዳረሻ", "recent": "የቅርብ ጊዜ", "drives": "ድራይቮች", "depthlabel": "የፍለጋ ጥልቀት:", "unlimited": "ያልተገደበ (ሁሉንም ይቃኛል፣ በጣም ረጅም ሊሆን ይችላል)", "displaylabel": "ማሳያ: ትናንሽ ንጥሎችን ደብቅ", "showall": "ሁሉንም አሳይ", "analyze": "ተንትን", "quit": "አገልጋይ ዝጋ", "legend": "መግለጫ", "help": "እገዛ", "stop": "አቁም", "choose": "ይህን አቃፊ ምረጥ", "cancel": "ሰርዝ", "sortby": "ደርድር", "byname": "ስም", "bysize": "መጠን", "scangrey": "ግራጫ አቃፊዎችን ቃኝ", "leggreen": "የተቃኘ አቃፊ፣ መጠኑ የታወቀ", "legorange": "በአሁኑ ቅኝት የታቀደ አቃፊ", "legblue": "በሰልፍ ውስጥ ያለ አቃፊ (ከአሁኑ ቅኝት በኋላ ይቃኛል)", "legspin": "እየተቃኘ ያለ አቃፊ", "leggrey": "ያልተቃኘ፣ የተገለለ፣ ማገናኛ ወይም የተጠበቀ። ለመቃኘት ጠቅ ያድርጉ (2 ደረጃ)", "legfile": "ፋይል", "legbar": "በአሁኑ አቃፊ መጠን ውስጥ ያለ ድርሻ", "helpnav": "አሰሳ፦ ለመግባት አቃፊ ጠቅ ያድርጉ፣ ወደ ላይ ለመውጣት ከላይ ያለውን መንገድ ይጠቀሙ።", "helpgrey": "ግራጫ አቃፊዎች፦ ለመቃኘት ጠቅ ያድርጉ (ተጨማሪ 2 ደረጃ)። በቅኝት ጊዜ ወደ ሰልፍ (ሰማያዊ) ገብተው በመጨረሻ ይቃኛሉ።", "helpfiles": "ፋይሎች፦ አቃፊ ሲከፍቱ ይታያሉ፣ በመጠን ተደርድረው፣ እስከ 1000 ትልቁ ድረስ።", "helpfilter": "የማሳያ ማጣሪያ (አዲስ ቅኝት ቁልፍ)፦ ቅኝቱን ሳይቀይር ትናንሽ ንጥሎችን ይደብቃል።", "helpone": "በአንድ ጊዜ አንድ ቅኝት፦ ሁለት ትይዩ የዲስክ ቅኝቶች እርስ በርስ ያዘገዩ ነበር።", "noanalysis": "ትንተና የለም", "theme": "ብሩህ / ጨለማ ገጽታ", "ready": "ዝግጁ።", "launching": "ቅኝት በመጀመር ላይ...", "done": "ቅኝት ተጠናቀቀ። በነጻነት ይዳስሱ።", "inprogress": "በሂደት ላይ:", "scanningtag": "በመቃኘት ላይ", "parent": ".. (ወላጅ አቃፊ)", "scandots": "በመቃኘት ላይ...", "empty": "ባዶ አቃፊ", "loadingfiles": "ፋይሎችን በመጫን ላይ...", "excluded": "ተገልሏል", "junction": "ማገናኛ", "unscanned": "ያልተቃኘ", "protected": "የተጠበቀ", "error": "ስህተት", "scanshort": "ቅኝት...", "queued": "በሰልፍ", "levels": "ደረጃ", "hide1": "ከ1 ሜባ በታች ደብቅ", "hide100": "ከ100 ሜባ በታች ደብቅ", "hide1g": "ከ1 ጊባ በታች ደብቅ", "genby": "የተፈጠረው በ:", "invalidpath": "ልክ ያልሆነ ወይም የማይደረስ መንገድ።", "windowopen": "መስኮት ተከፍቷል...", "interrupted": "ቅኝት ተቋርጧል።", "enterpath": "እባክዎ መንገድ ያስገቡ።", "serverstopped": "አገልጋይ ቆሟል። ይህን ትር መዝጋት ይችላሉ።", "you": "(እርስዎ)", "network": "(አውታረ መረብ)", "drivefree": "ነጻ", "removeone": "አስወግድ", "clearhist": "ታሪክ አጽዳ", "exclusions": "ማግለያዎች", "scandepthfull": "ጥልቀት: ያልተገደበ", "scandepthn": "ጥልቀት: {n}", "browsetitle": "አስስ", "upfolder": "ወላጅ አቃፊ", "clickscan": "ይህን አቃፊ ለመቃኘት ጠቅ ያድርጉ (2 ደረጃ)", "depthhigh": "ዝርዝር፦ ብዙ ደረጃዎች አስቀድመው ተጭነዋል፣ ለማሳየት ከባድ።", "depthlow": "ቀላል እና ፈጣን፦ የመጀመሪያ ደረጃዎች ብቻ ይጫናሉ። አቃፊ ጠቅ በማድረግ ወደ ታች ይውረዱ።", "depthmid": "ጥሩ ሚዛን፦ ብዙ ደረጃዎች በአንድ ጊዜ ይታያሉ፣ አሁንም ለስላሳ።", "depthtail": "መጠኖች ሁልጊዜ ትክክል ናቸው፤ ከዚህ ጥልቀት ባሻገር፣ ግራጫ አቃፊ ጠቅ በማድረግ ያስሱ።", "depthunl": "ሙሉውን ዛፍ ያልፋል እና ያሳያል። በትልቅ ዲስክ ላይ በጣም ረጅም እና ከባድ ሊሆን ይችላል።", "exclcustom": "በተጨማሪ ያግልሉ (በመስመር አንድ መንገድ)፦", "exclexp": "እነዚህ የስርዓት አቃፊዎች በቅኝት ጊዜ ሁልጊዜ ይዘለላሉ።", "filesafter": "ፋይሎች የአሁኑ ቅኝት ሲጠናቀቅ ይታያሉ።", "filescap": "ማሳያው በዚህ አቃፊ ውስጥ ላሉ 1000 ትልቅ ፋይሎች ተወስኗል።", "filterexp": "ማሳያውን ብቻ ያጣራል፣ ለንባብ። ቅኝቱን አያፋጥንም፦ መጠኖች ሁልጊዜ ሙሉ በሙሉ ይሰላሉ። በማንኛውም ጊዜ ሊቀየር ይችላል።", "hiddenmsg": "ከመነሻው በታች {n} ንጥሎች በማሳያ ማጣሪያ ተደብቀዋል።", "modalhint": "የሚተነተነውን አቃፊ ይምረጡ፣ ከዚያ ጥልቀት እና ማሳያ ያዘጋጁ።", "pickerfail": "የተለመደ የአቃፊ መራጭ በዚህ ማሽን ላይ ሊከፈት አልቻለም። መንገዱን በቀጥታ ይተይቡ።", "queuedmsg": "አቃፊ በሰልፍ ተቀምጧል ({n} በመጠበቅ ላይ)። በአሁኑ ቅኝት መጨረሻ ይቃኛል።"},km:{"sub": "កម្មវិធីវិភាគការប្រើប្រាស់ថាស", "newscan": "ស្កេនថ្មី", "total": "សរុប៖", "newanalysis": "ការវិភាគថ្មី", "folderlabel": "ថតដែលត្រូវវិភាគ", "browse": "រកមើល...", "quickaccess": "ចូលដំណើរការរហ័ស", "recent": "ថ្មីៗ", "drives": "ដ្រាយ", "depthlabel": "ជម្រៅស្វែងរក៖", "unlimited": "គ្មានដែនកំណត់ (ស្កេនទាំងអស់ អាចយូរណាស់)", "displaylabel": "បង្ហាញ៖ លាក់ធាតុតូចៗ", "showall": "បង្ហាញទាំងអស់", "analyze": "វិភាគ", "quit": "បិទម៉ាស៊ីនមេ", "legend": "តំណាង", "help": "ជំនួយ", "stop": "បញ្ឈប់", "choose": "ជ្រើសថតនេះ", "cancel": "បោះបង់", "sortby": "តម្រៀប", "byname": "ឈ្មោះ", "bysize": "ទំហំ", "scangrey": "ស្កេនថតពណ៌ប្រផេះ", "leggreen": "ថតដែលបានស្កេន ដឹងទំហំ", "legorange": "ថតដែលបានគ្រោងក្នុងការស្កេនបច្ចុប្បន្ន", "legblue": "ថតក្នុងជួរ (ស្កេនក្រោយការស្កេនបច្ចុប្បន្ន)", "legspin": "ថតកំពុងស្កេន", "leggrey": "មិនទាន់ស្កេន បានដកចេញ junction ឬការពារ។ ចុចដើម្បីស្កេន (2 កម្រិត)", "legfile": "ឯកសារ", "legbar": "ចំណែកនៃទំហំថតបច្ចុប្បន្ន", "helpnav": "រុករក៖ ចុចថតដើម្បីចូល ផ្លូវខាងលើដើម្បីឡើងលើ។", "helpgrey": "ថតពណ៌ប្រផេះ៖ ចុចដើម្បីស្កេន (បន្ថែម 2 កម្រិត)។ ពេលស្កេន វាចូលជួរ (ខៀវ) ហើយស្កេននៅចុងបញ្ចប់។", "helpfiles": "ឯកសារ៖ បង្ហាញពេលអ្នកបើកថត តម្រៀបតាមទំហំ កំណត់ត្រឹម 1000 ធំបំផុត។", "helpfilter": "តម្រងបង្ហាញ (ប៊ូតុងស្កេនថ្មី)៖ លាក់ធាតុតូចៗសម្រាប់ភាពងាយអាន ដោយមិនប្តូរការស្កេន។", "helpone": "ស្កេនម្តងមួយ៖ ការស្កេនថាសពីរស្របគ្នានឹងធ្វើឲ្យគ្នាទៅវិញទៅមកយឺត។", "noanalysis": "គ្មានការវិភាគ", "theme": "រូបរាងភ្លឺ / ងងឹត", "ready": "រួចរាល់។", "launching": "កំពុងចាប់ផ្តើមស្កេន...", "done": "ស្កេនរួច។ រុករកដោយសេរី។", "inprogress": "កំពុងដំណើរការ៖", "scanningtag": "កំពុងស្កេន", "parent": ".. (ថតមេ)", "scandots": "កំពុងស្កេន...", "empty": "ថតទទេ", "loadingfiles": "កំពុងផ្ទុកឯកសារ...", "excluded": "បានដកចេញ", "junction": "junction", "unscanned": "មិនទាន់ស្កេន", "protected": "បានការពារ", "error": "កំហុស", "scanshort": "ស្កេន...", "queued": "ក្នុងជួរ", "levels": "កម្រិត", "hide1": "លាក់ក្រោម 1 MB", "hide100": "លាក់ក្រោម 100 MB", "hide1g": "លាក់ក្រោម 1 GB", "genby": "បង្កើតដោយ៖", "invalidpath": "ផ្លូវមិនត្រឹមត្រូវ ឬចូលមិនបាន។", "windowopen": "បានបើកបង្អួច...", "interrupted": "ការស្កេនត្រូវបានបញ្ឈប់។", "enterpath": "សូមបញ្ចូលផ្លូវ។", "serverstopped": "ម៉ាស៊ីនមេបានបញ្ឈប់។ អ្នកអាចបិទផ្ទាំងនេះ។", "you": "(អ្នក)", "network": "(បណ្តាញ)", "drivefree": "ទំនេរ", "removeone": "លុប", "clearhist": "សម្អាតប្រវត្តិ", "exclusions": "ការដកចេញ", "scandepthfull": "ជម្រៅ៖ គ្មានដែនកំណត់", "scandepthn": "ជម្រៅ៖ {n}", "browsetitle": "រកមើល", "upfolder": "ថតមេ", "clickscan": "ចុចដើម្បីស្កេនថតនេះ (2 កម្រិត)", "depthhigh": "លម្អិត៖ កម្រិតច្រើនផ្ទុកជាមុន ធ្ងន់ក្នុងការបង្ហាញ។", "depthlow": "ស្រាល និងលឿន៖ ផ្ទុកតែកម្រិតដំបូង។ ចុះក្រោមដោយចុចថត។", "depthmid": "តុល្យភាពល្អ៖ កម្រិតច្រើនមើលឃើញក្នុងពេលតែមួយ នៅតែរលូន។", "depthtail": "ទំហំត្រឹមត្រូវជានិច្ច លើសពីជម្រៅនេះ ចុចថតពណ៌ប្រផេះដើម្បីស្វែងរក។", "depthunl": "ដើរ និងបង្ហាញដើមឈើទាំងមូល។ អាចយូរ និងធ្ងន់ខ្លាំងលើថាសធំ។", "exclcustom": "ដកចេញបន្ថែម (មួយផ្លូវក្នុងមួយបន្ទាត់)៖", "exclexp": "ថតប្រព័ន្ធទាំងនេះត្រូវបានរំលងជានិច្ចពេលស្កេន។", "filesafter": "ឯកសារនឹងបង្ហាញនៅចុងបញ្ចប់នៃការស្កេនបច្ចុប្បន្ន។", "filescap": "ការបង្ហាញកំណត់ត្រឹមឯកសារធំបំផុត 1000 ក្នុងថតនេះ។", "filterexp": "ត្រងតែការបង្ហាញ ដើម្បីងាយអាន។ មិនធ្វើឲ្យស្កេនលឿនទេ៖ ទំហំគណនាពេញលេញជានិច្ច។ អាចប្តូរបានគ្រប់ពេល។", "hiddenmsg": "ធាតុ {n} ក្រោមកម្រិតត្រូវបានលាក់ដោយតម្រងបង្ហាញ។", "modalhint": "ជ្រើសថតដែលត្រូវវិភាគ រួចកំណត់ជម្រៅ និងការបង្ហាញ។", "pickerfail": "មិនអាចបើកឧបករណ៍ជ្រើសថតដើមនៅលើម៉ាស៊ីននេះទេ។ វាយផ្លូវដោយផ្ទាល់។", "queuedmsg": "ថតបានចូលជួរ ({n} កំពុងរង់ចាំ)។ នឹងស្កេននៅចុងបញ្ចប់នៃការស្កេនបច្ចុប្បន្ន។"}
};
function t(k){var d=I18N[LANG]||I18N.en;var v=(d&&d[k]!=null)?d[k]:(I18N.en[k]!=null?I18N.en[k]:k);return v;}
function applyI18n(){
  try{
    var els=document.querySelectorAll('[data-i18n]');for(var i=0;i<els.length;i++){els[i].textContent=t(els[i].getAttribute('data-i18n'));}
    var ph=document.querySelectorAll('[data-i18n-ph]');for(var j=0;j<ph.length;j++){ph[j].placeholder=t(ph[j].getAttribute('data-i18n-ph'));}
    var ti=document.querySelectorAll('[data-i18n-title]');for(var k=0;k<ti.length;k++){ti[k].title=t(ti[k].getAttribute('data-i18n-title'));}
  }catch(e){}
}
var MODES=__MODES_JSON__;
var selMode=3, es=null, DISPLAY_MIN=0;

function q(id){return document.getElementById(id);}
function toggleTheme(){var h=document.documentElement;h.setAttribute('data-theme',h.getAttribute('data-theme')==='dark'?'light':'dark');}
function showSettings(){q('overlay').classList.add('on');try{q('setclose').style.display=ROOT?'':'none';}catch(e){}}
function closeModal(){if(ROOT){q('overlay').classList.remove('on');}}
function hideStop(){try{q('stopscan').style.display='none';}catch(e){}}
function stopScan(){try{if(es)es.close();}catch(e){}try{if(ses)ses.close();}catch(e){}stopTimer();BUSY=false;SUBSCAN=null;SUBPARENT=null;subQueue=[];SCANNING=null;ACTIVE_MAP={};HOMEVIEW=null;try{q('gbar').classList.remove('loading');}catch(e){}q('dscan').innerHTML='';q('dcur').textContent=t('interrupted');hideStop();renderTree();showSettings();}
document.addEventListener('keydown',function(e){if(e.key==='Escape'&&ROOT&&q('overlay').classList.contains('on')){closeModal();}});
var pathCk=null;
function setPathValid(v){var el=q('pathok');if(!el)return;if(v===null){el.className='pathok';el.textContent='';}else if(v){el.className='pathok ok';el.textContent='\u2713';}else{el.className='pathok bad';el.textContent='\u2717';}}
function onPathInput(){setPathValid(null);if(pathCk)clearTimeout(pathCk);var p=q('path').value.trim();if(!p)return;pathCk=setTimeout(function(){fetch('/api/check?token='+TOKEN+'&path='+encodeURIComponent(p)).then(function(r){return r.json();}).then(function(c){if(q('path').value.trim()===p)setPathValid(!!(c&&c.ok));}).catch(function(){});},400);}
function saveSettings(path){try{localStorage.setItem('psncdu_prefs',JSON.stringify({path:path,depth:q('depth').value,unl:q('depthUnl').checked,min:q('minsize').value}));}catch(e){}}
function restoreSettings(){try{var p=JSON.parse(localStorage.getItem('psncdu_prefs')||'null');if(!p)return;if(p.path)q('path').value=p.path;if(p.depth)q('depth').value=p.depth;if(typeof p.unl==='boolean')q('depthUnl').checked=p.unl;if(p.min!=null)q('minsize').value=p.min;DISPLAY_MIN=parseInt(q('minsize').value,10)||0;}catch(e){}}
function setLang(l){if(!I18N[l])return;LANG=l;try{localStorage.setItem('psncdu_lang',l);}catch(e){}document.documentElement.lang=l;document.documentElement.dir=/^(ar|ur|fa|he)$/.test(l)?'rtl':'ltr';applyI18n();try{updateSortBtn();}catch(e){}try{if(BUSY&&q('stdepth').textContent){q('stdepth').textContent=q('depthUnl').checked?t('scandepthfull'):t('scandepthn').replace('{n}',q('depth').value);}}catch(e){}onDepth();if(CUR)renderTree();}
var BROWSE_PATH='';
function openBrowse(){q('browseOverlay').classList.add('on');browseTo((q('path').value||'').trim());}
function closeBrowse(){q('browseOverlay').classList.remove('on');}
function chooseBrowse(){if(BROWSE_PATH){q('path').value=BROWSE_PATH;onPathInput();}closeBrowse();}
function browseTo(p){q('browseList').innerHTML='<div class="tempty">...</div>';fetch('/api/browse?token='+TOKEN+'&path='+encodeURIComponent(p||'')).then(function(r){return r.json();}).then(function(d){BROWSE_PATH=d.path||'';q('browsePath').textContent=BROWSE_PATH||t('drives');var h='';if(d.parent!==null&&d.parent!==undefined){h+='<div class="browseitem up" onclick="browseTo(&#39;'+jsq(d.parent)+'&#39;)"><span class="bic">&#8617;</span>'+esc(t('upfolder'))+'</div>';}(d.dirs||[]).forEach(function(x){h+='<div class="browseitem" onclick="browseTo(&#39;'+jsq(x.path)+'&#39;)"><span class="bic">&#128193;</span>'+esc(x.name)+'</div>';});if(!h)h='<div class="tempty">'+esc(t('empty'))+'</div>';q('browseList').innerHTML=h.replace(/&#39;/g,String.fromCharCode(39));}).catch(function(){q('browseList').innerHTML='<div class="tempty">'+esc(t('invalidpath'))+'</div>';});}
function toggleLegend(){q('legend').classList.toggle('on');}
function legTab(t){
  q('legtab-leg').classList.toggle('on',t==='leg');
  q('legtab-aide').classList.toggle('on',t==='aide');
  q('legpane-leg').classList.toggle('on',t==='leg');
  q('legpane-aide').classList.toggle('on',t==='aide');
}

function loadQuick(){
  fetch('/api/quick?token='+TOKEN).then(function(r){return r.json();}).then(function(d){
    if(d.users&&d.users.length){
      var h='';d.users.forEach(function(u){var up=np(u.path);h+='<button class="drive" title="'+esc(up)+'" onclick="q(&#39;path&#39;).value=&#39;'+jsq(up)+'&#39;;onPathInput()"><b>'+esc(u.name)+'</b></button>';});
      q('users').innerHTML=h.replace(/&#39;/g,String.fromCharCode(39));q('usersWrap').style.display='';
    }else{q('usersWrap').style.display='none';}
    var seen={},hh='',n=0;
    if(d.history){d.history.forEach(function(p){var pp=np(p),k=pp.toLowerCase();if(seen[k])return;seen[k]=1;n++;hh+='<span class="rec"><button class="drive" title="'+esc(pp)+'" onclick="q(&#39;path&#39;).value=&#39;'+jsq(pp)+'&#39;;onPathInput()">'+esc(pp)+'</button><button class="recx" title="'+esc(t('removeone'))+'" onclick="removeRecent(&#39;'+jsq(pp)+'&#39;)">'+String.fromCharCode(215)+'</button></span>';});}
    if(n){hh+='<button class="clearhist" onclick="clearHistory()">'+esc(t('clearhist'))+'</button>';q('history').innerHTML=hh.replace(/&#39;/g,String.fromCharCode(39));q('histWrap').style.display='';}
    else{q('histWrap').style.display='none';}
  }).catch(function(){});
}
function removeRecent(p){fetch('/api/history?token='+TOKEN+'&remove='+encodeURIComponent(p)).then(function(){loadQuick();}).catch(function(){});}
function clearHistory(){fetch('/api/history?token='+TOKEN+'&clear=1').then(function(){loadQuick();}).catch(function(){});}
function exclParam(){var el=q('exclcustom');if(!el)return '';var v=el.value.split(/[\r\n;]+/).map(function(x){return x.trim();}).filter(Boolean);return v.length?('&exclude='+encodeURIComponent(v.join(';'))):'';}
function toggleExcl(){var b=q('exclbody');if(!b)return;var open=b.style.display==='none';b.style.display=open?'':'none';q('exclchev').textContent=open?String.fromCharCode(9652):String.fromCharCode(9662);if(open&&!b.dataset.loaded){loadExclusions();b.dataset.loaded='1';}}
function loadExclusions(){fetch('/api/exclusions?token='+TOKEN).then(function(r){return r.json();}).then(function(a){var h='';(a||[]).forEach(function(x){h+='<li><code>'+esc(np(x))+'</code></li>';});q('excllist').innerHTML=h;}).catch(function(){});}

function pickFolder(){
  var b=q('pickbtn');b.disabled=true;var old=b.textContent;b.textContent=t('windowopen');
  fetch('/api/pick?token='+TOKEN+'&path='+encodeURIComponent(q('path').value.trim())).then(function(r){return r.json();}).then(function(d){
    if(d&&d.path){q('path').value=d.path;onPathInput();}
    else if(d&&d.err){showErr(t('pickerfail'));}
    b.disabled=false;b.textContent=old;
  }).catch(function(){b.disabled=false;b.textContent=old;});
}

function onDepth(){
  var unl=q('depthUnl').checked;
  q('depth').disabled=unl;
  var v=unl?0:parseInt(q('depth').value,10);
  q('depthval').textContent=unl?'\u221E':v;
  var msg;
  if(unl)msg=t('depthunl');
  else if(v<=2)msg=t('depthlow');
  else if(v<=5)msg=t('depthmid');
  else msg=t('depthhigh');
  q('depthExp').textContent=msg+' '+t('depthtail');
}

function onMinSize(){DISPLAY_MIN=parseInt(q('minsize').value,10)||0;if(CUR)renderTree();}

function loadDrives(){
  fetch('/api/drives?token='+TOKEN).then(function(r){return r.json();}).then(function(list){
    var html='';
    list.forEach(function(d){
      if(!d.ready){return;}
      var used=(d.totalGB>0)?Math.max(0,Math.min(100,Math.round((d.totalGB-d.freeGB)/d.totalGB*100))):0;var bc=used>=90?'bad':(used>=75?'warn':'ok');html+='<button class="drive" onclick="q(\'path\').value=\''+d.name.replace(/\\/g,'\\\\')+'\';onPathInput()">'+'<b>'+d.name+'</b> '+(d.label||'')+'<small>'+d.freeGB+' / '+d.totalGB+' GB '+t('drivefree')+'</small>'+'<span class="dbar '+bc+'"><i style="width:'+used+'%"></i></span></button>';
    });
    q('drives').innerHTML=html;
  }).catch(function(){});
}

var scanTimer=null, scanStart=0;
function fmtElapsed(s){if(s<60)return s+'s';var m=Math.floor(s/60);return m+'m'+(s%60)+'s';}
function stopTimer(){if(scanTimer){clearInterval(scanTimer);scanTimer=null;}}

function startScan(){
  var path=q('path').value.trim();
  if(!path){showErr(t('enterpath'));return;}
  q('err').classList.remove('on');
  fetch('/api/check?token='+TOKEN+'&path='+encodeURIComponent(path)).then(function(r){return r.json();}).then(function(c){
    if(c&&c.ok===false){showErr(t('invalidpath'));setPathValid(false);return;}
    runScan(path);
  }).catch(function(){runScan(path);});
}
function runScan(path){
  saveSettings(path);
  if(es){try{es.close()}catch(e){}}if(ses){try{ses.close()}catch(e){}}subQueue=[];HOMEVIEW=null;
  // Ferme la modale de reglages, l'app affiche l'arbre
  q('overlay').classList.remove('on');
  q('dpct').textContent='0%';
  q('gbarfill').style.width='0%';
  q('gbar').classList.add('loading');
  q('dcur').textContent=t('launching');
  q('dopwrap').textContent='';
  q('derr').textContent='';
  NODES={};ROOT=null;CUR=null;SCANNING=null;ACTIVE_MAP={};BUSY=true;SUBSCAN=null;
  q('tlist').innerHTML='';q('tbc').innerHTML='';q('ttotal').innerHTML='';q('dscan').innerHTML='';
  scanStart=Date.now();
  stopTimer();
  scanTimer=setInterval(function(){q('delapsed').textContent=fmtElapsed(Math.floor((Date.now()-scanStart)/1000));},1000);

  var depth=q('depthUnl').checked?0:q('depth').value;
  try{q('stdepth').textContent=q('depthUnl').checked?t('scandepthfull'):t('scandepthn').replace('{n}',q('depth').value);}catch(e){}
  var url='/api/scan?token='+TOKEN+'&path='+encodeURIComponent(path)+'&depth='+encodeURIComponent(depth)+'&mode='+selMode+exclParam();
  es=new EventSource(url);
  es.onmessage=function(ev){
    try{var d=JSON.parse(ev.data);
      if(typeof d.pct==='number'){var p=Math.max(0,Math.min(100,d.pct));q('gbarfill').style.width=p+'%';q('dpct').textContent=p+'%';}
      if(d.status){q('dcur').textContent=d.status;}
      q('dopwrap').textContent=d.op?(' \u00b7 '+d.op):'';
    }catch(e){}
  };
  es.addEventListener('root',onTreeRoot);
  es.addEventListener('node',onTreeNode);
  es.addEventListener('special',onTreeSpecial);
  es.addEventListener('active',onTreeActive);
  es.addEventListener('totals',onTreeTotals);
  es.addEventListener('scanning',onTreeScanning);
  es.addEventListener('done',function(ev){
    for(var k in NODES){var n=NODES[k];if(n.kind==='dir'&&n.state!=='done'){if(n.size===null)n.size=0;n.state='done';}}
    SCANNING=null;ACTIVE_MAP={};q('dscan').innerHTML='';
    q('gbar').classList.remove('loading');
    q('gbarfill').style.width='100%';q('dpct').textContent='100%';
    q('dcur').textContent=t('done');
    stopTimer();es.close();BUSY=false;
    q('dback').classList.add('on');
    renderTree();
    if(CUR)loadFiles(CUR);
    processQueue();
  });
  es.addEventListener('failed',function(ev){
    es.close();stopTimer();
    var msg='Echec du scan.';try{msg=JSON.parse(ev.data).message||msg;}catch(e){}
    showDisplayErr(msg);
  });
  es.onerror=function(){
    if(es.readyState===2){stopTimer();showDisplayErr('Connexion au serveur interrompue.');}
  };
}

function showDisplayErr(m){BUSY=false;SCANNING=null;ACTIVE_MAP={};q('dscan').innerHTML='';try{q('gbar').classList.remove('loading');}catch(e){}q('derr').textContent=m;q('derr').classList.add('on');q('dcur').textContent=t('interrupted');q('dback').classList.add('on');hideStop();}
function showErr(m){var e=q('err');e.textContent=m;e.classList.add('on');}

// ===== Arbre navigable live (v5.0) =====
var NODES={},ROOT=null,CUR=null,SCANNING=null,FILES={},ACTIVE_MAP={};

function tfmt(b){
  if(b===null||b===undefined)return '...';
  if(b<0)return '';
  var u=['o','Ko','Mo','Go','To','Po'],i=0,n=b;
  while(n>=1024&&i<u.length-1){n/=1024;i++;}
  var v=(i===0)?String(n):n.toLocaleString(LANG,{minimumFractionDigits:2,maximumFractionDigits:2});
  return v+' '+u[i];
}
function esc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}
function jsq(s){return s.replace(/\\/g,'\\\\').replace(/'/g,"\\'");}
function pj(ev){try{return JSON.parse(ev.data);}catch(e){return null;}}
function isAncestorOf(anc,path){if(!anc||!path)return false;var a=anc.replace(/\\+$/,'');return path===a||path.indexOf(a+'\\')===0;}
function np(p){if(typeof p!=='string')return p;var unc=p.slice(0,2)==='\\\\';var q=p.replace(/\\+/g,'\\');return unc?'\\'+q:q;}
function dotIcon(c){return '<span class="ic">&#128193;<span class="sdot sdot-'+c+'"></span></span>';}
var EXTICON={};
(function(){function A(ic,l){l.forEach(function(e){EXTICON[e]=ic;});}
A('&#128444;&#65039;',['jpg','jpeg','png','gif','bmp','webp','svg','ico','tif','tiff','heic','heif','psd','raw','cr2','nef']);
A('&#127916;',['mp4','mkv','avi','mov','wmv','flv','webm','m4v','mpg','mpeg','3gp','ts','vob']);
A('&#127925;',['mp3','wav','flac','aac','ogg','wma','m4a','opus','mid','aiff']);
A('&#128213;',['pdf']);
A('&#128216;',['doc','docx','odt','rtf','pages']);
A('&#128215;',['xls','xlsx','xlsm','csv','ods','tsv']);
A('&#128217;',['ppt','pptx','odp']);
A('&#128221;',['txt','md','markdown','log','nfo','ini','cfg','conf']);
A('&#128476;&#65039;',['zip','rar','7z','tar','gz','bz2','xz','iso','cab','tgz','z','lz']);
A('&#9881;&#65039;',['exe','msi','bat','cmd','com','ps1','psm1','sh','app','apk','deb','rpm']);
A('&#129513;',['dll','sys','drv','so','o','a','lib','ocx']);
A('&#128187;',['js','ts','jsx','tsx','py','java','c','cpp','h','hpp','cs','go','rb','php','html','htm','css','scss','json','xml','yml','yaml','sql','swift','kt','rs','pl','lua','r','vb']);
A('&#128292;',['ttf','otf','woff','woff2','eot','fon']);
A('&#128189;',['vhd','vhdx','vmdk','vdi','img','dmg','ova','ovf']);
A('&#128451;&#65039;',['db','sqlite','sqlite3','mdb','accdb','bak','dat','mdf','ldf']);
A('&#128218;',['epub','mobi','azw','azw3','fb2','djvu']);
A('&#128273;',['pem','crt','cer','pfx','p12','pub','asc','gpg','key']);
A('&#128279;',['lnk','url','desktop']);
A('&#129522;',['torrent']);
})();
function fileIcon(name){var i=name.lastIndexOf('.');if(i<=0||i===name.length-1)return '&#128196;';return EXTICON[name.slice(i+1).toLowerCase()]||'&#128196;';}

function ensureNode(path,name,parent){
  if(name)name=np(name);
  var nd=NODES[path];
  if(!nd){nd={name:name||path,parent:(parent!==undefined?parent:null),size:null,state:'pending',kind:'dir',children:[]};NODES[path]=nd;}
  else{if(name)nd.name=name;if(parent!==undefined&&parent!==null&&nd.parent===null)nd.parent=parent;}
  return nd;
}
function linkChild(parent,path){if(!parent)return;var p=ensureNode(parent);if(p.children.indexOf(path)<0)p.children.push(path);}

function onTreeRoot(ev){var d=pj(ev);if(!d)return;d.path=np(d.path);var nd=ensureNode(d.path,d.name);nd.state='active';
  if(SUBSCAN){
    if(SUBPARENT){if(!nd.parent)nd.parent=SUBPARENT;linkChild(SUBPARENT,d.path);}
    CUR=d.path;
  } else {ROOT=d.path;if(CUR===null){CUR=d.path;}}
  renderTree();
}
function onTreeNode(ev){var d=pj(ev);if(!d)return;d.path=np(d.path);d.parent=np(d.parent);ensureNode(d.path,d.name,d.parent);linkChild(d.parent,d.path);if(CUR===d.parent)renderTree();}
function onTreeSpecial(ev){var d=pj(ev);if(!d)return;d.path=np(d.path);d.parent=np(d.parent);var nd=ensureNode(d.path,d.name,d.parent);nd.kind=d.kind;nd.state='done';nd.size=-1;linkChild(d.parent,d.path);if(CUR===d.parent)renderTree();}
function onTreeActive(ev){var d=pj(ev);if(!d)return;d.path=np(d.path);var nd=ensureNode(d.path);nd.state='active';if(CUR===d.path||CUR===nd.parent)renderTree();}
function onTreeTotals(ev){var d=pj(ev);if(!d||!d.items)return;var upd=false;d.items.forEach(function(it){var p=np(it.path);var nd=ensureNode(p);nd.size=it.size;if(typeof it.count==='number')nd.count=it.count;if(p!==ROOT)nd.state='done';if(p===CUR||nd.parent===CUR)upd=true;});if(upd)renderTree();}
function cfmt(n){try{return Number(n).toLocaleString(LANG);}catch(e){return ''+n;}}
function descDirs(p,memo){if(memo[p]!=null)return memo[p];var nd=NODES[p];if(!nd||!nd.children)return memo[p]=0;var c=0;for(var i=0;i<nd.children.length;i++){var cn=NODES[nd.children[i]];if(cn&&cn.kind==='dir'){c+=1+descDirs(nd.children[i],memo);}}return memo[p]=c;}
function counts(nd,cp,memo){var f=(typeof nd.count==='number'&&nd.count>=0)?nd.count:null;var s=descDirs(cp,memo);var parts=[];if(s>0)parts.push('&#128193; '+cfmt(s));if(f!=null)parts.push('&#128196; '+cfmt(f));return parts.join(' &middot; ');}
function onTreeScanning(ev){var d=pj(ev);if(!d)return;var p=np(d.path);SCANNING=p;
  if(d.anc){ACTIVE_MAP={};d.anc.forEach(function(a){var ap=np(a.path);ACTIVE_MAP[ap]={size:a.size,count:a.count};var an=NODES[ap];if(an&&an.state!=='done'){if(typeof a.size==='number'&&(an.size==null||a.size>an.size))an.size=a.size;if(typeof a.count==='number'&&(an.count==null||a.count>an.count))an.count=a.count;}});}
  var segs=p.split('\\'),clean=[];for(var i=0;i<segs.length;i++){if(segs[i]!=='')clean.push(segs[i]);}
  var leaf=clean.length?clean[clean.length-1]:p;
  var parents=clean.slice(0,clean.length-1),prefix='',MAXSEG=4;
  if(parents.length>MAXSEG){prefix='\u2026\\';parents=parents.slice(parents.length-MAXSEG);}
  var phtml='';for(var j=0;j<parents.length;j++){phtml+=esc(parents[j])+'<span class="segsep">\\</span>';}
  q('dscan').innerHTML='<span class="spin"></span><span class="lbl">'+t('inprogress')+'</span>'
    +'<span class="scanparents" title="'+esc(p)+'">'+prefix+phtml+'</span>'
    +'<span class="segleaf" title="'+esc(p)+'">'+esc(leaf)+'</span>';
  renderTree();
}

function goTo(p){if(!NODES[p])return;CUR=p;renderTree();loadFiles(p);}

function loadFiles(p){
  if(BUSY)return;                      // serveur occupe : fichiers apres le scan
  if(FILES[p]!==undefined)return;      // deja charge ou en cours
  var nd=NODES[p];
  if(nd&&nd.kind!=='dir')return;
  FILES[p]={loading:true,list:[]};
  if(CUR===p)renderTree();
  fetch('/api/files?token='+TOKEN+'&path='+encodeURIComponent(p)).then(function(r){return r.json();}).then(function(d){
    FILES[p]={loading:false,list:(d&&d.files)?d.files:[],capped:!!(d&&d.capped)};
    if(CUR===p)renderTree();
  }).catch(function(){FILES[p]={loading:false,list:[]};if(CUR===p)renderTree();});
}

var ses=null,BUSY=false,SUBSCAN=null,SUBPARENT=null,subQueue=[],HOMEVIEW=null;

function scanAllGrey(){
  if(!CUR||!NODES[CUR])return;
  var cur=NODES[CUR],greys=[];
  (cur.children||[]).forEach(function(cp){var nd=NODES[cp];if(nd&&(nd.kind==='unscanned'||nd.kind==='denied'||nd.kind==='excluded'||nd.kind==='junction')){greys.push(cp);}});
  if(!greys.length)return;
  HOMEVIEW=CUR;
  greys.forEach(function(cp){var nd=NODES[cp];if(nd&&!nd.queued){nd.kind='dir';nd.queued=true;nd.size=null;if(subQueue.indexOf(cp)<0)subQueue.push(cp);}});
  renderTree();
  if(!BUSY)processQueue();
}

function subScan(path){
  var nd=NODES[path];if(!nd)return;
  if(nd.queued)return;
  if(BUSY){
    // Un scan tourne : on met en file et on marque le dossier "en file" (bleu)
    if(subQueue.indexOf(path)<0)subQueue.push(path);
    nd.kind='dir';nd.queued=true;nd.size=null;
    q('dcur').textContent=t('queuedmsg').replace('{n}',subQueue.length);
    renderTree();
    return;
  }
  HOMEVIEW=null;
  runSubScan(path);
}

function processQueue(){
  if(BUSY)return;
  while(subQueue.length){
    var p=subQueue.shift();
    if(NODES[p]){runSubScan(p);return;}
  }
  HOMEVIEW=null;
}

function runSubScan(path){
  var nd=NODES[path];if(!nd)return;
  BUSY=true;SUBSCAN=path;SUBPARENT=(nd.parent||null);SCANNING=null;ACTIVE_MAP={};
  nd.kind='dir';nd.state='active';nd.queued=false;nd.size=null;
  CUR=HOMEVIEW?HOMEVIEW:path;
  q('dscan').innerHTML='';q('gbar').classList.add('loading');
  q('dcur').textContent=t('launching');
  renderTree();
  var url='/api/scan?token='+TOKEN+'&path='+encodeURIComponent(path)+'&depth=2&mode='+(selMode||1)+'&sub=1';
  ses=new EventSource(url);
  ses.onmessage=function(ev){try{var d=JSON.parse(ev.data);if(d.status)q('dcur').textContent=d.status;}catch(e){}};
  ses.addEventListener('root',onTreeRoot);
  ses.addEventListener('node',onTreeNode);
  ses.addEventListener('special',onTreeSpecial);
  ses.addEventListener('active',onTreeActive);
  ses.addEventListener('totals',onTreeTotals);
  ses.addEventListener('scanning',onTreeScanning);
  ses.addEventListener('done',function(){
    ses.close();BUSY=false;SUBSCAN=null;SUBPARENT=null;SCANNING=null;
    q('dscan').innerHTML='';q('gbar').classList.remove('loading');
    for(var k in NODES){if((k===path||isAncestorOf(path,k))&&NODES[k].kind==='dir'&&NODES[k].state!=='done'){if(NODES[k].size===null)NODES[k].size=0;NODES[k].state='done';}}
    q('dcur').textContent=t('done');
    renderTree();
    if(CUR)loadFiles(CUR);
    processQueue();
  });
  ses.addEventListener('failed',function(ev){
    ses.close();BUSY=false;SUBSCAN=null;SUBPARENT=null;q('gbar').classList.remove('loading');
    var msg='Echec.';try{msg=JSON.parse(ev.data).message||msg;}catch(e){}
    q('dcur').textContent='Echec du scan : '+msg;renderTree();processQueue();
  });
  ses.onerror=function(){if(ses.readyState===2){BUSY=false;SUBSCAN=null;SUBPARENT=null;q('gbar').classList.remove('loading');renderTree();processQueue();}};
}

var sortMode='size',_szKeys=null,_szTs=0,_szView=null;
function effSort(){return sortMode;}
function itKey(it){return it.dir?('d:'+it.path):('f:'+(it.name||''));}
function itName(it){return((it.dir?it.nd.name:it.name)||'').toLowerCase();}
function sortItems(items){
  var mode=effSort();
  if(mode!=='size'){items.sort(function(a,b){var an=itName(a),bn=itName(b);return an<bn?-1:(an>bn?1:0);});return;}
  function bySize(a,b){var sa=(typeof a.size==='number'&&a.size>=0)?a.size:0,sb=(typeof b.size==='number'&&b.size>=0)?b.size:0;if(sb!==sa)return sb-sa;var an=itName(a),bn=itName(b);return an<bn?-1:(an>bn?1:0);}
  if(!BUSY){items.sort(bySize);return;}
  var now=Date.now();
  if(_szView!==CUR||!_szKeys||now-_szTs>1000){items.sort(bySize);_szKeys={};items.forEach(function(it,i){_szKeys[itKey(it)]=i;});_szTs=now;_szView=CUR;}
  else{items.sort(function(a,b){var ia=_szKeys[itKey(a)],ib=_szKeys[itKey(b)];if(ia==null&&ib==null)return bySize(a,b);if(ia==null)return 1;if(ib==null)return -1;return ia-ib;});}
}
function updateSortBtn(){var b=q('sortbtn');if(b)b.textContent=t('sortby')+' : '+(effSort()==='size'?t('bysize'):t('byname'));}
function toggleSort(){sortMode=(sortMode==='size')?'name':'size';try{localStorage.setItem('psncdu_sort',sortMode);}catch(e){}_szKeys=null;updateSortBtn();renderTree();}
function renderTree(){
  if(!CUR||!NODES[CUR])return;
  var cur=NODES[CUR];
  var dmemo={};
  var chain=[],p=CUR;
  while(p){chain.unshift(p);p=NODES[p]?NODES[p].parent:null;}
  var bc='';
  chain.forEach(function(pp,i){
    var nm=NODES[pp]?NODES[pp].name:pp;
    bc+=(i>0?'<span class="tsep">&rsaquo;</span>':'')+'<a class="tcrumb" onclick="goTo(\''+jsq(pp)+'\')">'+esc(nm||pp)+'</a>';
  });
  q('tbc').innerHTML=bc;
  // Fusion dossiers + fichiers, tries par taille (le plus gros en haut)
  var items=[];
  (cur.children||[]).forEach(function(cp){var nd=NODES[cp];if(nd)items.push({dir:true,path:cp,nd:nd,size:(nd.size!=null&&nd.size>=0)?nd.size:(ACTIVE_MAP[cp]?ACTIVE_MAP[cp].size:-1)});});
  var fe=FILES[CUR];
  if(fe&&fe.list){fe.list.forEach(function(f){items.push({dir:false,name:f.name,size:f.size});});}
  // Tri : "nom" (stable, remplissage de haut en bas) ou "taille" (gros en
  // haut). Par defaut nom pendant le scan et taille a la fin ; le bouton de
  // tri force l'un ou l'autre. En mode taille pendant le scan, re-tri lisse.
  sortItems(items);
  updateSortBtn();
  // Total et compteurs calcules depuis le contenu du dossier (robuste : la
  // taille remontee du serveur pour le dossier courant peut etre partielle).
  var childSum=0,childFiles=0;
  items.forEach(function(it){if(typeof it.size==='number'&&it.size>=0)childSum+=it.size;});
  (cur.children||[]).forEach(function(cp){var cn=NODES[cp];if(cn&&typeof cn.count==='number'&&cn.count>=0)childFiles+=cn.count;});
  var ownFiles=(fe&&fe.list)?fe.list.length:0;
  var totFiles=childFiles+ownFiles;
  var totSize=Math.max((typeof cur.size==='number'&&cur.size>=0)?cur.size:0,childSum);
  var denom=totSize>0?totSize:1;
  var hcnt=[];var ds=descDirs(CUR,dmemo);if(ds>0)hcnt.push('&#128193; '+cfmt(ds));if(totFiles>0)hcnt.push('&#128196; '+cfmt(totFiles));
  q('ttotal').innerHTML=t('total')+' <strong>'+tfmt(totSize)+'</strong>'+(hcnt.length?' <span class="tcnt">'+hcnt.join(' &middot; ')+'</span>':'')+((cur.state==='active'||isAncestorOf(CUR,SCANNING))?' <span class="tactive-tag">'+t('scanningtag')+'</span>':'');
  var greyCount=0;
  (cur.children||[]).forEach(function(cp){var gn=NODES[cp];if(gn&&(gn.kind==='unscanned'||gn.kind==='denied'||gn.kind==='excluded'||gn.kind==='junction'))greyCount++;});
  var gb=q('scangrey');if(gb){if(greyCount>0){gb.style.display='';gb.textContent=t('scangrey')+' ('+greyCount+')';}else{gb.style.display='none';}}
  var html='';
  if(cur.parent){html+='<div class="trow tup" onclick="goTo(\''+jsq(cur.parent)+'\')"><span class="ticon">&#8617;</span><span class="tname">'+t('parent')+'</span><span class="tsize"></span></div>';}
  var hidden=0;
  items.forEach(function(it){
    if(it.dir){
      var nd=it.nd,cp=it.path;
      if(DISPLAY_MIN>0&&nd.kind==='dir'&&nd.state==='done'&&nd.size>=0&&nd.size<DISPLAY_MIN){hidden++;return;}
      var greyKind=(nd.kind==='excluded'||nd.kind==='junction'||nd.kind==='unscanned'||nd.kind==='denied');
      var scanning=(SCANNING&&nd.kind==='dir'&&nd.state!=='done'&&isAncestorOf(cp,SCANNING));
      var icon,cls;
      if(greyKind){icon=dotIcon('grey');cls='skip';}
      else if(scanning){icon='<span class="spin"></span>';cls='active';}
      else if(nd.state==='active'){icon='<span class="spin"></span>';cls='active';}
      else if(nd.queued){icon=dotIcon('blue');cls='pending';}
      else if(nd.state==='pending'){icon=dotIcon('orange');cls='pending';}
      else{icon=dotIcon('green');cls='done';}
      var szt;
      if(nd.kind==='excluded')szt=t('excluded');else if(nd.kind==='junction')szt=t('junction');else if(nd.kind==='unscanned')szt=t('unscanned');else if(nd.kind==='denied')szt=t('protected');else if(nd.size===-1)szt=t('error');else if(scanning&&ACTIVE_MAP[cp])szt=tfmt(ACTIVE_MAP[cp].size);else if(scanning)szt=t('scanshort');else if(nd.queued)szt=t('queued');else if(nd.state==='pending')szt=(nd.size!=null&&nd.size>=0)?tfmt(nd.size):'...';else szt=tfmt(nd.size);
      var cnt=(greyKind)?'':counts(nd,cp,dmemo);
      if(scanning&&ACTIVE_MAP[cp]){var ds2=descDirs(cp,dmemo);var cp2=[];if(ds2>0)cp2.push('&#128193; '+cfmt(ds2));if(ACTIVE_MAP[cp].count!=null)cp2.push('&#128196; '+cfmt(ACTIVE_MAP[cp].count));cnt=cp2.join(' &middot; ');}
      var pct=(nd.size!=null&&nd.size>=0)?Math.min(100,Math.round(nd.size/denom*100)):null;
      var clickFn=greyKind?'subScan':(nd.kind==='dir'?'goTo':null);
      var titleAttr=greyKind?' title="'+esc(t('clickscan'))+'"':'';
      html+='<div class="trow '+cls+'"'+titleAttr+(clickFn?' onclick="'+clickFn+'(\''+jsq(cp)+'\')"':'')+'>'
        +'<span class="ticon">'+icon+'</span><span class="tname">'+esc(nd.name)+'</span>'
        +'<span class="tbarwrap"><span class="tbar" style="width:'+(pct!=null?pct:0)+'%"></span></span>'
        +'<span class="tpct">'+(pct!=null?pct+'%':'')+'</span><span class="tcount">'+cnt+'</span><span class="tsize">'+szt+'</span></div>';
    } else {
      if(DISPLAY_MIN>0&&it.size>=0&&it.size<DISPLAY_MIN){hidden++;return;}
      var fpct=(it.size>=0)?Math.min(100,Math.round(it.size/denom*100)):null;
      html+='<div class="trow file" title="'+esc(it.name)+'">'
        +'<span class="ticon"><span class="fic">'+fileIcon(it.name)+'</span></span><span class="tname">'+esc(it.name)+'</span>'
        +'<span class="tbarwrap"><span class="tbar tbar-file" style="width:'+(fpct!=null?fpct:0)+'%"></span></span>'
        +'<span class="tpct">'+(fpct!=null?fpct+'%':'')+'</span><span class="tcount"></span><span class="tsize">'+tfmt(it.size)+'</span></div>';
    }
  });
  if(fe&&fe.loading){html+='<div class="tempty">'+t('loadingfiles')+'</div>';}
  else if(!items.length){html+='<div class="tempty">'+((cur.state==='active'||isAncestorOf(CUR,SCANNING))?t('scandots'):(BUSY?t('filesafter'):t('empty')))+'</div>';}
  if(fe&&fe.capped){html+='<div class="tempty">'+t('filescap')+'</div>';}
  if(hidden>0){html+='<div class="tempty">'+t('hiddenmsg').replace('{n}',hidden)+'</div>';}
  else if(BUSY&&items.length&&!fe&&cur.state==='done'){html+='<div class="tempty">'+t('filesafter')+'</div>';}
  q('tlist').innerHTML=html;
  try{q('stopscan').style.display=BUSY?'':'none';}catch(e){}
}

function quitServer(){
  fetch('/api/quit?token='+TOKEN).then(function(){
    document.body.innerHTML='<div style="padding:60px;text-align:center;font-family:sans-serif;color:#4a6080">'+t('serverstopped')+'</div>';
  }).catch(function(){
    document.body.innerHTML='<div style="padding:60px;text-align:center;font-family:sans-serif;color:#4a6080">'+t('serverstopped')+'</div>';
  });
}

function getParam(n){var m=new RegExp('[?&]'+n+'=([^&]*)').exec(window.location.search);return m?decodeURIComponent(m[1].replace(/\+/g,' ')):null;}

try{var _sl=localStorage.getItem('psncdu_lang');if(_sl&&I18N[_sl])LANG=_sl;}catch(e){}
document.documentElement.lang=LANG;
document.documentElement.dir=/^(ar|ur|fa|he)$/.test(LANG)?'rtl':'ltr';
applyI18n();
try{if(q('langsel'))q('langsel').value=LANG;}catch(e){}
try{var _sm=localStorage.getItem('psncdu_sort');if(_sm==='name'||_sm==='size')sortMode=_sm;}catch(e){}
try{updateSortBtn();}catch(e){}
loadDrives();loadQuick();restoreSettings();onDepth();

// Auto-lancement quand on arrive depuis un bouton "Scanner ce dossier"
(function(){
  var sc=getParam('scan');
  if(sc){
    q('path').value=sc;
    var d=getParam('depth');
    if(d!==null&&d!==''){q('depth').value=d;}
    setTimeout(startScan,200);
  }
})();
</script>
</body>
</html>
'@

# --- Langue du systeme (repli anglais) ----------------------
$SUPPORTED_LANGS = @('en','zh','hi','es','fr','ar','bn','pt','ru','ur','id','de')
$sysLang = 'en'
try { $sysLang = ([System.Globalization.CultureInfo]::CurrentUICulture.TwoLetterISOLanguageName).ToLower() } catch {}
if ($SUPPORTED_LANGS -notcontains $sysLang) { $sysLang = 'en' }
Write-Log "[WEB] Langue systeme : $([System.Globalization.CultureInfo]::CurrentUICulture.Name) -> UI en '$sysLang'"

$SETTINGS_HTML = $SETTINGS_HTML.
    Replace('__VERSION__',     $SCRIPT_VERSION).
    Replace('__EMAIL__',       $USER_EMAIL).
    Replace('__MODEL__',       "Claude Opus 4.8").
    Replace('__DEFAULTPATH__', $DEFAULT_PATH).
    Replace('__DEFAULTDEPTH__',[string]$DEFAULT_DEPTH).
    Replace('__MODES_JSON__',  $MODES_JSON).
    Replace('__LOGO__',        $APP_LOGO_SVG).
    Replace('__LANG__',        $sysLang).
    Replace('__TOKEN__',       $script:Token)

# --- Handler du scan (SSE) ----------------------------------
function Invoke-ScanRequest {
    param($Req, $Resp)

    if ($script:ScanBusy) { Send-Status $Resp 409 "scan en cours"; return }
    $script:ScanBusy = $true
    $script:CancelScan = $false

    $pathIn  = $Req.QueryString["path"]
    $depthIn = $Req.QueryString["depth"]
    $modeIn  = $Req.QueryString["mode"]
    $exIn    = $Req.QueryString["exclude"]

    $maxDepth = $DEFAULT_DEPTH
    try { $maxDepth = [int]$depthIn } catch {}
    if ($maxDepth -lt 0 -or $maxDepth -gt 10) { $maxDepth = $DEFAULT_DEPTH }

    $modeKey = 1
    try { $mk = [int]$modeIn; if ($SCAN_MODES.ContainsKey($mk)) { $modeKey = $mk } } catch {}
    $selectedMode = $SCAN_MODES[$modeKey]

    # Prepare le flux SSE
    $Resp.StatusCode  = 200
    $Resp.ContentType = "text/event-stream"
    try { $Resp.Headers.Add("Cache-Control","no-cache") } catch {}
    try { $Resp.SendChunked = $true } catch {}
    $enc = New-Object System.Text.UTF8Encoding($false)
    $sw  = New-Object System.IO.StreamWriter($Resp.OutputStream, $enc)
    $sw.AutoFlush = $true
    # Script-scoped : indispensable pour que le ProgressSink, invoque
    # depuis Update-Progress (autre portee), retrouve bien le writer.
    $script:SseWriter = $sw

    $sendEvent = {
        param([string]$EventName, [string]$Json)
        try {
            if ($EventName) { $sw.Write("event: $EventName`n") }
            $sw.Write("data: $Json`n`n")
        } catch {}
    }

    try {
        $startPath = Normalize-Path $pathIn
        if ($req.QueryString["sub"] -ne "1") { Add-History $startPath }
        if ([string]::IsNullOrWhiteSpace($startPath) -or -not (Test-Path -LiteralPath $startPath -ErrorAction SilentlyContinue)) {
            & $sendEvent 'failed' "{""message"":""Chemin invalide ou inaccessible : $(ConvertTo-JsonSafe $pathIn)""}"
            return
        }

        Write-Log "[WEB] Scan demande : '$startPath' depth=$maxDepth mode=$modeKey"

        # Le sink pousse chaque Update-Progress vers SSE.
        # Il ecrit via $script:SseWriter (portee script) car il est
        # invoque depuis Update-Progress, hors de cette fonction.
        $script:ProgressSink = {
            param($act,$status,$op,$pct)
            try {
                $j = "{""pct"":$pct,""status"":""$(ConvertTo-JsonSafe $status)"",""op"":""$(ConvertTo-JsonSafe $op)""}"
                $script:SseWriter.Write("data: $j`n`n")
            } catch { $script:CancelScan = $true }
        }
        # v5.0 : evenements d'arbre (structure + tailles) vers SSE
        $script:TreeSink = {
            param($ev,$json)
            try { $script:SseWriter.Write("event: $ev`ndata: $json`n`n") } catch { $script:CancelScan = $true }
        }

        $savedExcluded = $EXCLUDED_DIRS
        $userEx = @()
        if ($exIn) { foreach ($e in ($exIn -split ';')) { $e = $e.Trim(); if ($e) { $userEx += (Normalize-Path $e) } } }
        if ($userEx.Count -gt 0) { $script:EXCLUDED_DIRS = @($EXCLUDED_DIRS + $userEx); Write-Log "[WEB] Exclusions personnalisees : $($userEx -join ', ')" }

        $result = Start-FastScan -RootPath $startPath -MaxDepth $maxDepth -Mode $selectedMode

        # v5.4 : interface unique. L'arbre live EST la vue finale, plus de
        # rapport separe a generer. On signale juste la fin.
        if ($script:CancelScan) {
            Write-Log "[WEB] Scan interrompu (client deconnecte) : $startPath"
        } else {
            & $sendEvent 'done' "{}"
            Write-Log "[WEB] Scan termine : $($result['DirCount']) dossiers en $($result['ElapsedSec'])s"
        }
    }
    catch {
        Write-Log "[WEB] Erreur scan : $_" -Level ERROR
        & $sendEvent 'failed' "{""message"":""$(ConvertTo-JsonSafe ([string]$_))""}"
    }
    finally {
        $script:ProgressSink = $null
        $script:DirEventSink = $null
        $script:TreeSink     = $null
        $script:SseWriter    = $null
        if ($null -ne $savedExcluded) { $script:EXCLUDED_DIRS = $savedExcluded }
        $script:ScanBusy     = $false
        try { $sw.Close() } catch {}
        try { $Resp.OutputStream.Close() } catch {}
    }
}

# --- Demarrage du serveur -----------------------------------
Clear-Host
Write-Host ""
Write-Host "  +-----------------------------------------+" -ForegroundColor Cyan
Write-Host "  |  PS-NCDU  -  Disk Usage Analyzer         |" -ForegroundColor DarkYellow
Write-Host "  |  Application web locale        v$SCRIPT_VERSION      |" -ForegroundColor Cyan
Write-Host "  |  Eric Guiffault - CL SASU                |" -ForegroundColor DarkGray
Write-Host "  +-----------------------------------------+" -ForegroundColor Cyan
Write-Host ""

$listener = New-Object System.Net.HttpListener
$port     = $SERVER_PORT
$bound    = $false
for ($try = 0; $try -lt 40; $try++) {
    $tryPort = $SERVER_PORT + $try
    $listener = New-Object System.Net.HttpListener
    $listener.Prefixes.Add("http://127.0.0.1:$tryPort/")
    try {
        $listener.Start()
        $port  = $tryPort
        $bound = $true
        break
    } catch {
        Write-Log "[WEB] Port $tryPort indisponible : $_" -Level DEBUG
        try { $listener.Close() } catch {}
    }
}

if (-not $bound) {
    Write-Host "  ERREUR : impossible de démarrer le serveur HTTP." -ForegroundColor Red
    Write-Host "  Cause probable : droits insuffisants (HttpListener)." -ForegroundColor Yellow
    Write-Host "  Solutions :" -ForegroundColor Yellow
    Write-Host "    1) Lancer PowerShell en tant qu'administrateur, ou" -ForegroundColor Yellow
    Write-Host "    2) Reserver l'URL une fois (admin) :" -ForegroundColor Yellow
    Write-Host "       netsh http add urlacl url=http://127.0.0.1:$SERVER_PORT/ user=$env:USERDOMAIN\$env:USERNAME" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  Appuyez sur ENTER pour quitter..." -ForegroundColor DarkGray
    $null = Read-Host
    exit 1
}

$rootUrl = "http://127.0.0.1:$port/?token=$($script:Token)"
Write-Host "  Serveur prêt : " -NoNewline -ForegroundColor Green
Write-Host $rootUrl -ForegroundColor White
Write-Host "  (Ouvrez cette URL si le navigateur ne s'ouvre pas seul)" -ForegroundColor DarkGray
Write-Host "  Ctrl+C dans cette console, ou bouton Quitter dans l'UI, pour arrêter." -ForegroundColor DarkGray
Write-Host ""
Write-Log "[WEB] Serveur demarre sur $rootUrl"

try { Start-Process $rootUrl | Out-Null } catch { Write-Log "[WEB] Start-Process navigateur : $_" -Level WARN }

# --- Boucle principale --------------------------------------
$running = $true
while ($running -and $listener.IsListening) {
    try {
        $ctx  = $listener.GetContext()
    } catch { break }
    $req  = $ctx.Request
    $resp = $ctx.Response
    $path = $req.Url.AbsolutePath
    $tok  = $req.QueryString["token"]

    # Toute route (sauf 404) exige le bon jeton
    if ($path -ne '/favicon.ico' -and $tok -ne $script:Token) {
        Send-Status $resp 403 "jeton invalide"
        continue
    }

    switch -Regex ($path) {
        '^/$'            { Send-Text   $resp $SETTINGS_HTML }
        '^/api/drives$'  { Send-Text   $resp (Get-DrivesJson) "application/json; charset=utf-8" }
        '^/api/check$'   {
            $cp = Normalize-Path $req.QueryString["path"]
            $okp = $false; $isd = $false
            try {
                if ($cp -and (Test-Path -LiteralPath $cp)) {
                    $okp = $true
                    $isd = (Test-Path -LiteralPath $cp -PathType Container)
                }
            } catch { $okp = $false }
            Send-Text $resp "{""ok"":$(if($okp){'true'}else{'false'}),""dir"":$(if($isd){'true'}else{'false'})}" "application/json; charset=utf-8"
        }
        '^/api/quick$'   { Send-Text   $resp (Get-QuickJson)  "application/json; charset=utf-8" }
        '^/api/history$' {
            if ($req.QueryString["clear"] -eq '1') { Remove-History -All }
            elseif ($req.QueryString["remove"]) { Remove-History -Path (Normalize-Path $req.QueryString["remove"]) }
            Send-Text $resp (Get-QuickJson) "application/json; charset=utf-8"
        }
        '^/api/exclusions$' {
            $ep=@(); foreach($ex in $EXCLUDED_DIRS){ $ep += """$(ConvertTo-JsonSafe $ex)""" }
            Send-Text $resp ("["+($ep -join ",")+"]") "application/json; charset=utf-8"
        }
        '^/api/files$'   {
            $fp = Normalize-Path $req.QueryString["path"]
            Send-Text $resp (Get-FilesJson $fp) "application/json; charset=utf-8"
        }
        '^/api/pick$'    {
            $init = $req.QueryString["path"]
            $pick = Show-FolderPicker -Initial $init
            Send-Text $resp "{""path"":""$(ConvertTo-JsonSafe $pick.Path)"",""err"":""$(ConvertTo-JsonSafe $pick.Err)""}" "application/json; charset=utf-8"
        }
        '^/api/browse$'  {
            $bp = $req.QueryString["path"]
            $sb = New-Object System.Text.StringBuilder
            if ([string]::IsNullOrWhiteSpace($bp)) {
                [void]$sb.Append('{"path":"","parent":null,"dirs":[')
                $first = $true
                foreach ($root in [System.IO.Directory]::GetLogicalDrives()) {
                    if (-not $first) { [void]$sb.Append(',') }; $first = $false
                    [void]$sb.Append('{"name":"' + (ConvertTo-JsonSafe $root) + '","path":"' + (ConvertTo-JsonSafe $root) + '"}')
                }
                [void]$sb.Append(']}')
            } else {
                $bp2 = Normalize-Path $bp
                if ($bp2 -match '^[A-Za-z]:\\?$') { $par = '' }
                else { $par = Get-ParentPath $bp2; if ($null -eq $par) { $par = '' } }
                [void]$sb.Append('{"path":"' + (ConvertTo-JsonSafe $bp2) + '","parent":"' + (ConvertTo-JsonSafe $par) + '","dirs":[')
                try {
                    $sr = Get-SubDirectories -Path $bp2
                    $items = @($sr["Items"] | Sort-Object Name)
                    $first = $true
                    foreach ($it in $items) {
                        if (-not $first) { [void]$sb.Append(',') }; $first = $false
                        [void]$sb.Append('{"name":"' + (ConvertTo-JsonSafe $it.Name) + '","path":"' + (ConvertTo-JsonSafe $it.FullName) + '"}')
                    }
                } catch {}
                [void]$sb.Append(']}')
            }
            Send-Text $resp $sb.ToString() "application/json; charset=utf-8"
        }
        '^/api/scan$'    { Invoke-ScanRequest $req $resp }
        '^/result$'      {
            if ($null -ne $script:LastResultHtml) { Send-Text $resp $script:LastResultHtml }
            else { Send-Status $resp 404 "aucun resultat" }
        }
        '^/api/quit$'    {
            Send-Text $resp '{"ok":true}' "application/json; charset=utf-8"
            $running = $false
        }
        default          { Send-Status $resp 404 "not found" }
    }
}

try { $listener.Stop(); $listener.Close() } catch {}
Write-Host ""
Write-Host "  Serveur arrêté." -ForegroundColor Green
Write-Log "[WEB] Serveur arrete"
