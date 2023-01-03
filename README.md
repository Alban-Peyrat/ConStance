# ConStance - Contrôle Sudoc

[![Abandonned](https://img.shields.io/badge/Maintenance%20Level-Abandoned-orange.svg)](https://gist.github.com/cheerfulstoic/d107229326a01ff0f333a1d3476e068d)

ConStance est un outil visant originellement à contrôler des données du Sudoc / IdRef en utilisant [le webservice MARCXML de l'Abes](http://documentation.abes.fr/sudoc/manuels/administration/aidewebservices/index.html#SudocMarcXML). En pratique, elle permet également de donner certaines statistiques ou informations sur le fonds renseigné.

**Évitez d'avoir d'autres fichiers Excel ouverts pendant l'analyse (dans le cas où une erreur de programmation pourrait faire intéragir ConStance avec des fichiers non prévus).**

_Document en cours de réalisation_

## Initialisation

### À partir d'une liste de PPN d'Alma

Exportez d'Alma une liste de Titres physiques, renommez-la `export_Alma_ConStance.xlsx` et placez-la dans le même dossier que `ConStance.xlsm` (ConStance dans le reste de la documentation).

Allumez ConStance, choisissez la feuille `Introduction` et remettez à zéro les données. Importez ensuite les données d'Alma. Sélectionnez en `H2` le contrôle à effectuer (en appuyant sur Alt + flèche du bas, une liste déroulante s'affichera).

Lancez ensuite l'analyse.

### À partir d'une liste de PPN déjà établie

Allumez ConStance, choisissez la feuille `Introduction` et remettez à zéro les données. Collez votre liste de PPN dans la colonne associée. Sélectionnez en `H2` le contrôle à effectuer (en appuyant sur Alt + flèche du bas, une liste déroulante s'affichera).

Lancez ensuite l'analyse.

Notes : ConStance prend en compte les 9 derniers caractères de la cellule, si votre liste se présente sous la forme `PPN 123456789` ou `(PPN)123456789` ce n'est pas la peine de la retoucher, ni de rajouter des 0 en début de PPN, elle les ajoutera automatiquement.

## Les analyses

Au lancement de l'analyse, ConStance trie les PPN du plus petit au plus grand (tels qu'ils lui ont été donnés).

Si ConStance n'arrive pas à se connecter à la notice dans le Sudoc (ou IdRef), elle écrit sur toute la ligne une erreur avec le numéro d'entrée dans la liste originale et l'information qu'elle a utilisée comme PPN (les 9 derniers caractères de la cellule, plus des zéros en début si elle n'obtient pas 9 caractères via la manipulation initiale).

### CS1 : équivalence champs 103 / 200$f (IdRef)

Compare les données présentes en 103 avec ce qui est indiqué en 200$f (date de naissance et date de mort).

Pour chaque PPN présent dans la liste, ConStance récupère les dates présentes en 103$a et 103$b, analysant au passage si celles-ci sont incertaines ou avant Jésus-Christ. Elle récupère ensuite la 200$f et la décompose pour retirer les mêmes informations que la 103. Elle compare ensuite les dates en prenant en compte diverses manières d'écrire les dates incertaines (19.., 19XX) et signale si certaines sont mal orthographiées dans leur champ. 

Dans sa réponse finale, ConStance divise ses commentaires entre les dates de naissance, les dates de décès et le résultat final de l'analyse. Sont signalés en rouge les PPN qui ont des problèmes de correspondance, et plus spécifiquement en bleu les cellules des notes type format incorrect ou absence de telle ou telle date.

### CS2 : présence d'un lien en 700

Vérifie s'il y a un lien en 700.

Pour chaque PPN présent dans la liste, ConStance récupère et écrit le premier dollar de la 700. Voici les résultats possibles :
* `OK` en vert : le premier dollar est un 3 ;
* `Problème` en rouge : le premier dollar n'est pas un 3 (soit il n'y a pas de 700, soit c'est un autre dollar).

### CS3 : présence d'un lien en 7XX

Le fonctionnement est le même que CS2, sauf qu'à la place d'interroger uniquement la 700, ConStance interroge l'intégralité des 7XX. Le PPN est signalé en rouge si au moins un des champs ne contient pas un lien. Pour l'affichage des données, chaque 7XX se voit attribué un numéro d'occurrence (commençant à 0), c'est ce numéro qui est ensuite reporté dans la colonne des résultats pour identifier tous les champs posant problème.

### CS4 [non pref.] : statistiques d'âge (champs 210-214)

Calcule l'âge moyen et l'âge médian d'une liste de PPN (entre 1900 et 2030).

_Préférez utiliser [l'analyse CA2 de ConAn](https://github.com/Alban-Peyrat/ConAn#ca2--statistiques-d%C3%A2ge-en-prenant-en-compte-les-exemplaires) pour les titres physiques (de manière générale, ConAn prend en compte le nombre d'exemplaires, donc peut être plus intéressant). Pour une analyse sur les titres (exemplaires exclus), préférez utiliser CS7. Par ailleurs, CS5 vise à utiliser un champ supposément plus précis, donc préférez CS5 à CS4._

Pour chaque PPN, ConStance récupère tous les sous-champs 210$d et 214$d et isole les dates présentes dedans, conservant toujours la date la plus élevée, comprise entre 1900 et 2030 (exclus). Une fois l'intégralité des PPN traités, elle trie les PPN par âge puis calcule l'âge moyen et l'âge médian. Elle signalera en rouge les titres pour lesquelles elle n'a pas pu identifier la date (qui sont forcément envoyés au bas de la liste) et inscrira le nombre de titres exclus au total.

### CS5 [non pref.] : statistiques d'âge (champ 100)

Calcule l'âge moyen et l'âge médian d'une liste de PPN (entre 1900 et 2030).

_Préférez utiliser [l'analyse CA2 de ConAn](https://github.com/Alban-Peyrat/ConAn#ca2--statistiques-d%C3%A2ge-en-prenant-en-compte-les-exemplaires) pour les titres physiques (de manière générale, ConAn prend en compte le nombre d'exemplaires, donc peut être plus intéressant). Pour une analyse sur les titres (exemplaires exclus), préférez utiliser CS7_

Même fonctionnement que CS4 mais en utilisant cette fois-ci les sous-champs 100$a et 100$c, ce qui impliquerait donc plus de précision (l'analyse de CS4 peut exclure des PPN car ConStance n'arrive pas à isoler la date ou donner des dates erronées malgré les garde-fous).

### CS6 : détection de multiples éditions d'un même titre

_Notes : contrôle assez complexe. Réflexion en cours sur la précision à revoir ? sur la prise en compte des volumes ? si le taux pour la  comparaison des titres minimum à 100% devrait être réduit ?_

_A pour alternative [CA3 de ConAn](https://github.com/Alban-Peyrat/ConAn#ca3--d%C3%A9tection-de-multiples-%C3%A9ditions-dun-m%C3%AAme-titre). Chacun des deux à ses avantages et ses inconvénients, la détection initiale via la clef de titre est différente._

Détermine les titres qui ont possiblement deux éditions dans la même liste.

_L'explication détaillée arrivera une fois ma réflexion plus éclairée._ En quelques mots, ConStance récupère le titre, la mention d'édition et les PPN présents en 7XX de chaque PPN de notices bibliographiques, puis génère des clefs de titre en excluant notamment les déteminants définis et indéfinis. Elle compare ensuite pour chaque PPN contenant une mention d'édition la clef qu'elle lui a attribué à l'intégralité des autres clefs de la liste et défini un score de correspondance transformé en pourcentage : au-delà de 80%, elle juge possible que les titres soient corespondants. Lorsqu'elle compare les mots, elle ne compare pas s'ils sont parfaitement égaux mais si les 4 premières et les 4 dernières lettres correspondent, ce qui fait varier le score (pour prendre en compte de légers changements de terminologie). Par ailleurs, si l'intégralité du titre le plus court correspond à 100 % au début du titre le plus long, elle considère également que les titres sont correspondants. Ensuite, pour tous ces titres correspondants, elle calcule deux taux de correspondances des PPN de mentions de responsabilités, un excluant les collectivités et l'autre en les prenant en compte.

Mal fonctionnement possible : la mise en couleur du jaune et du bleu, avec le jaune qui apparaît trop souvent.

Sur la colonne de résultats :
* un même PPN peut apparaître plusieurs fois : cela est lié à la prise en compte du champ 451, le comportement est normal ;
* trois pourcentages s'affichent :
  * le premier concerne le pourcentage de correspondance du titre. Il est forcément supérieur à 80 sauf si la correspondance du titre le plus court est parfaite avec le début du plus long ;
  * le deuxième correspond au taux de correspondance des mentions de responsabilités, en excluant les champs 710 et plus ;
  * le troisième est le taux de correspondance de l'intégralité des mentions de responsabilités. Il est forcément supérieur ou égal au deuxième.
* le code couleur, sachant que cet ordre de priorité est utilisé :
  * vert : aucune détection automatique ;
  * rouge : correspondance de titres et au moins une mention de responsabilité commune ;
  * orange / jaune : correspondance de titres, pas de mention de responsabilité champ inférieur à 710 commune mais au moins une mention de responsabilité commune ;
  * bleu : correspondance de titres, mais pas de correspondance de mentions de responsabilité.

Recommandation : ne pas filtrer mais parcourir la liste et vérifier les noms adjacents.

### CS7 : statistiques d'âge (version double dates)

Calcule l'âge moyen et l'âge médian d'une liste de PPN.

Fusion de CS4 et CS5 : CS7 récupère à la fois l'année via la 100 et via la 21X puis calcule l'âge moyen et l'âge médian en fonction des deux méthodes. Le champ 100 est prioritaire, les résultats sont triés en fonction de celui-ci, la ligne est colorée de rouge si celui-ci est vide. Si la date en 21X est vide, seule la cellule correspondante est colorée. Les deux résultats sont affichés, ainsi que le nombre de titres exclus pour chacun d'entre eux.
