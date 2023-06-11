# Projet "Calendrier"

Outil d'extraction du service annuel à partir du calendrier public de La Rochelle Université.

Le TEA n'est pas pris en compte dans le calcul du service car il dépend du nombre de groupe de 10 étudiants. Il faut également ajouter ensuite tout ce qui relève du REH et de la FC.

Calcul heures supplémentaires pour les EC : 
* Si le sous-total (CM*1,5 + TD + TP) est < au service dû de l’enseignant : toutes les heures TP sont comptabilisées pour 1 HeTD
* Si le sous-total (CM*1,5 + TD + TP) est > au service dû de l’enseignant :
1. On calcule la part des TP dans ce sous-total, et on reporte cette part pour identifier les heures TP comptabilisées pour 1 HeTD
2. On soustrait ce volume horaire au volume d’heures TP total pour identifier les
heures TP comptabilisées pour 2/3 HETD

## Récupération du code source

```
git clone https://gitlab.univ-lr.fr/dmondo01/calendrier.git
```

### Prérequis

Python 3.7 ou plus

Module xslxWriter :
```
pip install XlsxWriter
```

Module ics :
```
pip install ics
```

### Utilisation

Dans le fichier main.py, indiquez votre login ULR, le nombre d'heures que vous devez 
effectuer dans votre service, votre type (ATER, EC, PRAG, PRCE ou VACATAIRE), la date de début d'année à partir de laquelle le décompte des heures s'effectuera et en option une date de fin
```
time_table = TimeTable("LOGIN_ULR", 384, TeacherType.PRAG, datetime(2022, 9, 1))
time_table = TimeTable("LOGIN_ULR", 192, TeacherType.ATER, datetime(2021, 9, 1), datetime(2022, 8, 31))
```

Dans le second exemple :

login = "LOGIN_ULR"

nombre d'heures = 192 hetd

type enseignant = ATER

date de début = 01/09/2021

date de fin = 31/08/2022




## Author

* **Damien Mondou** 