# ExpUsers
Programme Java permettant d'extraire les utilisateurs d'un site Web dans un fichier Excel

##Utilisation:
```
java ExpUsers [-dbserver db] [-o fichier.xlsx] [-d] [-t] 
```
où :
* ```-dbserver db``` est la référence à la base de données, par défaut désigne la base de données de développement. Voir fichier *myDatabases.prop* (optionnel).
* ```-o fichier.xlsx``` est le nom du fichier Excel qui recevra les utilisateurs. Amorcé à *users.xlsx* par défaut (paramètre optionnel).
* ```-d``` le programme s'exécute en mode débug, il est beaucoup plus verbeux. Désactivé par défaut (paramètre optionnel).
* ```-t``` le programme s'exécute en mode test, les transcations en base de données ne sont pas faites. Désactivé par défaut (paramètre optionnel).

##Pré-requis :
- Java 6 ou supérieur.
- JDBC Informix
- JDBC MySql
- Driver MongoDB

##Fichier des paramètres : 

Ce fichier permet de spécifier les paramètres d'accès aux différentes bases de données.

A adapter selon les implémentations locales.

Ce fichier est nommé : *MyDatabases.prop*.

Le fichier *MyDatabases_Example.prop* est fourni à titre d'exemple.

##Références:

- [API Java Exel POI](http://poi.apache.org/download.html)
- [Tuto Java POI Excel](http://thierry-leriche-dessirier.developpez.com/tutoriels/java/charger-modifier-donnees-excel-2010-5-minutes/)
- [Tuto Java POI Excel](http://jmdoudoux.developpez.com/cours/developpons/java/chap-generation-documents.php)
