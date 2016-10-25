# ExpUsers
Programme Java permettant d'extraire les utilisateurs d'un site Web dans un fichier Excel

##Utilisation:
```
java ExpUsers [-dbserver db] [-o fichier.xlsx] [-d] [-t] 
```
o� :
* ```-dbserver db``` est la r�f�rence � la base de donn�es, par d�faut d�signe la base de donn�es de d�veloppement. Voir fichier *myDatabases.prop* (optionnel).
* ```-o fichier.xlsx``` est le nom du fichier Excel qui recevra les utilisateurs. Amorc� � *users.xlsx* par d�faut (param�tre optionnel).
* ```-d``` le programme s'ex�cute en mode d�bug, il est beaucoup plus verbeux. D�sactiv� par d�faut (param�tre optionnel).
* ```-t``` le programme s'ex�cute en mode test, les transcations en base de donn�es ne sont pas faites. D�sactiv� par d�faut (param�tre optionnel).

##Pr�-requis :
- Java 6 ou sup�rieur.
- JDBC Informix
- JDBC MySql
- Driver MongoDB
- [xmlbeans-2.6.0.jar](https://xmlbeans.apache.org/)
- [commons-collections4-4.1.jar](https://commons.apache.org/proper/commons-collections/download_collections.cgi)

##Fichier des param�tres : 

Ce fichier permet de sp�cifier les param�tres d'acc�s aux diff�rentes bases de donn�es.

A adapter selon les impl�mentations locales.

Ce fichier est nomm� : *MyDatabases.prop*.

Le fichier *MyDatabases_Example.prop* est fourni � titre d'exemple.

##R�f�rences:

- [API Java Exel POI](http://poi.apache.org/download.html)
- [Tuto Java POI Excel](http://thierry-leriche-dessirier.developpez.com/tutoriels/java/charger-modifier-donnees-excel-2010-5-minutes/)
- [Tuto Java POI Excel](http://jmdoudoux.developpez.com/cours/developpons/java/chap-generation-documents.php)
