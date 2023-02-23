# conversion de rtf vers docx

Projet mené dans le cadre de la refonte de Papyrus en Kaligraf.

## première méthode : change_word_format()

Le chemin du fichier fournit à l'appli doit être absolu.
Dans le cas contraire, un chemin par défaut est utilisé, peut-être lié à la configuration de Word.

Il est à noter que les tests ne fonctionnent qu'avec l'extension .doc : le format .docx entraîne une erreur :

`Word a rencontré une erreur`

Il reste donc à régler ce problème concernant la fonction change_word_format().

## seconde méthode : ConvertRtfToDocx()

J'ai adapté le code pour en éliminer ce qui concerne la gestion des images et simplifier la déclaration des chemins.
Cette méthode réussit à convertir en docx !

Reste à tester avec davantage de fichiers, notamment les cases à cocher.

## source

[stackoverflow](https://stackoverflow.com/questions/65724760/how-to-convert-rtf-to-docx-in-python)