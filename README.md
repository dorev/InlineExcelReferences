# InlineExcelReferences

Ce script permet d'insérer des valeurs contenues dans un document Excel **à l'intérieur du texte** d'un document Word.

## Usage

À même le texte, il suffit...

* d'écrire **{REF NomDuExcel.xlsx, NomDeLaFeuille, A6}{FINREF}** là où vous désirez obtenir la valeur du texte
* d'exécuter la macro **AcquérirValeursExternes** (dans l'onglet Affichage >> Macro)

Le code de la référence sera alors caché et la valeur contenu dans le fichier, la feuille et la cellule indiqué sera inséré entre les deux balises. Le format du texte inséré sera le même que celui qui est donné à la référence.

Pour faire réapparaître les références, il suffit d'exécuter la macro **AfficherRéférences**

## Ajout de ce script à votre document Word

1. Enregistrer votre document Word au format .docm (pour accepter les macros)
2. Ouvrir Visual Basic (Alt+F11)
3. S'assurer d'activer l'extension Excel dans Visual Basic (Outils >> Références, cocher Microsoft Excel)
4. Copier le contenu du fichier script.vba dans la page Visual Basic correspondant à votre document Word

