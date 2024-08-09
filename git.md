# Utilisation de Git pour la gestion des sources VBA

## Installation

Lancer

    initGit.bat

## Sources

[lien](https://medium.com/@ivanzhd/use-git-hooks-for-version-control-of-excel-vba-code-261fda960fc5)

## Principe

Exploit Git hooks written in Python to dump the Excel VBA modules content into text files each time when you execute git commit command. This method is the most transparent and sustainable as it only uses basic tools and methods of Git, Python and VBA.

## Requirement

- Python
- oletools
`pip install -U oletools`
