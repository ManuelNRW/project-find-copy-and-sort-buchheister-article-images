# project find, copy and sort buchheister article images
Finde jpg-Dateien in bestimmten Ordnern, anhand von angegebenen Matching Pattern, und kopiele diese in ein Zielverzeichnis.

In dem database Ordner befinden sich die Dateien destination_path.txt und source_path.txt
- destination_path.txt: Hier wird der Pfad zum Ordner angegeben, in welchen die Dateien kopiert werden sollen.
- source_path.txt: Hier können pro Zeile mehrere Pfade angegeben werden, in welchen die Dateien durchsucht werden sollen.
Achtung: Die Angabe des Pfades zu den Windows-Ordnern muss ohne Backslash am Ende erfolgen. Auch werden keine Unterordner durchsucht.

In der matching_phrases.txt wird in der ersten Zeile, ohne einem Slash, ein Unterordner angegeben, welcher für die zu kopierenden Bilder angelegt wird. Der Ordner wird im Verzeichnis der destination_path.txt erstellt.
Anschließend werden die zu suchenden Keywords, gesondert pro Zeile, hinterlegt. Sobald sich eines der Keywords in einem Dateinamen der Dateien befindet, wird diese Datei gematcht und in den Zielordner kopiert.
Mit einem * können Keywords aufgeteilt werden. Wenn Beispielsweise das Keyword night*day hinterlegt ist, würde beispielsweise auch die Datei Day&Night gefunden werden.
Auch ist die Suche nicht case sensitive, berücksichtigt also keine Groß- und Kleinschreibung.

Beispiel der mathing_phrases.txt:

```
day and night
day*night
134567
23456789
```

Hier sind drei Keywords hinterlegt. Die letzten beiden Keywords dienen exemplarisch als Artikelnummern.

Wichtig ist, dass die erste Zeile nicht für ein zu suchendes Keyword, sondern zur Anlage des Zielordners dient.
