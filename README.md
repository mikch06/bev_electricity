# bev_electricity
BEV Ladeabrechnung automatisiert ab WWZ pdf's - readyhome.

## Function
Mit diesem Script lassen sich readyhome Ladestationsabrechnung der WWZ einlesen.
Anschliessend wird nach Ladeschlüssel sortiert und abgerechnet, dies inkl. Mehrwertsteuer und Betriebspauschale. 

## Rquirements on the host:
    pip install pandas openpyxl --break-system-packages
    apt-get install poppler-utils tesseract-ocr tesseract-ocr-deu
