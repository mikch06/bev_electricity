# bev_electricity
BEV Ladeabrechnung automatisiert ab WWZ pdf's - readyhome.

## Function
Mit diesem Script lassen sich readyhome Ladestationsabrechnung der WWZ einlesen.
Anschliessend wird nach Ladeschlüssel sortiert und abgerechnet, dies inkl. Mehrwertsteuer und Betriebspauschale. 

## Rquirements on the host:
    pip install pandas openpyxl --break-system-packages
    apt-get install poppler-utils tesseract-ocr tesseract-ocr-deu

## Setup
Clone this repo:

    git clone https://github.com/mikch06/bev_electricity

Create virtual environment:

    python3 -m venv venv

Activate venv:

    source venv/bin/activate

Run script:

    python3 wwz_ladenachweis_auswerten.py ./pdf-directory
