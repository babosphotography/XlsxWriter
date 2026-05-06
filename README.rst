import pandas as pd
from datetime import datetime

with pd.ExcelWriter('PV_Dimensionnement_Complete_Victron.xlsx', engine='openpyxl') as writer:
    
    # 1. Projet
    project = pd.DataFrame({
        'Paramètre': ['Nom Projet', 'Localisation', 'Latitude', 'Longitude', 'Type Système', 'Puissance PV (kWp)', 'Tension Batterie', 'Autonomie (jours)', 'Date'],
        'Valeur': ['', 'Dakar, Sénégal', '14.69', '-17.44', 'Off-Grid / Hybrid', '', '48V', '2-3', datetime.now().strftime('%d/%m/%Y')]
    })
    project.to_excel(writer, sheet_name='1_Projet', index=False)
    
    # 2. Profil de Consommation
    load = pd.DataFrame({
        'Heure': list(range(0,24)),
        'Puissance (W)': [150,120,100,80,100,400,900,1400,1100,800,700,650,700,750,900,1100,1400,1600,1300,900,700,500,300,200],
        'Commentaire': ['Nuit','','','','','Lever','Matin','Pic matin','','','','','','','','','Soir','Pic soir','','','','','Nuit']
    })
    load.to_excel(writer, sheet_name='2_Load_Profile', index=False)
    
    # 3. Modules PV (exemples)
    modules = pd.DataFrame({
        'Fabricant': ['Jinko', 'Trina', 'Longi', 'JA Solar', 'Canadian Solar'],
        'Modèle': ['Tiger Neo 555W', 'Vertex S 450W', 'Hi-MO 6 580W', 'DeepBlue 550W', 'HiKu7 550W'],
        'Puissance STC (Wp)': [555,450,580,550,550],
        'Voc (V)': [49.5,45.2,50.1,49.8,49.5],
        'Isc (A)': [13.9,11.2,14.2,13.8,13.7],
        'Vmp (V)': [41.5,37.8,42.0,41.6,41.3],
        'Imp (A)': [13.4,11.9,13.8,13.2,13.3],
        'Efficacité (%)': [22.5,22.0,22.8,21.5,21.6]
    })
    modules.to_excel(writer, sheet_name='3_Modules_PV', index=False)
    
    # 4. Régulateurs Victron + Autres
    controllers = pd.DataFrame({
        'Marque': ['Victron']*12 + ['Autres'],
        'Modèle': ['SmartSolar MPPT 75/15','SmartSolar MPPT 100/30','SmartSolar MPPT 100/50','SmartSolar MPPT 150/35','SmartSolar MPPT 150/60','SmartSolar MPPT 150/70','SmartSolar MPPT 250/60','SmartSolar MPPT 250/100','MPPT RS 450/100','MPPT RS 450/200','BlueSolar MPPT 100/50','SmartSolar MPPT 100/20','EPEVER Tracer 4210AN'],
        'Type': ['MPPT']*12 + ['MPPT'],
        'Tension Batterie': ['12/24/48V']*11 + ['12/24/48V', '12/24V'],
        'Courant (A)': [15,30,50,35,60,70,60,100,100,200,50,20,40],
        'PV Max Voc (V)': [75,100,100,150,150,150,250,250,450,450,100,100,100],
        'Puissance Max 48V (W)': [880,1740,2900,2000,3480,4000,3480,5800,5800,11520,2900,1160,0],
        'Bluetooth': ['Oui','Oui','Oui','Oui','Oui','Oui','Oui','Oui','Oui','Oui','Non','Oui','Non']
    })
    controllers.to_excel(writer, sheet_name='4_Regulateurs', index=False)
    
    # 5. Onduleurs / Inverters Victron + Autres (très enrichi)
    inverters = pd.DataFrame({
        'Marque': ['Victron']*15 + ['Autres']*6,
        'Modèle': ['MultiPlus-II 48/3000','MultiPlus-II 48/5000','MultiPlus-II 48/8000','MultiPlus-II 48/10000','Quattro 48/5000','Quattro 48/8000','Quattro 48/10000','Quattro 48/15000','Multi RS Solar 48/6000','Inverter RS 48/6000','MultiPlus 48/5000','MultiPlus-II GX 48/5000','EasySolar-II GX 48/5000','MultiPlus 48/20kW','Inverter Compact 24/2000',
                   'Deye SUN-12K','Growatt SPF 5000 ES','SMA Sunny Boy 5.0','Fronius Primo 6.0','Goodwe ES 5kW','Luxpower 12k'],
        'Type': ['Hybrid','Hybrid','Hybrid','Hybrid','Hybrid (2 AC In)','Hybrid (2 AC In)','Hybrid','Hybrid','Hybrid avec MPPT','Off-Grid avec MPPT','Hybrid','Hybrid GX','All-in-One','Hybrid','Off-Grid',
                 'Hybrid','Off-Grid/Hybrid','Grid-Tie','Grid-Tie','Hybrid','Hybrid'],
        'Puissance (VA/W)': [3000,5000,8000,10000,5000,8000,10000,15000,6000,6000,5000,5000,5000,20000,2000,
                             12000,5000,5000,6000,5000,12000],
        'Tension Batterie': ['48V']*14 + ['24V','48V','48V','','','48V','48V','48V'],
        'MPPT intégré': ['Non','Non','Non','Non','Non','Non','Non','Non','Oui','Oui','Non','Non','Oui','Non','Non',
                         'Oui','Oui','Non','Non','Oui','Oui'],
        'Remarque': ['Best-seller ESS','Très polyvalent','Puissant','Haut de gamme','2 entrées AC','2 entrées AC','Très puissant','Pour grosses installations','MPPT intégré','RS Smart','Classique','Avec GX intégré','All-in-One','Très gros système','Compact',
                     'Populaire','Excellent rapport','Grid-Tie','Grid-Tie','Bon marché','Puissant']
    })
    inverters.to_excel(writer, sheet_name='5_Onduleurs', index=False)
    
    # 6. Composantes Complètes du Système PV
    components = pd.DataFrame({
        'Catégorie': ['Panneaux PV','Régulateur MPPT','Onduleur','Batterie','Câblage DC','Protections','Monitoring','Structure','Autres'],
        'Composantes': ['Modules PV (Mono PERC, TOPCon, etc.)','Victron SmartSolar / RS','MultiPlus / Quattro / RS','Lithium LiFePO4 (BYD, Pylontech, Victron, etc.)','Câbles solaires 4-16mm²','Disjoncteurs DC, Fusibles, Parafoudres','Cerbo GX, Touch GX, VRM Portal','Structures fixes / trackers','Compteur Bidirectionnel, Transfert ATS, etc.'],
        'Rôle': ['Production DC','Charge batterie optimisée','Conversion DC→AC + Chargeur','Stockage énergie','Transport courant','Sécurité','Suivi & Contrôle','Fixation panneaux','Intégration au site'],
        'Exemples Victron': ['---','SmartSolar 250/100','MultiPlus-II 48/5000','---','---','---','Cerbo GX + GX Touch','---','VE.Bus, VE.Can']
    })
    components.to_excel(writer, sheet_name='6_Composantes_Systeme', index=False)
    
    # 7. Calculs (avec formules)
    calc = pd.DataFrame({
        'Paramètre': ['Energie journalière (Wh)', 'PSH (heures soleil)', 'Puissance PV requise (Wp)', 'Nombre de modules', 'Nombre de MPPT', 'Capacité Batterie (kWh)', 'Nombre Onduleurs'],
        'Valeur / Formule': ['=SOMME(2_Load_Profile!B2:B25)', '4.5', '=B2/(C2*0.75)', '', '', '', ''],
        'Explication': ['Profil de charge', 'Peak Sun Hours', 'Avec pertes système ~25%', 'Selon puissance module', 'Selon courant PV', 'Selon autonomie et DoD', 'Selon puissance requise']
    })
    calc.to_excel(writer, sheet_name='7_Calculs', index=False)

print("✅ Fichier 'PV_Dimensionnement_Complete_Victron.xlsx' généré avec succès !")
print("Il contient maintenant tous les principaux modèles Victron + liste complète des composantes.")XlsxWriter
==========

**XlsxWriter** is a Python module for writing files in the Excel 2007+ XLSX
file format.

XlsxWriter can be used to write text, numbers, formulas and hyperlinks to
multiple worksheets and it supports features such as formatting and many more,
including:

* 100% compatible Excel XLSX files.
* Full formatting.
* Merged cells.
* Defined names.
* Charts.
* Autofilters.
* Data validation and drop down lists.
* Conditional formatting.
* Worksheet PNG/JPEG/GIF/BMP/WMF/EMF images.
* Rich multi-format strings.
* Cell comments.
* Integration with Pandas and Polars.
* Textboxes.
* Support for adding Macros.
* Memory optimization mode for writing large files.

It supports Python 3.8+ and PyPy3 and uses standard libraries only.

Here is a simple example:

.. code-block:: python

   import xlsxwriter

   # Create an new Excel file and add a worksheet.
   workbook = xlsxwriter.Workbook("demo.xlsx")
   worksheet = workbook.add_worksheet()

   # Widen the first column to make the text clearer.
   worksheet.set_column("A:A", 20)

   # Add a bold format to use to highlight cells.
   bold = workbook.add_format({"bold": True})

   # Write some simple text.
   worksheet.write("A1", "Hello")

   # Text with formatting.
   worksheet.write("A2", "World", bold)

   # Write some numbers, with row/column notation.
   worksheet.write(2, 0, 123)
   worksheet.write(3, 0, 123.456)

   # Insert an image.
   worksheet.insert_image("B5", "logo.png")

   workbook.close()

.. image:: https://raw.github.com/jmcnamara/XlsxWriter/master/dev/docs/source/_images/demo.png

See the full documentation at: https://xlsxwriter.readthedocs.io

Release notes: https://xlsxwriter.readthedocs.io/changes.html

