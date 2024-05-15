# T-ETL
## by AZOUGA Mourad
### Version 0.0.1
- Reads excel sheets from XML file 'register.xml'
- Spreads them into 3 classes (achat, ventes and articles)
- Create if not already exists a table for each class + Bilan table when called
- If anything is added to the excel sheet, and the sheet is read by the app it updates the table with the new values
- If the same file is called repeatedly it only reads it the first time
- Filters null values from the excel sheet
- Generates a Bilan on call with: bilan(num, art_id, qte_actuelle), qte_actuelle  = total des achats - total des ventes
- Stores new data in the bilan table when it's called and clears the previous ones