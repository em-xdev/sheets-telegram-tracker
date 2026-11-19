# Template builder

`build.py` generates the four spreadsheets in `templates/` and `demo/`. You only need it if you want to change the templates.

```bash
pip install openpyxl
cd tools
python3 build.py ../templates/Family_Budget.xlsx
python3 build.py ../templates/Presupuesto_Familiar.xlsx --es
python3 build.py ../demo/Family_Budget_DEMO.xlsx --demo
python3 build.py ../demo/Presupuesto_Familiar_DEMO.xlsx --demo --es
```

`howto.py` holds the text of the "How to use" tab in both languages. The demo data is generated (fixed random seed), so it is always the same and contains no real information.

If you rename a sheet, a column header or a Settings label, update the matching words in `bot/FamilyBudgetBot.gs` (`WB_ES` and the English names), because the bot finds everything by those names.
