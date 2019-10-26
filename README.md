## Top 5 Audit at L'Oréal Perú

This project is meant to automize many reports/tables needed for audit processes related to our *Master Data* (Customers, Suppliers & Pricing).

It generates an R function that will do the heavy work of generating the tables (in Excel).

The data is confidential, so the data and reports generated are not shared in this repo.

Instructions:

- Customers data must be an Excel file contained in this folder (relative to your working directory): `./data/clientes/`. The code will read the last Excel file added/modified in the folder (according to the modification date).
- Suppliers data must be an Excel file contained in this folder (relative to your working directory): `./data/proveedores/<current_year>/`. The code will read the last Excel file added/modified in the folder (if today is 2017/06/20, the path for the data is `./data/proveedores/2017/`).
- Create a `reportes` folder in working directory. This will be the place where the reports will be written.
- Run the script `auditoriaTop5.R`:
```
source("auditoriaTop5.R")
```

- It will generate a function that you can use to generate the reports. Run this function:
```
Top5Answers()
```

- It will write a report (in Excel) inside the folder `reportes`