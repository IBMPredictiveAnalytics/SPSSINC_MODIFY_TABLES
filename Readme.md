# SPSSINC MODIFY TABLES
## Modify the appearance of pivot tables.
 This procedure allows you to modify the appearance of data cells and row and column labels. You can modify text style, text color or background color. You can also set column widths or the width of row labels and you can hide specified rows or columns.

**Note: Consider using the built-in OUTPUT MODIFY command**.

---
Requirements
----
- IBM SPSS Statistics 18 or later and the corresponding IBM SPSS Statistics-Integration Plug-in for Python.

---
Installation intructions
----
1. Open IBM SPSS Statistics
2. Navigate to Utilities -> Extension Bundles -> Download and Install Extension Bundles
3. Search for the name of the extension and click Ok. Your extension will be available.

---
Tutorial
----

### Installation Location

Utilities →

&nbsp;&nbsp;Modify Table Appearance

### UI
<img width="893" alt="image" src="https://user-images.githubusercontent.com/19230800/196509728-2353219b-b9eb-417a-a5a2-d203b1bbd319.png">
<img width="1071" alt="image" src="https://user-images.githubusercontent.com/19230800/196509755-1f003ca7-0c30-4248-b5f2-c2e309b9ee16.png">
<img width="513" alt="image" src="https://user-images.githubusercontent.com/19230800/196509793-f3f4f762-c5bd-4d88-9ab9-8197157ec650.png">

### Syntax
Example
> SPSSINC MODIFY TABLES subtype="Frequencies" <br />
> SELECT=Count  <br />
> DIMENSION= COLUMNS LEVEL = -1  SIGLEVELS=BOTH  <br />
> PROCESS = PRECEDING  <br />
> /STYLES  APPLYTO=DATACELLS. <br />

---
License
----

- Apache 2.0
                              
Contributors
----

  - IBM SPSS
