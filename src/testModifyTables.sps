SPSSINC modify tables /help.

FILE HANDLE samples NAME = "c:/spss17/samples/english".
get file='samples/cars.sav'.
compute foreign = origin > 1.
LOGISTIC REGRESSION VARIABLES foreign
  /METHOD=ENTER accel engine horse mpg 
  /PRINT=GOODFIT CORR ITER(1) CI(95)
  /CRITERIA=PIN(0.05) POUT(0.10) ITERATE(20) CUT(0.5).

output open FILE ="c:/python25/lib/site-packages/misc/tests/testSPSSmodifytables.spv".

* hide selected columns.
SPSSINC MODIFY TABLES subtype='Variables in the Equation' SELECT=2 3 'Upper'.

* set selected column widths.  pivot table editor will not necessarily honor all sizing.
SPSSINC MODIFY TABLES subtype='Variables in the Equation' select=2 3 'Upper'
/WIDTHS WIDTHS=30.

SPSSINC MODIFY TABLES subtype='Variables in the Equation' select=2 3 'Upper'
/WIDTHS WIDTHS=80.

* <<ALL>> option.
SPSSINC MODIFY TABLES SUBTYPE='Contingency Table' SELECT="<<ALL>>"
/WIDTHS WIDTHS=100.

* <<ALL>> on styles with level.
SPSSINC MODIFY TABLES SUBTYPE='Contingency Table' SELECT="<<ALL>>" LEVEL=-2
/STYLES BACKGROUNDCOLOR=0 200 0.


*negative indexing.
SPSSINC MODIFY TABLES subtype='Variables in the Equation' select=2 3 'Upper'
/WIDTHS WIDTHS=30 
/STYLES TEXTSTYLE=ITALIC TEXTCOLOR=0 0 255 BACKGROUNDCOLOR = 0 255 0.

*positive indexing.
SPSSINC MODIFY TABLES subtype='Variables in the Equation' select=2 3 'Upper'  LEVEL=2
/WIDTHS WIDTHS=30
/STYLES TEXTSTYLE=ITALIC TEXTCOLOR=0 0 255 BACKGROUNDCOLOR = 255 0 0.

*negative indexing - rows..
SPSSINC MODIFY TABLES subtype='Variables in the Equation' select='horse'
DIMENSION=ROWS
/STYLES TEXTSTYLE=ITALIC TEXTCOLOR=0 0 255 BACKGROUNDCOLOR = 0 0 255.

* positive indexing - rows.
SPSSINC MODIFY TABLES subtype='Variables in the Equation' select='mpg' 'Constant' LEVEL=2
DIMENSION=ROWS
/STYLES TEXTSTYLE=ITALIC TEXTCOLOR=0 0 255 BACKGROUNDCOLOR = 0 0 128.

* positive indexing - rows.
SPSSINC MODIFY TABLES subtype='Variables in the Equation' select='Step 1a' LEVEL=1
DIMENSION=ROWS
/STYLES TEXTSTYLE=ITALIC TEXTCOLOR=0 0 255 BACKGROUNDCOLOR = 200 200 200.

* using custom function.
SPSSINC MODIFY TABLES subtype='Variables in the Equation' select='Step 1a' LEVEL=1
DIMENSION=ROWS
/STYLES CUSTOMFUNCTION="customstylefunctions.stripeOddRows".

* wash background columns.
SPSSINC MODIFY TABLES subtype='Variables in the Equation' select="<<ALL>>"
DIMENSION=COLUMNS LEVEL=1
/STYLES APPLYTO=DATACELLS CUSTOMFUNCTION="customstylefunctions.washColumnBackgrounds".

* wash even numbered background columns.
SPSSINC MODIFY TABLES subtype='Variables in the Equation' select=0  2  4  6 
DIMENSION=COLUMNS LEVEL=1
/STYLES APPLYTO=BOTH CUSTOMFUNCTION="customstylefunctions.washColumnBackgrounds".

* blue text and green background for B and Sig columns.
SPSSINC MODIFY TABLES subtype='variables in the equation' SELECT="B" "Sig."
/STYLES TEXTCOLOR = 0 0 255 BACKGROUNDCOLOR=0 255 0.

* Bold only selected column labels.
SPSSINC MODIFY TABLES subtype='Variables in the Equation' select=2 3 'Upper'
/STYLES TEXTSTYLE=BOLD APPLYTO=LABELS.

* style only datacells: bold italic white.
SPSSINC MODIFY TABLES subtype='Variables in the Equation' select=2 3 'Upper'
/STYLES TEXTSTYLE=BOLDITALIC APPLYTO=DATACELLS TEXTCOLOR=255 255 255.

* blue text for all table types in previous procedure.
SPSSINC MODIFY TABLES SUBTYPE=* SELECT = 0 2 4 6 8
/STYLES TEXTCOLOR = 0 0 255.

SPSSINC MODIFY TABLES subtype='Variables in the Equation' select=1 3 5
DIMENSION= ROWS
/STYLES TEXTSTYLE=BOLDITALIC APPLYTO=DATACELLS TEXTCOLOR=255 255 255.

SPSSINC MODIFY TABLES subtype='Variables in the Equation' SELECT='Wald'
/WIDTHS WIDTHS=30.

SPSSINC MODIFY TABLES subtype='Variables in the Equation' select=0 2 4
DIMENSION= ROWS
/STYLES TEXTSTYLE=BOLDITALIC APPLYTO=BOTH TEXTCOLOR=255 0 0.

SPSSINC MODIFY TABLES subtype='Variables in the Equation' select=0 2 4
DIMENSION= COLUMNS
/STYLES TEXTSTYLE=BOLDITALIC APPLYTO=BOTH TEXTCOLOR=255 0 0.

SPSSINC MODIFY TABLES subtype='Variables in the Equation' select=2 3 'Upper'
/WIDTHS WIDTHS=70 60 50.

SPSSINC MODIFY TABLES subtype='Variables in the Equation' select=0 1 2 3 'upper'
/WIDTHS WIDTHS=20.

SPSSINC MODIFY TABLES subtype='Variables in the Equation' select= 'Lower' 'Upper'
/STYLES APPLYTO=LABELS TEXTCOLOR=16777215 BACKGROUNDCOLOR=255 0 0.

SPSSINC MODIFY TABLES subtype='Variables in the Equation' select= '95% C.I.for EXP(B)' LEVEL=-2
/STYLES APPLYTO=LABELS TEXTCOLOR=255 BACKGROUNDCOLOR=0 255 0.

SPSSINC MODIFY TABLES subtype='Variables in the Equation' select=0 2 4
DIMENSION= ROWS
/STYLES  APPLYTO=BOTH BACKGROUNDCOLOR=0 255 0.

SPSSINC MODIFY TABLES subtype= 'Variables in the Equation'
/WIDTHS ROWLABELS=-1 -2 ROWLABELWIDTHS=80 60.

GET FILE='samples/employee data.sav'.
***format salary(f8.2).
CTABLES
  /VLABELS VARIABLES=jobcat gender bdate salary DISPLAY=DEFAULT
  /TABLE jobcat + gender BY bdate [MEAN] + salary [MEAN]
  /CATEGORIES VARIABLES=jobcat gender ORDER=A KEY=VALUE EMPTY=INCLUDE TOTAL=YES POSITION=AFTER 
    MISSING=EXCLUDE.
SPSSINC MODIFY TABLES subtype='Custom Table' SELECT="Total" DIMENSION=ROWS
/STYLES TEXTSTYLE=BOLD.
SPSSINC MODIFY TABLES subtype='Custom Table' SELECT=-1 PROCESS=PRECEDING
/STYLES TEXTCOLOR=0 0 255 TEXTSTYLE=BOLD APPLYTO="x > 50000".
*next command should do nothing, because selected column is a date.
SPSSINC MODIFY TABLES subtype='Custom Table' SELECT=0 PROCESS=PRECEDING
/STYLES TEXTCOLOR=0 0 255 TEXTSTYLE=BOLD APPLYTO="x > 50000".
SPSSINC MODIFY TABLES subtype='Custom Table' SELECT=-1 PROCESS=ALL
/STYLES TEXTCOLOR=0 255 255 TEXTSTYLE=BOLD APPLYTO="x > 50000".
DATASET ACTIVATE DataSet1.
LOGISTIC REGRESSION VARIABLES foreign
  /METHOD=ENTER accel engine horse mpg 
  /PRINT=GOODFIT CORR ITER(1) CI(95)
  /CRITERIA=PIN(0.05) POUT(0.10) ITERATE(20) CUT(0.5).


SPSSINC modify tables subtype="Custom Table" select=0 1 2/styles backgroundcolor=0 0 128
textcolor=0 255 0 /widths widths=100.

SPSSINC modify tables subtype="Custom Table" select="Total"
dimension=rows
/styles backgroundcolor=0 0 128
textcolor=0 255 0.

* Use qualitative scheme.
SPSSINC MODIFY TABLES subtype='Contingency Table' select="<<ALL>>" process=ALL
/styles applyto=datacells customfunction="customstylefunctions.qualitative".

* Use Pastel Qualitative scheme.
SPSSINC MODIFY TABLES subtype=* select="<<ALL>>" process=ALL
/styles applyto=datacells customfunction="customstylefunctions.pastelqualitative".

* overflow color set.
CTABLES
  /TABLE BY year [C][COUNT F40.0].
SPSSINC MODIFY TABLES subtype='Custom Table' select="<<ALL>>" 
/styles applyto=datacells customfunction="customstylefunctions.pastelqualitative".

SPSSINC MODIFY TABLES subtype='Custom Table' select="<<ALL>>" 
/styles applyto=labels customfunction="customstylefunctions.pastelqualitative".

SPSSINC MODIFY TABLES subtype='Custom Table' select="<<ALL>>" level=0
/styles applyto=labels customfunction="customstylefunctions.pastelqualitative".

SPSSINC MODIFY TABLES subtype='Custom Table' select="<<ALL>>" 
/styles applyto=datacells customfunction="customstylefunctions.qualitative".
get file='samples/cars.sav'.
compute foreign = origin > 1.
REGRESSION
  /STATISTICS COEFF OUTS R ANOVA
  /DEPENDENT mpg
  /METHOD=ENTER accel engine horse weight foreign.
SPSSINC MODIFY TABLES subtype='Coefficients' select=0
/styles applyto=datacells customfunction="customstylefunctions.makeSigCoefsBold".


* Bold the total summaries.
CTABLES
  /TABLE year BY origin > mpg [MEAN]
  /CATEGORIES VARIABLES=year ['70', '71', '72', SUBTOTAL='Average', '73', '74', '75', 
    SUBTOTAL='Average', '76', '77', '78', SUBTOTAL='Average', '79', '80', '81', '82', 
    SUBTOTAL='Average', OTHERNM] EMPTY=INCLUDE POSITION=AFTER.
SPSSINC MODIFY TABLES SUBTYPE="Custom Table" SELECT = "Average"
DIMENSION=ROWS
/STYLES BACKGROUNDCOLOR=192 150 150 TEXTSTYLE = BOLD TEXTCOLOR=255 255 255.

* Bold the counts in a crosstab table.
CROSSTABS
  /TABLES=cylinder BY origin
  /CELLS=COUNT ROW COLUMN TOTAL.
SPSSINC MODIFY TABLES SUBTYPE='Crosstabulation' SELECT="Count" DIMENSION=ROWS
/STYLES TEXTSTYLE=BOLD.

* Bold the cells with large residuals.
CROSSTABS  /TABLES=cylinder BY origin
  /CELLS=COUNT SRESID TOTAL.
SPSSINC MODIFY TABLES SUBTYPE='Crosstabulation' SELECT="Std. Residual"
DIMENSION = ROWS
/STYLES TEXTSTYLE=BOLD APPLYTO="abs(x) > 2".

* set background, text color, and text style for totals.
CTABLES
  /TABLE jobcat [C] + gender [C] BY bdate [S][MEAN] + salary [S][MEAN]
  /CATEGORIES VARIABLES=jobcat gender ORDER=A KEY=VALUE EMPTY=INCLUDE TOTAL=YES LABEL='Average' 
    POSITION=AFTER MISSING=EXCLUDE.
SPSSINC MODIFY TABLES SUBTYPE="Custom Table" SELECT = "Average"
DIMENSION=ROWS
/STYLES BACKGROUNDCOLOR=150 150 150 TEXTSTYLE = BOLD TEXTCOLOR=255 255 255.
DATASET ACTIVATE DataSet7.
* Custom Tables.
get file="samples/telco.sav".
dataset name telco.
CTABLES
  /TABLE region [C] > ed [C][COUNT F40.0] BY custcat [C]
  /CATEGORIES VARIABLES=region ed TOTAL=YES POSITION=AFTER.
SPSSINC MODIFY TABLES SUBTYPE="Custom Table" SELECT = "Total"
DIMENSION=ROWS
/STYLES BACKGROUNDCOLOR=255 255 88 TEXTSTYLE = BOLD.

* Custom Tables.


SPSSINC MODIFY TABLES subtype="Rotated Factor Matrix"
SELECT="<<ALL>>" PROCESS = PRECEDING
/STYLES  APPLYTO="abs(x) < .2"
TEXTCOLOR = 255 255 255.

DATASET ACTIVATE DataSet0.
SPSSINC MODIFY TABLES subtype="'Classification Table'"
SELECT=1 
REGEXP=yes DIMENSION= COLUMNS
LEVEL = -1  PROCESS = PRECEDING
/STYLES  APPLYTO=BOTH 
TEXTSTYLE=italic.

* custom functions with parameters.
SPSSINC MODIFY TABLES subtype="'Variables in the Equation'"
SELECT="<<ALL>>" 
DIMENSION= ROWS
LEVEL = -1  PROCESS = PRECEDING 
/STYLES  APPLYTO=DATACELLS 
CUSTOMFUNCTION="customstylefunctions.stripeOddRows2(r=20,g=20,b=200)"
"mycustomfuncs.second(r=200, g=50, b=50)".

SPSSINC MODIFY TABLES subtype='Contingency Table'
SELECT="<<ALL>>" 
DIMENSION= COLUMNS
LEVEL = -1  PROCESS = PRECEDING 
/STYLES  APPLYTO=BOTH 
CUSTOMFUNCTION="customstylefunctions.washColumns(color='red')".

SPSSINC MODIFY TABLES subtype='Contingency Table'
SELECT="<<ALL>>" 
DIMENSION= COLUMNS
LEVEL = -1  PROCESS = PRECEDING 
/STYLES  APPLYTO=BOTH 
CUSTOMFUNCTION="customstylefunctions.washColumns(color='green')".

* invalid custom function.
SPSSINC MODIFY TABLES SUBTYPE='Contingency Table' SELECT="<<ALL>>" LEVEL=-2
/STYLES BACKGROUNDCOLOR=0 200 0 customfunction="mycustomfuncs.badfunction".

SPSSINC MODIFY TABLES subtype="*"
SELECT=1 
DIMENSION= COLUMNS
LEVEL = -1  PROCESS = PRECEDING 
/STYLES  APPLYTO=DATACELLS.


DATASET NAME DataSet3 WINDOW=FRONT.
COMMENT BOOKMARK;LINE_NUM=47;ID=3.
COMMENT BOOKMARK;LINE_NUM=62;ID=5.
COMMENT BOOKMARK;LINE_NUM=67;ID=4.
COMMENT BOOKMARK;LINE_NUM=75;ID=1.
COMMENT BOOKMARK;LINE_NUM=153;ID=2.
