
get file="c:/spss18/samples/english/employee data.sav".
compute random = rv.normal(0,1).
SORT CASES  BY jobcat.
SPLIT FILE LAYERED BY jobcat.
split file off.
CORRELATIONS
  /VARIABLES=salbegin salary random bdate educ minority
  /PRINT=TWOTAIL NOSIG
  /MISSING=PAIRWISE.

*do all defaults.
SPSSINC MODIFY TABLES SUBTYPE='Correlations' SELECT=0 DIMENSION=ROWS
/styles customfunction="formatcorrmat.cleancorr".


CORRELATIONS
  /VARIABLES=salbegin salary bdate educ
  /PRINT=TWOTAIL NOSIG
  /MISSING=PAIRWISE.
NONPAR CORR
  /VARIABLES=salbegin salary bdate educ
  /PRINT=BOTH TWOTAIL NOSIG
  /MISSING=PAIRWISE.
