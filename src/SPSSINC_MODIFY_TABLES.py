#/***********************************************************************
# * Licensed Materials - Property of IBM 
# *
# * IBM SPSS Products: Statistics Common
# *
# * (C) Copyright IBM Corp. 1989, 2020
# *
# * US Government Users Restricted Rights - Use, duplication or disclosure
# * restricted by GSA ADP Schedule Contract with IBM Corp. 
# ************************************************************************/

# SPSSINC MODIFY TABLES extension command

"""This module implements the SPSS MODIFY TABLES extension command.
It delegates the implementation to the modifytables.py module."""

__author__ = "SPSS, JKP"
__version__ = "1.4.2"

# history
# 07-jan-2015 add countinvis keyword
# 25-nov-2015 add significance keywords


from extension import Template, Syntax, processcmd

import modifytables

#try:
    #import wingdbstub
#except:
    #pass

helptext="""SPSS MODIFY TABLES SUBTYPE=subtypes SELECT=list of columns or rows to operate on
    [PROCESS={PRECEDING* | ALL}
    [DIMENSION={COLUMNS* | ROWS}] 
    [LEVEL=number]
    [HIDE={TRUE|FALSE}
    [REGEXP={NO*|YES}]
    [PRINTLABELS={YES|NO*}]
[/WIDTHS [WIDTHS=list-of-widths] [ROWLABELS=list of row label numbers] 
    [ROWLABELVALUES=list of widths]]
[/STYLES [TEXTSTYLE={REGULAR|BOLD|ITALIC|BOLDITALIC}]
    [TEXTCOLOR=RGB values]
    [BACKGROUNDCOLOR=RGB values]
    [CUSTOMFUNCTION="module.function" ...]
    [TLOOK="filespec"]
    [APPLYTO={BOTH* | LABELS | DATACELLS| True/False expression}]
[/HELP].

Operating on the specified set of rows or columns, you can
- hide them
- set the widths (columns and row labels only)
- set style characteristics for the selected row or column labels and/or values
- call custom functions.

SUBTYPE is the OMS table subtype   It specifies which types of tables to process.
You can find the subtype by right clicking in the outline on an instance 
or from Utilties/OMS identifiers.
The subtype name should be in quotes.  Any extra matching outer quotes are removed.
Case and white space do not matter.  Multiple subtypes can be specified.
SUBTYPE="*" processes tables regardless of subtype.

By default, the immediately preceding procedure output is processed.  Specify
PROCESS=ALL to process all existing tables matching the specified subtypes.

SELECT is a list of one or more data columns or rows to modify.
They can be specified by number, counting from zero, or by the text at the 
chosen level in the row or column header.
Use negative numbers to count backwards from the end:
-1 is the last row or column.
When selecting based on the text in the label, put the text in quotes.  The case
and white space in the text matter when matching.
If the text includes a footnote, add the footnote character to the end of the cell text.
Use "<<ALL>>" to match all rows or columns.
Out of range column numbers and nonmatching text are ignored.

If REGEXP=YES, the text is interpreted as a regular expression.  This can be
used to match patterns.  For example, the pattern
"^Sig"
would match any label starting with Sig.

By default, when selecting by the row or column header text, the lowest or
innermost text is tested.  LEVEL can be specified to use outer layers
of the labels.  LEVEL=-1, the default, is the innermost layer.  More negative
numbers move out.  Positive numbers  count in from the outside.  
LEVEL=1 is usually the first visible cell (level 0 is the dimension label).

In some cases, there may be invisible levels inbetween
what you see, so some experimentation may be required to find the level
you need.  The level specification only affects text matching.  Column and row
numbers always count at the innermost level.

Use PRINTLABELS=TRUE to display the full label structure of selected tables
in the specified dimension in order to assist in specifying the level.

Note that hiding a category hides that category in all dimensions.

DIMENSION=COLUMNS, the default, indicates operating on columns.
DIMENSION=ROWS causes rows to be operated on.

HIDE=TRUE causes the selected rows or columns to be hidden.  It cannot
be combined with any other specification and is assumed if no width or
styles are specified.

Specify WIDTHS=list of widths to set the width of the specified
columns to the specified values in points (one point = 1/72 inch).
If WIDTHS specifies a single value, it is applied to all selected columns.
WIDTHS cannot be used with DIMENSION=ROWS.

ROWLABELS selects labels in the row dimension.  It only accepts numbers;
-1 is the innermost label.  It is used only with
ROWLABELWIDTHS, which is a set of widths to apply to the row labels.

The STYLES subcommand sets styles for the text in each selected row or
column and/or applies a tableLook
TLOOK specifies a tablelook to be applied before other styling.

TEXTSTYLE sets the text style to
REGULAR, BOLD, ITALIC, or BOLDITALIC.

TEXTCOLOR sets the text color.  Color specifications are a triple of numbers
for the Red, Green, and Blue components.  Each value must be between 0 (none),
and 255 (maximum).  For example, 255 0 0 is red, and 255 255 255 is white.

BACKGROUNDCOLOR sets the cell background color.

Besides the built in functionality, you can create a Python function to be called for each qualifying
cell.  CUSTOMFUNCTION gives the module and function name in quotes and separated by ".".
See module customstylefunctions for documentation on how to write such functions.

Custom functions can have user-specified parameters written in Python notation.  For example,
"myfuncs.decorate(p1=100, p2='xyz')"
specifies parameters p1 and p2 with values 100 and 'xyz'.  Details on retrieving these values
within the function are found in customstylefunctions.py.

APPLYTO can be LABELS, DATACELLS, or BOTH and determines what the styles are
applied to.  For labels, the styles are applied according to the LEVEL specification, i.e.,
from the specified level inwards.

APPLYTO can also be an expression written in Python syntax that evaluates to True or False.  
With an expression, only data cells where the expression is true get the styles applied.
Use x (in lower case) to stand for the value of the cell when writing the expression and
i to stand for the row or column number in the opposite dimension from SELECT.
ii stands for the row or column number in the SELECT dimension.
Valid comparison operators are
==, !=, <=, <, >, >=.
For example, "x > 50000" would cause the styles to be applied only to cells with a 
value greater than 50000 within the selected rows or columns.
ii % 2 == 1
would be True for the odd-numbered rows or columns.
All cells in the row or column should be numeric.
Date formats cannot be used.
If the expression cannot be evaluated for a cell, it is considered False.

You can only operate on one dimension with a single command, but you can use as many
commands as needed.

/HELP displays this text and does nothing else.

Example:

LOGISTIC REGRESSION <syntax for logistic>.
SPSSINC MODIFY TABLES SUBTYPE='variables in the equation' SELECT = 2 3 'Upper'.

This causes the third, fourth, and the "Upper" columns to be 
hidden in the "variables in the equation" table of the logistic 
regression output.

DISCRIMINANT <syntax for discriminant>
SPSS MODIFY TABLES subtype='Group Statistics' 'Prior Probabilities for Groups' columns='Weighted'.
SPSS MODIFY TABLES SUBTYPE="Group Statistics", COLUMNS="Unweighted" 
/WIDTHS WIDTHS=50 ROWLABELS=-1 ROWLABELWIDTHS=40.

EXAMINE VARIABLES=accel engine foreign horse
  /PLOT NONE
  /PERCENTILES(5,10,25,50,75,90,95) HAVERAGE.
SPSS MODIFY TABLES SUBTYPE='Percentiles' LEVEL = -3
COLUMNS="Tukey's Hinges" DIMENSION=ROWS.

This hides the Tukey Hinges in the Percentiles table.  Note that this requires LEVEL=-3,
because there is a hidden level inbetween the two visible parts.

SPSS MODIFY TABLES subtype='variables in the equation' SELECT="B" "Sig."
/STYLES TEXTCOLOR = 0 0 255 BACKGROUNDCOLOR=0 255 0.

* bold cell standarized residuals larger than 2 in a Crosstab.
CROSSTABS  /TABLES=cylinder BY origin
  /CELLS=COUNT SRESID TOTAL.
SPSS MODIFY TABLES SUBTYPE='Crosstabulation' SELECT="Std. Residual"
DIMENSION = ROWS
/STYLES TEXTSTYLE=BOLD APPLYTO="abs(x) > 2".
"""
def Run(args):
    """Execute the MODIFY TABLES extension command"""

    args = args[list(args.keys())[0]]

    oobj = Syntax([
        Template("SUBTYPE", subc="",  ktype="str", var="subtype", islist=True),
        Template("PROCESS", subc="", ktype="str", var="process", islist=False),
        Template("SELECT", subc="",  ktype="literal", var="select", islist=True),
        Template("REGEXP", subc="", ktype="bool", var="regexp"),
        Template("DIMENSION", subc="", ktype="str", var="dimension"),
        Template("LEVEL", subc="", ktype="int", var= "level"),
        Template("HIDE", subc="", ktype="bool", var="hide", islist=False),
        Template("PRINTLABELS", subc="", ktype="bool", var="printlabels"),
        Template("COUNTINVIS", subc="", ktype="bool", var="countinvis"),
        Template("SIGCELLS", subc="", ktype="str", var="sigcells"),
        Template("SIGLEVELS", subc="", ktype="str", var="siglevels",
            vallist=["both", "upper", "lower"]),
        
        Template("WIDTHS", subc="WIDTHS", ktype="int", var="widths", vallist=(0,), islist=True),
        Template("ROWLABELS", subc="WIDTHS", ktype="str", var="rowlabels", islist=True),
        Template("ROWLABELWIDTHS", subc="WIDTHS", ktype="int", var="rowlabelwidths", islist=True),
        Template("TLOOK", subc="STYLES", ktype="literal", var="tlook"),
        Template("TEXTSTYLE", subc="STYLES", ktype="str", var="textstyle", islist=False),
        Template("TEXTCOLOR", subc="STYLES", ktype="int", var="textcolor", vallist=(0, 255), islist=True),
        Template("BACKGROUNDCOLOR", subc="STYLES", ktype="int", var="bgcolor", vallist=(0, 255), islist=True),
        Template("APPLYTO", subc="STYLES", ktype="literal", var="applyto", islist=False),
        Template("CUSTOMFUNCTION", subc="STYLES", ktype="literal", var="customfunction", islist=True),
        
        Template("HELP", subc="", ktype="bool")])

    # A HELP subcommand overrides all else
    if "HELP" in args:
        #print helptext
        helper()
    else:
        ###import cProfile
        ###cProfile.runctx("import extension, modifytables;extension.processcmd(oobj, args, modifytables.modify)",
            ###globals(), locals())
        processcmd(oobj, args, modifytables.modify)

def helper():
    """open html help in default browser window
    
    The location is computed from the current module name"""
    
    import webbrowser, os.path
    
    path = os.path.splitext(__file__)[0]
    helpspec = "file://" + path + os.path.sep + \
         "markdown.html"
    
    # webbrowser.open seems not to work well
    browser = webbrowser.get()
    if not browser.open_new(helpspec):
        print(("Help file not found:" + helpspec))
try:    #override
    from extension import helper
except:
    pass
