#Licensed Materials - Property of IBM
#IBM SPSS Products: Statistics General
#(c) Portions Copyright IBM Corp. 2009, 2020
#US Government Users Restricted Rights - Use, duplication or disclosure 
#restricted by GSA ADP Schedule Contract with IBM Corp.

"""This module contains some sample custom functions.  Custom functions, if specified, are
called for each label and.or data cell as specified in the STYLES syntax.  They are called after
any other styling functions specified.

It is better not to put your own custom functions in this module as they will be overwritten
if you install a newer version of the SPSSINC MODIFY TABLES command.

The custom function must have this signature if it does not have custom parameters.
f(obj, i, j, numrows, numcols, section, more)
where
obj is the table labels or datacells array currently being processed
i, j are the coordinates of the current label or data cell numbered from zero.
numrows and numcols provide the dimensions of the object
section is "labels" or "datacells"
more is an object with other items sometimes needed.
more.rowlabelarray
more.columnlabelarray
more.datacells
are the corresponding parts of the pivot table.
more.thetable is the pivot table object itself.  This might be used for apis such as ClearSelection that
require the table object.
Their main use is for doing something to a part of the table not being passed in the call.  See the Regression
coefficient example below.

A custom function would typically call a scripting pivot table cell function such as SetBackgroundColorAt(i,j),
but it can do anything it wants.

custom functions can also get custom parameters specified in the syntax.  See StripeOddRows2 
below for an example.
Custom parameters must be written in the syntax in Python function notation with parameter names.
E.g.,
CUSTOMFUNCTION="customstylefunctions.stripeOddRows2(r=20,g=20,b=200)".

If the custom function definition has an additional final parameter named custom,
that argument will be passed as an updatable dictionary of the parameters and values specified in
the syntax.
There is always a dictionary available with parameter "_first" that is initially set True.  The function can
use this to do initial processing and then set it to False.
Other parameters in the custom dictionary should be retrieved with the get function, 
since there is no prior checking that the function parameters are valid.

A custom function does not have to return a value, but if it returns False, all further
processing of the table will be skipped.  If the custom function modifies the structure
of the table, returning False may be necessary to avoid raising an exception.
If the function knows that nothing else needs to be done, returning False may speed things up.

Exceptions raised by custom functions are suppressed, and execution continues."""

__author__ = "SPSS, JKP"
__version__ = "1.9.2"

# 21-oct-2015 Add exception protection to SetNumericFormatAndDecimals
# 06-aug-2016 Add spreadsig function to move new style significance levels to their own row



# function RGB takes a list of three values and returns the RGB value
# function floatex decodes a numeric string value to its float value taking the cell format into account


import SpssClient   # for text constants
from modifytables import RGB
from extension import floatex  # strings to floats
import sys

#debugging (move this code appropriately for repeated debugging)
#import wingdbstub
#try:
    #import wingdbstub
    #if wingdbstub.debugger != None:
        #import time
        #wingdbstub.debugger.StopDebug()
        #time.sleep(2)
        #wingdbstub.debugger.StartDebug()
    #import thread
    #wingdbstub.debugger.SetDebugThreads({thread.get_ident(): 1}, default_policy=0)
    ## for V19 use
    ##    ###SpssClient._heartBeat(False)
#except:
    #pass




# generic function
# apply a specified scripting api to the appropriate object.
# This provides access to apis not already provided for in MODIFY TABLES
# functionality or other custom functions.

# example - rotating inner column labels for all Descriptives tables
#SPSSINC MODIFY TABLES subtype="'Descriptive Statistics'"
#  SELECT=1 DIMENSION= COLUMNS LEVEL = -1  PROCESS = ALL 
#  /STYLES  APPLYTO=LABELS 
#  CUSTOMFUNCTION="customstylefunctions.generic(item='thetable', func='SetRotateColumnLabels', onceonly=False, parms=[True])".

# example, change font size in selected data cells (can also be done with built-in MODIFY TABLES functionality)
#SPSSINC MODIFY TABLES subtype="'Descriptive Statistics'"
#SELECT=1 DIMENSION= COLUMNS LEVEL = -1  PROCESS = ALL 
#/STYLES  APPLYTO=DATACELLS 
#CUSTOMFUNCTION="customstylefunctions.generic(item='datacells', func='SetTextSizeAt', onceonly=False, parms=['i','j', 20])".

# Note that api and function names are quoted strings.


import copy
def generic(obj, i, j, numrows, numcols, section, more, custom):
    """Apply a scripting function to the appropriate part of a pivot table
    
    custom arguments:
    item - the name of one of the items listed in the more parameter as a string:
        "thetable", "rowlabelarray","columnlabelarray", "datacells",
    func - name of function to call as a string
    onceonly - True or False on whether to stop applying after the first time
    Note that onceonly applies to the whole command, not just the current object
    parms - parameters required by that particular function as a list or None if no parameters
    If strings 'i' or 'j' appear in the parms list (quoted) in either of the first two positions, 
    they are replaced by the current row and column numbers respectively
    exceptions in this function will generate error messages"""
    
    try:
        if custom["_first"]:
            if not "onceonly" in custom:   # for orderly error reporting
                print((_("""Error: Required parameter onceonly was not specified""")))
                return False
            custom["_first"] = False
        
        # get the function to apply from the current pivot table object or part thereof
        f = getattr(getattr(more, custom["item"]), custom["func"])
        if custom["parms"] is None:
            f()
        else:
            # substitute current positional parameters if referenced in parms
            parms2 = copy.copy(custom["parms"])
            currentloc = [i,j]
            for pos, z in enumerate(['i', 'j']):
                try:
                    loc = parms2.index(z, 0, 2)
                    parms2[loc] = currentloc[pos]
                except:
                    pass
            f(*parms2)
    except:
        print((_("""Error in custom function generic:"""), sys.exc_info()[1]))
    finally:
        if custom["onceonly"]:
            return False
        else:
            return
    
    
    
    

# example usage:

def stripeOddDataRows(obj, i, j, numrows, numcols, section, more):
    """Color background of odd number rows for data portion"""
    
    if section == "datacells" and i % 2 == 1:
        obj.SetBackgroundColorAt(i, j, RGB((0,0,200)))
        
def stripeOddRows(obj, i, j, numrows, numcols, section, more):
    """Color background of odd number rows for data and labels"""
    
    if i % 2 == 1:
        obj.SetBackgroundColorAt(i, j, RGB((0,0,200)))

# Example usage:
#SPSSINC MODIFY TABLES subtype="'Variables in the Equation'"
#SELECT="<<ALL>>" 
#DIMENSION= ROWS
#LEVEL = -1  PROCESS = PRECEDING 
#/STYLES  APPLYTO=DATACELLS 
#CUSTOMFUNCTION="customstylefunctions.stripeOddRows2(r=20,g=20,b=200)".

def stripeOddRows2(obj, i, j, numrows, numcols, section, more, custom):
    """stripe odd rows with color parameters
    
    extra parameters are r, g, b"""
    

    if i % 2 == 1:
        # retrieve three parameters with defaults and calculate color value first time then add to dictionary        
        if custom["_first"]:
            custom["_color"] = RGB((custom.get('r',0), custom.get('g', 0), custom.get('b', 200))) 
            custom["_first"] = False
        obj.SetBackgroundColorAt(i, j, custom["_color"])

# Example usage:
#SPSS MODIFY TABLES subtype='Variables in the Equation' select="<<ALL>>"
#DIMENSION=COLUMNS LEVEL=1
#/STYLES APPLYTO=DATACELLS CUSTOMFUNCTION="customstylefunctions.washColumnBackgrounds".

def washColumnBackgrounds(obj, i, j, numrows, numcols, section, more):
    """color cells backgrounds from dark  to light"""

    mincolor=180.
    maxcolor=255.
    increment = (maxcolor - mincolor)/(numcols-1)
    colorvalue = round(mincolor + increment * j)
    obj.SetBackgroundColorAt(i,j, RGB((colorvalue, colorvalue, colorvalue)))
    
def washColumnsBlue(obj, i, j, numrows, numcols, section, more):
    """color cells backgrounds from dark  to light"""

    mincolor=150.
    maxcolor=255.
    increment = (maxcolor - mincolor)/(numcols-1)
    colorvalue = round(mincolor + increment * j)
    obj.SetBackgroundColorAt(i,j, RGB((mincolor, mincolor, colorvalue)))
    
# Usage example:
# CUSTOMFUNCTION="customstylefunctions.washColumns(color='green')"
    
def washColumns(obj, i, j, numrows, numcols, section, more, custom):
    """color cells backgrounds from dark  to light
    
    parameter color = "red", "green", or "blue" determines the background color"""

    
    validcolors= ["red", "green", "blue"]
    dim = custom.get("color", "blue").lower()   # get user-specified color
    try:
        index = validcolors.index(dim)
    except:
        raise ValueError("Invalid color parameter for function washColumns: %s" % dim)

    mincolor=150.
    maxcolor=255.
    increment = (maxcolor - mincolor)/(numcols-1)
    colorvalue = round(mincolor + increment * j)
    color = [mincolor, mincolor, mincolor]
    color[index] = colorvalue
    obj.SetBackgroundColorAt(i,j, RGB(color))
    
# Next function might be invoked with syntax like this after running REGRESSION
#SPSS MODIFY TABLES subtype='Coefficients' select=0
#/styles applyto=datacells customfunction="customstylefunctions.makeSigCoefsBold".
# Note that the variable names or labels are column 3 in the row label array.    

def makeSigCoefsBold(obj, i, j, numrows, numcols, section, more):
    """Working on the REGRESSION Coefficients table, make each coefficient
    that is significant at the 5% level bold.
    Coefficients are column 0, and significance is column 4"""
    
    # a little misuse protection
    if section != 'datacells':
        return
    sig = floatex(obj.GetValueAt(i, 4), obj.GetNumericFormatAt(i, 4))
    if sig <= 0.05:
        obj.SetTextStyleAt(i,0, SpssClient.SpssTextStyleTypes.SpssTSBold)
        more.rowlabelarray.SetTextStyleAt(i, 3, SpssClient.SpssTextStyleTypes.SpssTSBold)

# color crosstab output cells red if the selected cell is greater than a specified value.
# E.g., if the cells consist of COUNT followed by ASRESID, it would color both.
# CROSS agecat BY pop  /CELLS=COUNT asresid.
#SPSSINC MODIFY TABLES subtype="'Crosstabulation'"
#SELECT="Adjusted Residual"  DIMENSION= ROWS LEVEL = -1  PROCESS = PRECEDING 
#/STYLES  CUSTOMFUNCTION="customstylefunctions.colorCrosstabResiduals(thresh=1.5)".
        
# Parameters can specify the threshold and the number of  rows to color in
# addition to the criterion row.

def colorCrosstabResiduals(obj, i, j, numrows, numcols, section, more,custom):
    if section == "datacells":
        try:
            if abs(floatex(obj.GetValueAt(i,j))) >= custom.get("thresh", 2.0):
                for k in range(custom.get("number", 1) + 1):
                    obj.SetBackgroundColorAt(i - k, j, RGB((255,0,0)))
                obj.SetBackgroundColorAt(i-1, j, RGB((255,0,0))) 
        except:
            pass

        
        
# The next function sets the selected cells to have two decimal places.
# Usage example:
# FREQUENCIES var.
#SPSSINC MODIFY TABLES subtype="frequencies"
#SELECT="Percent"  DIMENSION= COLUMNS LEVEL = -1  PROCESS = PRECEDING 
#/STYLES  CUSTOMFUNCTION="customstylefunctions.SetTwoDecimalPlaces".


def SetTwoDecimalPlaces(obj, i, j, numrows, numcols, section, more):
    "Set the cell format to two decimals"
    
    if section != 'datacells':
        return
    obj.SetHDecDigitsAt(i, j, 2)
    
# The next function is similar but takes an optional decimals parameter
# Usage example:
# FREQUENCIES var.
#SPSSINC MODIFY TABLES subtype="frequencies"
#SELECT="Percent"  DIMENSION= COLUMNS LEVEL = -1  PROCESS = PRECEDING 
#/STYLES  CUSTOMFUNCTION="customstylefunctions.SetDecimalPlaces(decimals=4)".

def SetDecimalPlaces(obj, i, j, numrows, numcols, section, more, custom):
    """Set the cell format to user-specified decimals
    
    parameter must be decimals=n"""
    
    if section != 'datacells':
        return
    decimals = custom.get("decimals", 2)
    obj.SetHDecDigitsAt(i, j, decimals)
    
def setLeadingZero(obj, i, j, numrows, numcols, section, more):
    """Set leading zero on cell by converting to string and prepending a zero
    and convert decimal symbol to comma
    Input is expected to be only cells with abs(x) < 1"""
    
    if section != 'datacells':
        return
    value = obj.GetUnformattedValueAt(i,j)
    decimals = obj.GetHDecDigitsAt(i,j)
    try:
        value2 = str(round(float(value), decimals)).replace(".", ",")  # comma locale
        
        obj.SetValueAt(i,j, chr(160) + value2)  # a NBSP is used to prevent conversion to a number
        obj.SetHAlignAt(i,j, SpssClient.SpssHAlignTypes.SpssHAlRight)
    except:
        pass

# The next function changes the selected cell values.  The full-precision value is
# changed to the value with decimals as displayed.  Thus the cell value in edit mode
# or in a full-precision export format such as Excel will be the same as the displayed
# value.
# roundToFormat takes no parameters.  Used via
# customfunction="customstylefunctions.roundToFormat" in SPSSINC MODIFY TABLES.
# Has no effect on date values or other values that cannot be converted to floats.
# It does work with DOLLAR and Custom Currency formats.
# Requires at least SPSS version 17.0.2

def roundToFormat(obj, i, j, numrows, numcols, section, more):
    """Set the actual cell value to the number of decimals in the format.
    This eliminates the undisplayed precision and can be used to
    force full-precision exports, e.g., for Excel, to show only the
    formatted values.  Sysmis and blank values and nonnumeric values
    are not affected."""
    
    if section != 'datacells':
        return
    value = obj.GetUnformattedValueAt(i,j)
    decimals = obj.GetHDecDigitsAt(i,j)
    try:
        value = round(float(value), decimals)
        obj.SetValueAt(i,j, str(value))
    except:
        pass


# The next function sets the format of a data cell.
# Usage example:
# FREQUENCIES var
# SPSSINC MODIFY TABLES subtype="frequencies"
# SELECT "Percent" "Valid Percent" "Cumulative Percent"
# STYLES CUSTOMFUNCTION='customstylefunctions.SetNumericFormat(format="##.#%")'
def SetNumericFormat(obj, i, j, numrows, numcols, section, more, custom):
    """Set numeric cell format.  Format definitions can be seen in the pivot table
    cell editor or Appendix C of the scripting manual
    
    custom parameter is format"""
    
    if section == "datacells":
        obj.SetNumericFormatAt(i, j, custom.get("format", "#.#"))
                               
#The next function is the same as SetNumericFormats except that it
# has a parameter for the number of decimals

def SetNumericFormatAndDecimals(obj, i, j, numrows, numcols, section, more, custom):
    """Set numeric cell format with decimals.  Format definitions can be seen in the pivot table
    cell editor or Appendix C of the scripting manual
    
    custom parameters:
    format= 'formatspec'   (default is "#.#"
    decimals=n                (default is 2)  """

    # If a pivot table has a short row such as with FREQUENCIES, an exception may
    # be raised on such a cell
    if section == "datacells":
        try:
            obj.SetNumericFormatAtWithDecimal(i, j, custom.get("format", "#.#"),
                custom.get("decimals", 2))
        except:
            pass
    
# The next function can be used to hide portions of tables where there is a sequence of numbered
# repetitions of blocks.  For example, running REGRESSION with stepwise methods or other blocks
# produces sets of equation estimates.  The last block is the final result.  The function below hides
# all the rows except for the last set.
    
# Usage
# REGRESSION
#  /DEPENDENT mpg
#  /METHOD=STEPWISE engine horse weight
#  /METHOD=ENTER year.
    
# SPSSINC MODIFY TABLES DIMENSION=ROWS PROCESS=PRECEDING
# /STYLES CUSTOMFUNCTION="customstylefunctions.hideNonfinalRows".

def hideNonfinalRows(obj, i, j, numrows, numcols, section, more):
    if section == "labels":
        try:
            lastouterlabel = obj.GetValueAt(numrows-1, 1)
            value = obj.GetValueAt(i, 1)
            if value != lastouterlabel:
                if value != more.previousUsedValue:
                    obj.HideLabelsWithDataAt(i,1)
                more.previousUsedValue = value
        except:
            print(("Pivot table exception.  row: %s" % i))

    
# Next function is intended for use with SPSSINC MERGE TABLES when the test table has been
# merged into a custom table, but it also has other uses.

AtoZ = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

def boldIfEndsWithAtoZLetter(obj, i, j, numrows, numcols, section, more):
    if section == 'datacells':
        v = obj.GetValueAt(i,j)
    if v[-1] in AtoZ:
        obj.SetTextStyleAt(i,j, SpssClient.SpssTextStyleTypes.SpssTSBold)
        
# color the background of cells if value ends with a letter code

# Usage example with a custom table after having merged significance
# using SPSSINC MERGE TABLES
#SPSSINC MODIFY TABLES subtype="customtable"
#SELECT="<<ALL>>" DIMENSION = COLUMNS PROCESS = PRECEDING 
#/STYLES APPLYTO=DATACELLS 
#CUSTOMFUNCTION="customstylefunctions.colorIfEndsWithAtoZLetter(r=255, g=255, b=50)".

def colorIfEndsWithAtoZLetter(obj, i, j, numrows, numcols, section, more, custom):
    """set the background color of any cell that ends with a letter code
    
    custom parameters are the red, green, and blue codes as r, g, b
    Default color is yellow (255,255,0)"""
    
    if custom["_first"]:
        custom["_color"] = RGB((custom.get('r',255), custom.get('g', 255), custom.get('b', 0))) 
        custom["_first"] = False  
    
    if section == 'datacells':
        v = obj.GetValueAt(i,j)
        if v[-1] in AtoZ:
            obj.SetBackgroundColorAt(i, j, custom["_color"])  


def SetCellMargins(obj, i, j, numrows, numcols, section, more, custom):
    """Set cell margins as an integer multiple of current margins
    
    custom parameters are left, right, top, bottom"""
    
    marginApis = {'top':(obj.GetTopMarginAt, obj.SetTopMarginAt),
    'left':(obj.GetLeftMarginAt, obj.SetLeftMarginAt),
    'right':(obj.GetRightMarginAt, obj.SetRightMarginAt),
    'bottom':(obj.GetBottomMarginAt, obj.SetBottomMarginAt)}
    
    if custom["_first"]:
        custom["_first"] = False
        custom["mlist"] = dict([(k, custom[k])\
            for k in ['left','right','top','bottom'] if k in custom])
        
    for k, val in list(custom["mlist"].items()):
        curvalue = marginApis[k][0](i,j)
        marginApis[k][1](i,j, curvalue * int(val))


# Sort the rows of a simple pivot table
# example usage:
#get file="c:/spss18/samples/english/employee data.sav".
#correlations salary salbegin prevexp with educ.
#dataset name emp.

#*table must be pivoted before sorting (v18.0.1 could have a timing problem here).
#SPSSINC MODIFY TABLES subtype="'Correlations'"
#SELECT="Pearson Correlation" DIMENSION= ROWS
#/STYLES  APPLYTO=DATACELLS 
#CUSTOMFUNCTION="customstylefunctions.moveRowsToColumns(fromdim=0, todim=0)".
#SPSSINC MODIFY TABLES subtype="'Correlations'"
#SELECT="Pearson Correlation" DIMENSION= COLUMNS
#/STYLES  APPLYTO=DATACELLS 
#CUSTOMFUNCTION="customstylefunctions.sortTable".

def sortTable(obj, i, j, numrows, numcols, section, more, custom):
    """Sort the rows of the table according to the selected column values
    Cell formats are NOT updated, so the formats for all cells in a column
    should be the same.
    
    Since it is not practical to move table footnotes along with the cell values,
    all footnotes are hidden.
    
    custom parameter is direction ('a', the default, or 'd')"""
    
    if not section == "datacells":
        return
    
    direction = custom.get("direction", "a")
    if not direction in ['a', 'd']:
        print("direction must be 'a' or 'd'")
        raise ValueError
    
    PvtMgr = more.thetable.PivotManager()
    numrowdims = PvtMgr.GetNumRowDimensions()
    if numrowdims != 1:
        print("Cannot sort table unless there is exactly one row dimension")
        raise ValueError

    col = j   # sorting column
    colvalues = []

    # store each row of pivot table as a list item consisting of the row label, the unformatted cell value
    # as a number if possible followed by the formatted string values
    for i in range(numrows):
        cval = []
        cval.append(more.rowlabelarray.GetValueAt(i,1))   # row label
        # append sorting key after trying to convert to a number
        kv = obj.GetUnformattedValueAt(i, col)
        try:
            kv = float(kv)
        except:
            pass
        cval.append(kv)
        
        for j in range(numcols):
            cval.append(obj.GetUnformattedValueAt(i, j))
            nf = obj.GetNumericFormatAt(i,j)
            if nf == '.-.$... ':
                nf = "$#,###.##"
            cval.append(nf)
        colvalues.append(cval)
    colvalues.sort(key=lambda arow: arow[1], reverse = direction == "d")

    for i in range(numrows):
        more.rowlabelarray.SetValueAt(i, 1, colvalues[i][0])
        
    for i in range(numrows):
        for j in range(0, numcols):
            try:
                index = j * 2 + 2
                obj.SetValueAt(i,j, colvalues[i][index])
                # for versions prior to 19, not setting the format string
                #obj.SetNumericFormatAt(i, j, colvalues[i][index + 1])
            except:
                pass
    # hide all footnotes
    more.thetable.SelectAllFootnotes()
    more.thetable.HideFootnote()
    return False

# Move a row in front of another row.
# Usage example
# CROSSTABS  /TABLES=year BY origin  /CELLS=COUNT ROW.
# SPSSINC MODIFY TABLES subtype="'Crosstabulation'" SELECT=0 DIMENSION= ROWS
# LEVEL = -1  PROCESS = PRECEDING 
# /STYLES  APPLYTO=LABELS CUSTOMFUNCTION="customstylefunctions.moveRowDimension(fromr=1, tor=0)".

def moveRowDimension(obj, i, j, numrows, numcols, section, more, custom):
    """Move a row in front of another row.
    Note that this will change the nesting hierarchy
    parameters:
    from: row number to move from
    to: row number to move to
    Default action interchanges the first two columns of the row labels
    
    Warning: once this function has changed the structure, it would be unwise
    to do anything else to the table on the same invocation of MODIFY TABLES!
    It returns False if it does a move in order to suppress further processing.
    
    Note: some tables do not have the dimensions you might think.  CTABLES tables, for example,
    generally are not intended to be pivoted and do not have separate dimensions that you
    would expect from their appearance.
    """
    
    # dimensions are numbered in the api in reverse: the innermost dimension is 0!
    # parameters are reflected in order to allow the natural usage


    PvtMgr = more.thetable.PivotManager()
    numrowdims = PvtMgr.GetNumRowDimensions() - 1
    fromr = numrowdims - custom.get("fromr", 1)
    tor = numrowdims - custom.get("tor", 0)
    try:
        row = PvtMgr.GetRowDimension(fromr)
        row.MoveToRow(tor)
        return False
    except:
        print("Row move failed: probable invalid row number")
        
# transpose the rows and columns of a table
# Usage example
# MEANS TABLES=engine BY origin  /CELLS MEAN COUNT STDDEV.
# SPSSINC MODIFY TABLES subtype="'Report'" SELECT=0 DIMENSION= ROWS
# /STYLES  APPLYTO=LABELS CUSTOMFUNCTION="customstylefunctions.transpose".

def transpose(obj, i, j, numrows, numcols, section, more, custom):
    """Transpose the rows and columns of a table
    
    Warning: once this function has changed the structure, it would be unwise
    to do anything else to the table on the same invocation of MODIFY TABLES!
    Therefore, it returns False in order to stop any further processing.
    """
    # this api raises a harmless exception on some tables after completing
    # the transpose
    try:
        PvtMgr = more.thetable.PivotManager().TransposeRowsWithColumns() 
    except:
        pass
    return False
    
# Move the layers of a table to the columns
# Usage example
# OLAP CUBES engine horse BY origin BY year
#  /CELLS=SUM COUNT MEAN STDDEV SPCT NPCT.
# SPSSINC MODIFY TABLES subtype="'LayeredReports'" SELECT=0 DIMENSION= ROWS
# /STYLES  APPLYTO=LABELS CUSTOMFUNCTION="customstylefunctions.moveLayersToColumns".
    
# Move the first layer dimension before the first column dimension.  The fromdim parameter is the
# dimension number in the layer (numbered from zero).  That dimension is moved before the todim
# dimension number in the columns.

# SPSSINC MODIFY TABLES subtype="'LayeredReports'" SELECT=0 DIMENSION= ROWS
# /STYLES  APPLYTO=LABELS CUSTOMFUNCTION="customstylefunctions.moveLayersToColumns(fromdim=0, todim=0)".

def moveLayersToColumns(obj, i, j, numrows, numcols, section, more, custom):
    """Move specified or all the dimensions in the layer to the columns
    
    Warning: once this function has changed the structure, it would be unwise
    to do anything else to the table on the same invocation of MODIFY TABLES!
    Therefore, it returns False in order to stop any further processing.
    """

    PvtMgr = more.thetable.PivotManager()
    fromdim = custom.get("fromdim")
    if fromdim is None:
        PvtMgr.MoveLayersToColumns()
    else:
        todim = custom.get("todim")
        if not todim is None:
            fd = PvtMgr.GetLayerDimension(fromdim)
            fd.MoveToColumn(todim)
    return False

# Move the layers to the rows
# Usage example
# OLAP CUBES engine horse BY origin BY year
#  /CELLS=SUM COUNT MEAN STDDEV SPCT NPCT.
# SPSSINC MODIFY TABLES subtype="'LayeredReports'" SELECT=0 DIMENSION= ROWS
# /STYLES  APPLYTO=LABELS CUSTOMFUNCTION="customstylefunctions.moveLayersToRows".

# Move the first layer dimension before the first row dimension.  The fromdim parameter is the
# dimension number in the layer (numbered from zero).  That dimension is moved before the todim
# dimension number in the rows.
# SPSSINC MODIFY TABLES subtype="'LayeredReports'" SELECT=0 DIMENSION= ROWS
# /STYLES  APPLYTO=LABELS CUSTOMFUNCTION="customstylefunctions.moveLayersToRows(fromdim=0, todim=0)".

def moveLayersToRows(obj, i, j, numrows, numcols, section, more, custom):
    """Move specified or all layer dimensions to the rows of the table.
    
    Warning: once this function has changed the structure, it would be unwise
    to do anything else to the table on the same invocation of MODIFY TABLES!
    Therefore, it returns False in order to stop any further processing.
    """
    PvtMgr = more.thetable.PivotManager()
    fromdim = custom.get("fromdim")
    if fromdim is None:
        PvtMgr.MoveLayersToRows()
    else:
        todim = custom.get("todim")
        if not todim is None:
            fd = PvtMgr.GetLayerDimension(fromdim)
            fd.MoveToRow(todim)
    return False

# If using this custom function without others, the SPSSINC MODIFY TABLES
# selection should just specify a single cell

# The following two functions work similarly to the preceding ones except that the
# fromdim and todim parameters are required.

def moveColumnsToLayers(obj, i, j, numrows, numcols, section, more, custom):
    """Move specified dimension in the columns to the layer.
    
    Warning: once this function has changed the structure, it would be unwise
    to do anything else to the table on the same invocation of MODIFY TABLES!
    Therefore, it returns False in order to stop any further processing.
    """

    PvtMgr = more.thetable.PivotManager()
    fromdim = custom.get("fromdim")
    todim = custom.get("todim")
    if fromdim is None or todim is None:
        raise ValueError("Error: fromdim or todim was not specified")

    fd = PvtMgr.GetColumnDimension(fromdim)
    fd.MoveToLayer(todim)
    return False

def moveRowsToLayers(obj, i, j, numrows, numcols, section, more, custom):
    """Move specified dimension in the rows to the layer.
    
    Warning: once this function has changed the structure, it would be unwise
    to do anything else to the table on the same invocation of MODIFY TABLES!
    Therefore, it returns False in order to stop any further processing.
    """

    PvtMgr = more.thetable.PivotManager()
    fromdim = custom.get("fromdim")
    todim = custom.get("todim")
    if fromdim is None or todim is None:
        raise ValueError("Error: fromdim or todim was not specified")

    fd = PvtMgr.GetRowDimension(fromdim)
    fd.MoveToLayer(todim)
    return False
    
# Example:
#SPSSINC MODIFY TABLES subtype="'Descriptive Statistics'"
#SELECT=0 
#DIMENSION= COLUMNS
#LEVEL = -1  PROCESS = PRECEDING 
#/STYLES  APPLYTO=DATACELLS 
#CUSTOMFUNCTION="customstylefunctions.hideFootnotes".
    
# To hide specific footnotes, specify customfunction like this:
#CUSTOMFUNCTION="customstylefunctions.hideFootnotes(fnlist=[1,3])".

# The following function moves a specified dimension from the columns to the rows
# Usage example:
#MEANS TABLES=salary BY jobcat BY gender  /CELLS MEAN COUNT.
# creates a table with statistics in the rows and jobcat and gender dimensions in the columns
#SPSSINC MODIFY TABLES subtype="'Report'"
#SELECT=0 
#DIMENSION= COLUMNS
#LEVEL = -1  PROCESS = PRECEDING 
#/STYLES  APPLYTO=DATACELLS 
#CUSTOMFUNCTION="customstylefunctions.moveColumnsToRows(fromdim=0,todim=1)".
# moves the GENDER column dimension to be the outermost row dimension.

def moveColumnsToRows(obj, i, j, numrows, numcols, section, more, custom):
    """Move specified dimension in the columns to the rows.
    
    Warning: once this function has changed the structure, it would be unwise
    to do anything else to the table on the same invocation of MODIFY TABLES!
    Therefore, it returns False in order to stop any further processing.
    """

    PvtMgr = more.thetable.PivotManager()
    fromdim = custom.get("fromdim")
    todim = custom.get("todim")
    if fromdim is None or todim is None:
        raise ValueError("Error: fromdim or todim was not specified")

    fd = PvtMgr.GetColumnDimension(fromdim)
    fd.MoveToRow(todim)
    return False

def moveRowsToColumns(obj, i, j, numrows, numcols, section, more, custom):
    """Move specified dimension in the rows to the columns.
    
    Warning: once this function has changed the structure, it would be unwise
    to do anything else to the table on the same invocation of MODIFY TABLES!
    Therefore, it returns False in order to stop any further processing.
    """

    PvtMgr = more.thetable.PivotManager()
    fromdim = custom.get("fromdim")
    todim = custom.get("todim")
    if fromdim is None or todim is None:
        raise ValueError("Error: fromdim or todim was not specified")

    fd = PvtMgr.GetRowDimension(fromdim)
    fd.MoveToColumn(todim)
    return False
    
def hideFootnotes(obj, i, j, numrows, numcols, section, more, custom):
    """Hide the footnotes in the pivot table.
    
    if the parameter fnlist is specified, just those sequence numbers are hidden.
    Otherwise all footnotes are hidden."""

    # Use the footnote array and make null markers
    fna = more.thetable.FootnotesArray()
    count = fna.GetCount()
    fnlist = custom.get("fnlist", list(range(count)))
    for fn in fnlist:
        if fn < count:
            fna.SetTextHiddenAt(fn, True)
            fna.ChangeMarkerToSpecial(fn, "")

def hideAllFootnotes(obj, i, j, numrows, numcols, section, more, custom):
    """Hide all footnotes in the table"""
    
    # Returns False to prevent being called multiple times
    # This function is much more efficient than using hideFootnotes
    more.thetable.SelectAllFootnotes()
    more.thetable.HideFootnote()
    return False

def blankTableTriangle(obj, i, j, numrows, numcols, section, more, custom):
    """Move specified dimension in the rows to the columns.
    
    custom parameter is triangle="uppper" or "lower".  upper is the default.
    upper means to blank the upper triangle.
    Blanking occurs according to whichever dimension is smaller:
    if there are more rows than columns, it is assumed that this is due
    to having multiple statistics in the rows and vice versa.
    Operation is only performed once regardless of the cell selection.
    PROCESS=ALL does not work with this function due to repetition protection
    """
    if not custom["_first"]:
        return
    custom["_first"] = False
    
    triangle = custom.get("triangle", "upper")
    rowstep = max(numrows/numcols,1)
    colstep = max(numcols/numrows,1)

    if numrows >= numcols:
        if triangle == "upper":    # blank upper triange
            for j in range(numcols):   # yes, numcols
                for i in range(0, j*rowstep):
                    _zap(more.datacells, i, j)
        else:   # blank lower triange
            for j in range(numcols):   # yes, numcols
                for i in range((j+1) * rowstep, numrows):
                    _zap(more.datacells, i, j)
    else:
        if triangle == "upper":
            for i in range(numrows):   # yes, numrows
                for j in range((i+1)*colstep, numcols):
                    _zap(more.datacells, i, j)
        else:
            for i in range(numrows):   # yes, numcols
                for j in range(0, i*colstep):
                    _zap(more.datacells, i, j)

def _zap(cells, i, j):
    cells.SetValueAt(i,j,"")
    try:
        cells.HideFootnotesAt(i,j)
    except:
        pass
    
    
# Hide rows if values are all below a threshold
# Usage example
#CTABLES
  #/TABLE educ BY gender > jobcat [COLPCT.COUNT PCT40.1]
  #/CATEGORIES VARIABLES=educ ORDER=A KEY=VALUE EMPTY=INCLUDE TOTAL=YES POSITION=AFTER 
    #MISSING=EXCLUDE
  #/CATEGORIES VARIABLES=gender jobcat ORDER=A KEY=VALUE EMPTY=INCLUDE MISSING=EXCLUDE.

#SPSSINC MODIFY TABLES subtype="'Custom Table'"
#SELECT="<<ALL>>" 
#DIMENSION= ROWS
#LEVEL = -1  PROCESS = ALL 
#/STYLES  APPLYTO=LABELS 
#CUSTOMFUNCTION="customstylefunctions.HideRowBasedOnValues(threshold=5)".

# APPLYTO must be LABELS
# All columns except optionally the first or last will be tested

def HideRowBasedOnValues(obj, i, j, numrows, numcols, section, more, custom):
    """Check whether all row values that are not missing are <= a threshold and hide
    
    custom parameters are
    threshold - test value for hiding: no default value
    By default, all column values are checked.  Specify
    omitfirst=n and/or omitlast=n where n is the number of columns to omit
    when testing values.  This accomodates preceding or following totals
    (at the highest level only).
    
    All row values except possibly the first or last are checked.  If any exceeds the
    threshold the row is not modified; if none exceed the threshold, the entire
    row is hidden.
    
    The command specification should reference the labels and applies to the innermost label only
    
    Requires at least version  Statistics version17.0.2"""
    
    
    if section == 'labels':
        thresh = float(custom.get("threshold", -1e8))
        numcols = more.datacells.GetNumColumns()
        omitfirst = custom.get("omitfirst", 0)
        omitlast = custom.get("omitlast", 0)
        
        start, stop = 0+omitfirst, numcols-omitlast
        
        for col in range(start, stop):
            v = more.datacells.GetUnformattedValueAt(i, col)
            try:
                if float(v) > thresh:
                    break
            except:
                pass
        else:
            obj.HideLabelsWithDataAt(i,j)
            
# The next function takes the first, outermost, row label and makes it the table title.
# Usage example:
# CTABLES
#     /TABLE jobcat [COUNT]  BY gender
#     /TITLES TITLE='any title'.
# 
# SPSSINC MODIFY TABLES subtype='Custom Table'
# SELECT="<<ALL>>" 
# /STYLES  APPLYTO=LABELS 
# CUSTOMFUNCTION="customstylefunctions.SetTitleFromStub".
            
def SetTitleFromStub(obj, i, j, numrows, numcols, section, more, custom):
    """Copy the first row label outermost element and set as table title
    
    The target table may need to have a title, which will be replaced."""
    
    rowtext = more.rowlabelarray.GetValueAt(0,1)
    more.pt.SetTitleText(rowtext)
    more.rowlabelarray.SetValueAt(0,1,"")
    more.rowlabelarray.SetRowLabelWidthAt(0, 1, 0)
    return False
        
# The following set of two color schemes is governed by the following license:
#Apache-Style Software License for ColorBrewer software and ColorBrewer Color Schemes

#Copyright (c) 2002 Cynthia Brewer, Mark Harrower, and The Pennsylvania State University.

#Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except in compliance with the License.
#You may obtain a copy of the License at

#http://www.apache.org/licenses/LICENSE-2.0

#Unless required by applicable law or agreed to in writing, software distributed
#under the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR
#CONDITIONS OF ANY KIND, either express or implied. See the License for the
#specific language governing permissions and limitations under the License.

#This text from my earlier Apache License Version 1.1 also remains in place for guidance on attribution and permissions:
#Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:
#1. Redistributions as source code must retain the above copyright notice, this list of conditions and the following disclaimer.
#2. The end-user documentation included with the redistribution, if any, must include the following acknowledgment:
#"This product includes color specifications and designs developed by Cynthia Brewer (http://colorbrewer.org/)."
#Alternately, this acknowledgment may appear in the software itself, if and wherever such third-party acknowledgments normally appear.
#4. The name "ColorBrewer" must not be used to endorse or promote products derived from this software without prior written permission. For written permission, please contact Cynthia Brewer at cbrewer@psu.edu.
#5. Products derived from this software may not be called "ColorBrewer", nor may "ColorBrewer" appear in their name, without prior written permission of Cynthia Brewer.

Set3Qualitative=[
   (141,211,199),
(255,255,	179),
(190,186,	218),
(251,128,	114),
(128,177,	211),
(253,180,	98),
(179,222,	105),
(252,205,	229),
(217,217,	217),
(188,128,	189),
(204,235,	197),
(255,237,	111)]

PastelQualitative=[
    (251,180,174),
(179,205,227),
(204,235,197),
(222,203,228),
(254,217,166),
(255,255,204),
(229,216,189),
(253,218,236),
(242,242,242)]


def qualitative(obj, i, j, numrows, numcols, section, more):
    """color cells backgrounds using qualitative scheme
    good for up to 12 columns."""

    obj.SetBackgroundColorAt(i,j, RGB(Set3Qualitative[min(j, 11)]))

def pastelqualitative(obj, i, j, numrows, numcols, section, more):
    """color cells backgrounds using Pastel Qualitative scheme
    good for up to 9 columns."""

    obj.SetBackgroundColorAt(i,j, RGB(PastelQualitative[min(j, 8)]))
    
def stripeOddDataRowsAndAlign(obj, i, j, numrows, numcols, section, more):
    """Color background of odd number rows for data portion"""
    
    if section == "datacells" and i % 2 == 1:
        obj.SetBackgroundColorAt(i, j, RGB((0,0,200)))
        obj.SetHAlignAt(i,j, SpssClient.SpssHAlignTypes.SpssHAlLeft)
        
    
# Hide all rows where all the data cells appear to be blank
def hideBlankRow(obj, i, j, numrows, numcols, section, more):
    """hide rows that appear entirely blank"""
    
    if not section =="datacells":
        return
    for c in range(numcols):
        if obj.GetValueAt(i,c) != "":
            break
    else:
        innerlabelcolumn = more.rowlabelarray.GetNumColumns() -1
        more.rowlabelarray.HideLabelsWithDataAt(i, innerlabelcolumn)
        

# Set horizontal alignment.  Usage example:
#SPSSINC MODIFY TABLES subtype="'Custom Table'"
#SELECT=-1
#DIMENSION= COLUMNS
#/STYLES  APPLYTO=DATA 
#CUSTOMFUNCTION="customstylefunctions.SetAlignment(align='center')".

alignparms= {"left": SpssClient.SpssHAlignTypes.SpssHAlLeft,
    "center" : SpssClient.SpssHAlignTypes.SpssHAlCenter,
    "right" : SpssClient.SpssHAlignTypes.SpssHAlRight,
    "mixed" : SpssClient.SpssHAlignTypes.SpssHAlMixed,
    "decimal" : SpssClient.SpssHAlignTypes.SpssHAlDecimal}

def SetAlignment(obj, i, j, numrows, numcols, section, more, custom):
    """set cell horizontal alignment
    
    parameter is align
    possible values are "left", "center", "right", "mixed", and "decimal"
    Default value is "right"
    """
    
    if custom["_first"]:
        custom["_first"] = False
        align = custom.get("align", "right")
        custom["align"] = alignparms[align]   # invalid value will raise exception
        
    obj.SetHAlignAt(i,j, custom["align"])
    
letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

# Usage: This function is useful for a Custom Table where the significance
# table has been merged into the main table. It replace occurrences of
# (letter) in the label or letter in the data cells with the specified substitution
# character

#SPSSINC MODIFY TABLES subtype="'Custom Table'"
#SELECT="<<ALL>>" DIMENSION= COLUMNS
#LEVEL = -1  PROCESS = PRECEDING 
#/STYLES  APPLYTO=BOTH 
#CUSTOMFUNCTION="customstylefunctions.reletter(letters='xyzw')".

#Alternatively, 
#CUSTOMFUNCTION="customstylefunctions.reletter(letters='red,white,blue,green')".

def reletter(obj, i, j, numrows, numcols, section, more, custom):
    """change occurrences of (A), (B(), etc to new set of symbols
    
    parameter letters is a sequence of characters to replace the corresponding
    letters A,B,... when they occur in a label in parentheses at the end, e.g., (B)
    If letters contains any comma characters, it is assumed that each comma-separated
    group (including any blanks) is the replacement for a single letter.  
    By using commas, you can generate multi-character sequences."""
    
    if custom["_first"]:
        custom["_first"] = False
        newletters = custom.get("letters", "")
        newletters = newletters.split(",")
        if len(newletters) == 1:
            newletters = newletters[0]    # no comma separators
        lennew = len(newletters)
        custom["map"] = dict([(a, b) for a, b in zip(letters[:lennew], newletters)])
        
    val = obj.GetValueAt(i,j)
    try:
        if section == "labels":
            if val[-1] == ")" and val[-3] == "(":
                val = val[:-2] + custom["map"][val[-2]] + val[-1]
                obj.SetValueAt(i, j, val)
        else:
            if isinstance(val, str):
                newval = []
                for c in val:
                    newval.append(custom["map"].get(c,c))
                obj.SetValueAt(i,j, "".join(newval))
    except:
        pass
    
def showcorner(obj, i, j, numrows, numcols, section, more, custom):
    obj.ShowHiddenDimensionLabelAt(i,j)

# Example usage:
#SPSSINC MODIFY TABLES subtype="customtable"
#SELECT="<<ALL>>" 
#DIMENSION= COLUMNS LEVEL = -1  SIGLEVELS=BOTH 
#PROCESS = PRECEDING 
#/STYLES  APPLYTO=DATACELLS 
#CUSTOMFUNCTION="customstylefunctions.spreadsig".

def spreadsig(obj, i, j, numrows, numcols, section, more, custom):
    """Move significance markers (new type) to their own row
    
    This function requires V24 or later
    It signals the caller to stop processing when it returns
    """
    
    labelcols = more.rowlabelarray.GetNumColumns() - 1  # off by 1
    row = numrows
    fail = False
    first = True
    while row > 0:
        row -= 1
        for col in range(numcols):
            markers = more.datacells.GetSigMarkersAt(row, col)
            ###print row, col, markers, more.datacells.GetValueAt(row, col)
            if markers:
                ###print row, col, markers
                break
        else:
            continue
        try:
            try:
                #more.rowlabelarray.InsertNewBefore(row, labelcols - 1, more.rowlabelarray.GetValueAt(row, labelcols-1))
                more.rowlabelarray.InsertNewBefore(row, labelcols, more.rowlabelarray.GetValueAt(row, labelcols))
            except:
                if first:
                    first = False
                    labelcols -= 1
                    more.rowlabelarray.InsertNewBefore(row, labelcols, more.rowlabelarray.GetValueAt(row, labelcols))
            for col in range(numcols):
                fmt = more.datacells.GetNumericFormatAt(row+1, col)
                decdigits = more.datacells.GetHDecDigitsAt(row+1, col)
                more.datacells.SetValueAt(row, col, str(more.datacells.GetUnformattedValueAt(row+1, col)))
                more.datacells.SetNumericFormatAt(row, col, fmt)
                more.datacells.SetHDecDigitsAt(row, col, decdigits)
                more.datacells.SetValueAt(row+1, col, "")
                more.datacells.SetVAlignAt(row, col, SpssClient.SpssVAlignTypes.SpssVAlTop)
            #more.rowlabelarray.SetValueAt(row + 1, labelcols - 1, "Significance")
            more.rowlabelarray.SetValueAt(row + 1, labelcols, "Significance")
        except:
            fail = True
            pass    #subtotals cause an exception in the pivot table
    if fail:
        print("Unable to process this table correctly to move significance levels to their own row")

    return False

        
            