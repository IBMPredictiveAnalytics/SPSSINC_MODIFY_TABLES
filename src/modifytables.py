#/***********************************************************************
# * Licensed Materials - Property of IBM 
# *
# * IBM SPSS Products: Statistics Common
# *
# * (C) Copyright IBM Corp. 1989, 2015
# *
# * US Government Users Restricted Rights - Use, duplication or disclosure
# * restricted by GSA ADP Schedule Contract with IBM Corp. 
# ************************************************************************/

# format specific rows or columns in a pivot table of given type
from __future__ import with_statement

__version__ = '1.4.2'
__author__ = "SPSS, JKP"

# Note: This module requires at least SPSS 17.0.0

# history
# 25-jun-2008 original version
# 30-jul-2008  add row hiding support
# 04-aug-2008 add column sizing and formatting features, rename command, implement wout scriptex
# 29-oct-2008  strip any matching outer quotes from any specified subtype
# 23-dec-2008 add PRINTLABELS keyword
# 27-dec-2008 add regular expression support for select
# 28-dec-2008 performance improvement with multiple styles
# 11-jan-2009  Allow custom functions with parameters.
# 05-may-2009 Make other index available as "ii" to applyto boolean
# 10-jul-2009 Support early termination from a custom function
# 12-jul-2009 Use deferred retrieval for pivot table parts that might not be used
# 23-jun-2010 translatable proc name
# 20-dec-2010 printtablelabels fix
# 03-feb-2014 Allow hiding based on spec not the innermost label
# 19-feb-2015 Allow testing on nonnumeric values in table

import SpssClient
from extension import floatex, _isseq
import re, functools, inspect, locale, sys
    
    
class fDataCellArray(object):
    def __init__(self, pt):
        self.pt = pt
        self.item = None
    def __getattr__(self, name):
        if self.item is None:
            self.item = self.pt.DataCellArray()
        return getattr(self.item, name)
    
class fRowLabelArray(object):
    def __init__(self, pt):
        self.pt = pt
        self.item = None
    def __getattr__(self, name):
        if self.item is None:
            self.item = self.pt.RowLabelArray()
        return getattr(self.item, name)

class fColumnLabelArray(object):
    def __init__(self, pt):
        self.pt = pt
        self.item = None
    def __getattr__(self, name):
        if self.item is None:
            self.item = self.pt.ColumnLabelArray()
        return getattr(self.item, name)


CUSTOMPARAMS={}
def modify(subtype, select=None,  skiplog=True, process="preceding", dimension='columns',
           level=-1, hide=False, widths=None, rowlabels=None, rowlabelwidths=None,
           textstyle=None, textcolor=None, bgcolor=None, applyto="both", customfunction=None, 
           printlabels=False, regexp=False, tlook=None, countinvis=True):
    """Apply a hide or show action to specified columns or rows of the specified subtype or resize columns

    subtype is the OMS subtype of the tables to process or a sequence of subtypes
    select is a sequence of column or row identifiers or ["ALL"].  Identifiers can be
    positive or negative numbers, counting from 0.  Negative numbers count from
    the end: -1 is the last row or column.
    Identifiers can also be the text of the lowest level in the column heading.
    If the value is or can be converted to an integer, it is assumed to be a column number.
    Numeric values are truncated to integers.
    You cannot hide all the items even though this routine will try.
    process specifies "preceding" to process the output of the preceding command or "all"
    to process all tables having any of the specified subtypes.
    level defaults to the innermost level (-1).  Specify a more negative number to move out or up in
    the label array.  -2, for example, would be the next-to-innermost level.
    When counting columns or rows, count at the innermost level regardless of the level setting.
    Hide cannot be combined with other actions.
    if skiplog is True, if the last item in the Viewer is a log, the search starts with the preceding item.
    It needs to be True if this function is run from the extension command and commands are echoing to the log.
    dimension == 'columns' indicates that columns should be operated on.  dimension == 'rows' specifies rows.
    widths specifies the width or widths to be applied to the selected rows or columns
    If it is a single element, it is used for all specified columns.
    Otherwise, it is a sequence of sizes in points of the same length as the select list.
    rowlabels and rowlabelwidths can be specified to set stub (row) widths.  rowlabels can only contain numbers.
    textstyle, textcolor, and bgcolor apply formatting.  colors are specified as three integers for RGB.
    textstyle can be REGULAR, BOLD, ITALIC, or BOLDITALIC
    APPLYTO can be BOTH, DATACELLS, LABELS, or a Boolean Python expression in which x stands for
    the cell value.

    customfunction is a list of module.function names of  functions to be called as cells are styled.

    This function processes the latest item in the designated Viewer: all pivot tables for that instance of
    the procedure are processed according to the subtype specification.
"""
    # ensure localization function is defined
    # pot file must be named SPSSINC_MODIFY_TABLES.pot
    global _
    try:
        _("---")
    except:
        def _(msg):
            return msg
        ###debugging
    #try:
        #import wingdbstub
        #if wingdbstub.debugger != None:
            #import time
            #wingdbstub.debugger.StopDebug()
            #time.sleep(2)
            #wingdbstub.debugger.StartDebug()
        #import thread
        #wingdbstub.debugger.SetDebugThreads({thread.get_ident(): 1}, default_policy=0)
    #except:
        #pass
    SpssClient.StartClient()
    try:
        info = NonProcPivotTable("INFORMATION", tabletitle=_("Information"))
        c = PtColumns(select, dimension, level, hide, widths, 
            rowlabels, rowlabelwidths, textstyle, textcolor, bgcolor, applyto, customfunction, 
            printlabels,regexp, tlook)
        if not _isseq(subtype):
            subtype=[subtype]
        # remove white space
        subtype = ["".join(st.lower().split()) for st in subtype]
        # remove matching outer quotes of any type
        subtype = [re.sub(r"""^('|")(.*)\1$""", r"""\2""", st) for st in subtype]
        if "*" in subtype:
            subtype = ["*"]
        items = SpssClient.GetDesignatedOutputDoc().GetOutputItems()
        itemcount = items.Size()
        if skiplog and items.GetItemAt(itemcount-1).GetType() == SpssClient.OutputItemType.LOG:
            itemcount -= 1
        for itemnumber in range(itemcount-1, -1, -1):
            item = items.GetItemAt(itemnumber)
            if process == "preceding" and item.GetTreeLevel() <= 1:
                break
            if item.GetType() == SpssClient.OutputItemType.PIVOT and\
               (subtype[0] == "*" or "".join(item.GetSubType().lower().split()) in subtype):
                c.thetable = item.GetSpecificType()
                if not countinvis:
                    set23(c.thetable)
                c.applyaction(c.thetable, info) 
    finally:
        info.generate()
        SpssClient.StopClient()


class PtColumns(object):
    """Modify display characteristics of a pivot table"""

    textstyles = {'regular' :SpssClient.SpssTextStyleTypes.SpssTSRegular,
                  'bold' : SpssClient.SpssTextStyleTypes.SpssTSBold,
                  'italic': SpssClient.SpssTextStyleTypes.SpssTSItalic,
                  'bolditalic' : SpssClient.SpssTextStyleTypes.SpssTSBoldItalic}

    def __init__(self, columns, dimension, level, hide,
                 widths, rowlabels, rowlabelwidths, textstyle, textcolor, bgcolor, applyto, customfunction, 
                 printlabels, regexp, tlook):
        """columns is a sequence of identifiers of columns to act on.
        It can include positive or negative numbers (or things that can be converted to these) and
        strings that will be matched to the lowest level of the column labels ignoring case.
        Use "<<ALL>>" to refer to all columns - not valid for hide.
        action is 'hide', 'show', or 'resize'
        If action == 'resize', widths must be specified.
        If dimension == 'rows' then rows are processed instead of columns.
        'rows' cannot be combined with 'resize'
        level specifies which level to check.  -1 is the innermost.  This only affects the behavior
        when using specific labels.
        widths is only used with action == 'resize'.  It is a sequence that specifies the width 
        or widths to be applied to the specified columns.  
        If it is a sequence of length 1, it is used for all specified columns.
        Otherwise, it is a sequence of sizes in points of the same length as the columns list.
	if regexp, nonnumerical text in columns is treated as a regular expression
        '"""

        if columns is None:
            columns = []
        attributesFromDict(locals())  # copy parameters
        self.encoding = locale.getlocale()[1]
        self.actionset = any([widths, rowlabelwidths, textstyle, textcolor, bgcolor, customfunction])
        if hide and self.actionset:
            raise ValueError(_("HIDE cannot be combined with other actions"))
        if  not (self.actionset or self.tlook):
            self.hide = True

        if not _isseq(columns):
            raise ValueError(_("Row or column specification must be a sequence of one or more values"))
        if widths:
            if dimension == 'rows':
                raise ValueError(_("The rows dimension cannot be combined with resizing"))
            if len(widths) == 1:
                self.widths = len(columns) * widths  # The <<ALL>> case is handled specially later
            if regexp and len(widths) > 0:
                raise ValueError(_("Regular expression mode is not available with widths"))
            if  len(columns) != len(self.widths):
                raise ValueError(_("The number of column sizes specified is different from number of columns specified."))
        rlgiven = rowlabels is not None
        rlwgiven = rowlabelwidths is not None
        if rlgiven ^ rlwgiven:
            raise ValueError(_("ROWLABELS and ROWLABELWIDTHS must be specified together"))
        if rowlabels:
            if len(rowlabelwidths) == 1:
                self.rowlabelwidths = len(rowlabels) * rowlabelwidths
            if  len(rowlabelwidths) != len(rowlabels):
                raise ValueError(_("The number of row label sizes specified is different from number of row labels specified."))
        else:
            if not columns:
                raise ValueError(_("SELECT is required unless ROWLABELS is specified"))
        self.columns = []
        regexplist = []
        for item in columns:
            try:
                self.columns.append(int(item))  # add all integer-like items to list
            except:
                # accumulate regexp text and regular text separately but <<ALL>> is special
                # self.regexp will be the compiled reg expression or False if no expressions or option is False
                if regexp and not columns[0] == "<<ALL>>":
                    regexplist.append(item)
                else:
                    self.columns.append(item)
        if regexplist:   # combine all regexp terms with or if any were found
            try:
                regexp = "|".join(["("+item+")" for item in regexplist])
                self.regexp = re.compile(regexp, re.LOCALE)
            except:
                reerr = sys.exc_info()[1]
                raise ValueError(_("Invalid regular expression: %s error: %s") % (regexp, str(reerr)))
        else:
            self.regexp = False
        if columns and columns[0] == "<<ALL>>" and hide:
            raise ValueError(_("HIDE cannot be applied with columns = <<ALL>>"))
        self.actionlist = [self.widths, self.rowlabelwidths, self.textstyle, self,textcolor, self.bgcolor]

        # build style code as a list of functions to call with user specifications
        # some parameters in the stock function signature are ignored in order to have
        # uniformity with the custom function

        if self.applyto.lower() in ["both", "labels", "datacells"]:
            self.applyto = self.applyto.lower()
        self.stylecalls = []
        if self.textstyle:
            def f(obj, row, column, numrows, numcols, section, ignore):
                obj.SetTextStyleAt(row, column, PtColumns.textstyles[self.textstyle])
            self.stylecalls.append(f)
        if self.textcolor:
            self.textcolor = RGB(self.textcolor)
            def f(obj, row, column, numrows, numcols, section, ignore):
                obj.SetTextColorAt(row, column, self.textcolor)
            self.stylecalls.append(f)
        if self.bgcolor:
            self.bgcolor = RGB(self.bgcolor)
            def f(obj, row, column, numrows, numcols, section, ignore):
                try:
                    obj.SetBackgroundColorAt(row, column, self.bgcolor)
                except:
                    raise SystemError(_("Set Background Color exception: %s %s %s") %(row, column, section))
            self.stylecalls.append(f)
        if not self.customfunction is None:
            for f in self.customfunction:
                self.stylecalls.append(resolvestr(f))
        self.previousUsedValue = ""


    def applyaction(self, pt, info):
        """Apply specified action to columns or rows of a pivot table.

        pt is the pivot table to process, on which GetSpecificType() is assumed to have been called."""

        #if self.action == 'show' and self.columns[0] == '<<ALL>>':
        #    pt.ShowAll()
        #    return

        if self.tlook:
            pt.SetTableLook(self.tlook)
        if self.widths and self.columns and self.columns[0] == '<<ALL>>':
            pt.SetDataCellWidths(self.widths[0])
        self.datacells = pt.DataCellArray()
        #self.datacells = fDataCellArray(pt)
        #self.rowlabelarray = pt.RowLabelArray()
        self.rowlabelarray = fRowLabelArray(pt)
        #self.columnlabelarray = pt.ColumnLabelArray()
        self.columnlabelarray = fColumnLabelArray(pt)

        self.numdatarows = self.datacells.GetNumRows()
        self.numdatacols = self.datacells.GetNumColumns()

        if self.dimension == 'columns':
            self.labels = self.columnlabelarray
            rowsorcols = self.labels.GetNumColumns()
            last = max(self.labels.GetNumRows() - 1, 0)
            self.printtablelabels(last+1, rowsorcols, "Columns", info)
            def swapper(i, j):
                return ((i + self.level + 1) % (last+1), j)
        else:
            self.labels = self.rowlabelarray
            rowsorcols = self.labels.GetNumRows()
            last = max(self.labels.GetNumColumns() -1, 0)
            self.printtablelabels(rowsorcols, last+1, "Rows", info)
            def swapper(i, j):
                return (j, (i + self.level + 1) % (last+1))

        specificrowsorcols = self.resolvecols(self.columns, rowsorcols, info)
        scset = set(specificrowsorcols)
        if self.widths:
            wdict = dict(zip(specificrowsorcols, self.widths))   # won't work with regexp

        try:
            pt.SetUpdateScreen(False)
            # process table data and label cells for width, hiding, and formatting
            for  roworcol in range(rowsorcols):
                i,j = swapper(last, roworcol)
                wkey = None
                # first see if row or column number was specified or taking all
                if roworcol in scset or "<<ALL>>" in scset:
                    wkey = roworcol
                else:
                    # otherwise see if text was specified or any regular expression matches
                    v = self.labels.GetValueAt(i,j)
                    if self.regexp:
                        if self.regexp.search(v):
                            wkey = v
                    else:
                        if v in scset:
                            wkey = v
                if not wkey is None:
                    if self.hide:
                        ###self.labels.HideLabelsWithDataAt(i,j)
                        self.hider(self.dimension, last, i, j)
                    else:
                        if self.widths and not "<<ALL>>" in scset:   #all case is already processed
                            self.datacells.ReSizeColumn(roworcol, wdict[wkey])
                        if self.actionset:
                            rc = self.dostyles(roworcol)
                            if rc is False:
                                break
                    #else:
                    #    self.labels.ShowAllLabelsAndDataInDimensionAt(i,j)
            if self.rowlabels:
                labels = self.rowlabelarray
                rowsorcols = labels.GetNumColumns()
                wdict = dict(zip(self.resolvecols(self.rowlabels, rowsorcols, info), self.rowlabelwidths))
                # process table data and label cells for width, hiding, and formatting
                for  roworcol in range(rowsorcols):
                    newwidth = wdict.get(roworcol, None)
                    if not newwidth is None:
                        labels.SetRowLabelWidthAt(1,roworcol, newwidth)
        finally:
            pt.SetUpdateScreen(True)

    def hider(self, dimension, last, i, j):
        """Hide specified row(s) or columns even if not innermost
		
		dimension is "rows" or "columns"
		last is the index of the innermost label in the dimension
		i and j index the label array for the matching label"""
        
        if dimension == "columns":
            hideloc = max(i, last)
            self.labels.HideLabelsWithDataAt(hideloc, j)
        else:
            hideloc = max(j, last)
            self.labels.HideLabelsWithDataAt(i,hideloc)
                    
    def dostyles(self, roworcol):
        """Apply any requested styles to labels and/or datacells.

        self.labels is the relevant dimension labels.
        self.datacells is the table data cell array
        roworcol is the current table row or column number"""

        # the datacells and labels objects are exposed as globals in case customfunctions need
        # the one not being passed in the style call

        if self.applyto != "labels":  #datacell styles
            expression = self.applyto not in ["both", "datacells"]
            rc = self.datacellstyles(roworcol, expression)
            if rc is False:
                return False

        if self.applyto in ["both", "labels"]:   # label styles
            rc = self.labelcellstyles(roworcol, self.labels.GetNumRows(), self.labels.GetNumColumns())
            if rc is False:
                return False

    def resolvecols(self, colarray, rowsorcols, info):
        """Return a list of column or row specifications for indexes with negative values resolved.

        rowsorcols is the actual number of data columns or rows in the current pivot table.
	If an out-of-bounds column number is found, it is removed from the list with a warning."""

        ret = []
        for item in colarray:
            try:
                item = int(item)
                if item < 0:
                    item += rowsorcols
                if not (0 <= item < rowsorcols):
                    info.addrow(_("""A specified row or column label number does not exist in a selected table.
It will be ignored.  Any column-specific width settings may be incorrect.
Label number: %s.  Table size: %s""")\
                           % (item, rowsorcols))
                    continue
            except ValueError:
                pass
            ret.append(item)
        return ret

    def printtablelabels(self, rows, cols, which, info):
        """Print row or column labels of table

	rows and cols are the dimensions.
	which is "rows" or "columns"."""
        if self.printlabels:
            info.addrow(_("Table Labels: %s.  Dimensions: %s, %s") % (which, rows, cols))
            for i in range(rows):
                for j in range(cols):
                    info.addrow("%d %d: %s" % (i, j, self.labels.GetValueAt(i, j)))
                    ###print i, j, self.labels.GetValueAt(i, j)

    def datacellstyles(self, roworcol, expression):
        """Apply datacell styles looping over rows or columns

	roworcol is the row or column number to process.
	expression is True if an applyto expression exists."""

        coldim = self.dimension == "columns"
        if coldim:
            limit = self.numdatarows
        else:
            limit = self.numdatacols

        for i in range(limit):
            outcome = True
            if coldim:
                row, col = i, roworcol
            else:
                row,col = roworcol, i
            if expression:
                try:   # this api is new in V18 or 17.0.2
                    x = float(self.datacells.GetUnformattedValueAt(row, col))
                except AttributeError:
                    x = floatex(self.datacells.GetValueAt(row, col), self.datacells.GetNumericFormatAt(row, col))
                except ValueError:
                    x = self.datacells.GetValueAt(row, col)   # 2/19/2015
                try:
                    if outcome:
                        outcome = eval(self.applyto, {'x':x, "i":i, "ii": roworcol})
                except (NameError, SyntaxError), e:
                    if not isinstance(self.applyto, unicode):
                        self.applyto = unicode(self.applyto, self.encoding)
                    raise ValueError(_("APPLYTO expression is invalid: %s") % self.applyto)
                except:
                    outcome = False

            if outcome:
                for f in self.stylecalls:
                    rc = f(self.datacells, row, col, self.numdatarows, self.numdatacols, "datacells",  self)
                    if rc is False:
                        return rc

    def labelcellstyles(self, roworcol, numlabelrows, numlabelcols):
        """Apply label styles

	roworcol is the row or column to process."""

        coldim = self.dimension == "columns"
        if coldim:
            limit = numlabelrows
        else:
            limit = numlabelcols

        for i in range((limit + self.level) % limit, limit):
            if coldim:
                row, col = i, roworcol
            else:
                row, col = roworcol, i
            try:
                for f in self.stylecalls:
                    rc = f(self.labels, row, col, numlabelrows, numlabelcols, "labels", self)
                    if rc is False:
                        return rc
            except:
                pass
def set23(pt):
    """Set pt incompatible if V23 or later
    
    pt is a pivot table"""
    # this function uses an api that does not exist before
    # V23 or perhaps a fixpack for V22 to turn off legacy 
    # table mode so that tables with hidden rows or columns can be
    # handled correctly
    
    try:
        if pt.IsLegacyTableCompatible():
            pt.SetLegacyTableCompatible(False)
    except:
        pass


def RGB(colorlist):
    """Return the integer RGB value for colorlist.

    colorlist is a triple of numbers between 0 and 255"""

    # type and range checks are assumed to be already done

    if len(colorlist) != 3:
        raise ValueError(_("Color specification must consist of three integers (R, G, and B values) between 0 and 255"))
    return int(colorlist[0]) + 256 * (int(colorlist[1]) + 256 * int(colorlist[2]))

def attributesFromDict(d):
    """build self attributes from a dictionary d."""

    # based on Python Cookbook, 2nd edition 6.18

    self = d.pop('self')
    for name, value in d.iteritems():
        setattr(self, name, value)

def resolvestr(afunc):
    """Return a callable for afunc and build its parameters

    afunc may be a callable object or a string in the form module.func
    or module.func(parm=value,...)to be imported.
    a parameter, _first is always created with value True if afunc is a string.
    """
    global CUSTOMPARAMS
    f, p = factor(afunc)
    CUSTOMPARAMS[f] = p   # add to parameter dictionary using module.function as the key

    if callable(f):
        return afunc   # no parameters
    else:
        bf = f.split(".")
        if len(bf) != 2:
            raise ValueError(_("function reference %s not valid") % f)
        exec "from %s import %s as customfunction" % (bf[0], bf[1])
        argspec = inspect.getargspec(customfunction)[0]
        nargs = len(argspec)
        if nargs < 7 or nargs > 8:
            argspecj = ", ".join(argspec)
            if not isinstance(argspecj, unicode):
                argspecj = unicode(argspecj, locale.getlocale()[1])
            raise ValueError(_("Invalid custom function signature.\nToo few arguments: %s") % argspecj)
        elif nargs > 7:       # indicates function provides for custom params
            customfunction = functools.partial(customfunction, custom=CUSTOMPARAMS[f])
        return customfunction

def factor(afunc):
    """decompose the string m.f or m.f(parms) and return function and parameter dictionary

    afunc has the form xxx or xxx(p1=value, p2=value,...)
    create a dictionary from the parameters consisting of at least _first:True.
    parameter must have the form name=value, name=value,... or name:value, name:value,...
    """

    mo = re.match(r"([^(]+)(\(.+\)$)", afunc)
    if not mo is None:       # parameters found, make a dictionary of them
        try:
            params = eval("dict" + mo.group(2) )
        except:
            mog = mo.group(2)
            if not isinstance(mog, unicode):
                mog = unicode(mog, locale.getlocale()[1])
            raise ValueError(_("Invalid customfunction parameter expression: %s") % mog)
        f = mo.group(1)
    else:
        params = {}
        f = afunc
    params["_first"] = True
    # basic sanity check for parameter expression
    if '(' in f or ')' in f or '=' in f or ',' in f:
        raise ValueError(_("Invalid customfunction parameter expression: %s") % afunc)
    return f, params

class NonProcPivotTable(object):
        """Accumulate an object that can be turned into a basic pivot table once a procedure state can be established"""
        
        def __init__(self, omssubtype, outlinetitle="", tabletitle="", caption="", rowdim="", coldim="", columnlabels=[],
                     procname="Messages"):
                """omssubtype is the OMS table subtype.
                caption is the table caption.
                tabletitle is the table title.
                columnlabels is a sequence of column labels.
                If columnlabels is empty, this is treated as a one-column table, and the rowlabels are used as the values with
                the label column hidden
                
                procname is the procedure name.  It must not be translated."""
                
                attributesFromDict(locals())
                self.rowlabels = []
                self.columnvalues = []
                self.rowcount = 0
        
        def addrow(self, rowlabel=None, cvalues=None):
            """Append a row labelled rowlabel to the table and set value(s) from cvalues.
            
            rowlabel is a label for the stub.
            cvalues is a sequence of values with the same number of values are there are columns in the table."""
                
            if cvalues is None:
                cvalues = []
            self.rowcount += 1
            if rowlabel is None:
                    self.rowlabels.append(str(self.rowcount))
            else:
                    self.rowlabels.append(rowlabel)
            if not _isseq(cvalues):
                    cvalues = [cvalues]
            self.columnvalues.extend(cvalues)
            
        def generate(self):
                """Produce the table assuming that a procedure state is now in effect if it has any rows."""
                
                privateproc = False
                if self.rowcount > 0:
                    import spss
                    try:
                            table = spss.BasePivotTable(self.tabletitle, self.omssubtype)
                    except:
                            StartProcedure(_("Messages"), self.procname)
                            privateproc = True
                            table = spss.BasePivotTable(self.tabletitle, self.omssubtype)
                    if self.caption:
                            table.Caption(self.caption)
                    # Note: Unicode strings do not work as cell values in 18.0.1 and probably back to 16
                    if self.columnlabels != []:
                            table.SimplePivotTable(self.rowdim, self.rowlabels, self.coldim, self.columnlabels, self.columnvalues)
                    else:
                            table.Append(spss.Dimension.Place.row,"rowdim",hideName=True,hideLabels=True)
                            table.Append(spss.Dimension.Place.column,"coldim",hideName=True,hideLabels=True)
                            colcat = spss.CellText.String("Message")
                            for r in self.rowlabels:
                                    cellr = spss.CellText.String(r)
                                    table[(cellr, colcat)] = cellr
                    if privateproc:
                            spss.EndProcedure()
def _isseq(obj):
        """Return True if obj is a sequence, i.e., is iterable.
        
        Will be False if obj is a string or basic data type"""
        
        if isinstance(obj, basestring):
                return False
        else:
                try:
                        iter(obj)
                except:
                        return False
                return True

def StartProcedure(procname, omsid):
    """Start a procedure
    
    procname is the name that will appear in the Viewer outline.  It may be translated
    omsid is the OMS procedure identifier and should not be translated.
    
    Statistics versions prior to 19 support only a single term used for both purposes.
    For those versions, the omsid will be use for the procedure name.
    
    While the spss.StartProcedure function accepts the one argument, this function
    requires both."""
    
    import spss
    try:
        spss.StartProcedure(procname, omsid)
    except TypeError:  #older version
        spss.StartProcedure(omsid)
