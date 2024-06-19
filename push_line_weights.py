#   "Line Weights Dialog Auto-Filler"
#   Copyright © 2021 Jared M. Holloway
#   License: MIT
#   Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
#   The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
#   THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


#   TODO:
#   - Move VBA functions into python (xlwings?)
#   - Add functions to auto-fill Annotation Line Weights and Perspective Line Weights tables
#   - Use CountColumns() to validate incoming vals string; handle invalid case (possible?  if not, just popup warn user; also do this first as placeholder)
#     Invalid Case: 
#       1) close Line Weights window; check that all incoming columns exist in project view scale table; if any do not, add them (how to call Revit API from CPython?)
#       2) reopen Line Weights window; Delete all existing columns (since we can't know or find out which view scales are already in the table); Add all incoming columns to table; InputVals()


import pywinauto as ui
from time import sleep


def GetRevitProc():
    # get all revit procs
    revit_procs = list(filter(lambda x:
            ui.application.handleprops.is_toplevel_window(x),
            ui.findwindows.find_windows(title_re="Autodesk Revit.*")
            ))
    # get 1 revit proc
    if len(revit_procs) < 1:
        print("Please open a Revit project first.")
    elif len(revit_procs) == 1:
        revit_proc = revit_procs[0]
    else:
        pick_list = []
        for n,proc in enumerate(revit_procs):
            pick_list.append(str(n))
            print( "{0})  {1}".format(n, ui.application.handleprops.text(proc)) )
        prompt = "Which Revit instance? ({0}):  ".format(str("|".join(pick_list)))
        revit_proc = revit_procs[int(input(prompt))]
    return revit_proc

def GetRevitApp(revit_proc):
    # get Revit application
    revit_app = ui.Application().connect(handle=revit_proc)
    return revit_app

def GetRevitWnd(revit_app):
    # get Revit main window
    revit_wnd = revit_app.top_window()
    return revit_wnd

def GetLineWeightsWnd(revit_app):
    # open Line Weights window
    GetRevitWnd(revit_app).post_command(32946)
    # wait for Line Weights window to open
    sleep(1)
    # get Line Weights dialog
    lw_wnd = revit_app.window(title_re="Line Weights")
    return lw_wnd

def InitVars(lw_wnd):
    # get help control
    hlp = lw_wnd['help'].set_focus()
    # get dialog control
    dlg = hlp.parent()
    # get tab control
    tab = lw_wnd['tabcontrol'].select(0)
    # cycle thru tabs to expose all controls in window
    for i in range(3):
        tab.select(i)
        sleep(.5)
    # get list of Line Weight table controls
    gxwnds = list(filter(lambda x: "GXWND" in str(x), lw_wnd.children()))

    # find correct window handle for each Line Weight table control
    tables = []
    for i in range(3):
        tab.select(i)
        sleep(.5)
        table = list(filter(lambda x: x.is_visible(), gxwnds))
        assert len(table) == 1
        tables.append(table[0])
    # Model Line Weights tab
    mlw = tables[0]
    # Perspective Line Weights tab
    plw = tables[1]
    # Annotation Line Weights tab
    alw = tables[2]

    # go to Model Line Weights tab
    tab.select(0)
    sleep(.5)
    return dlg,hlp,tab,mlw,plw,alw

def CountColumns(mlw):
    n = 0
    # activate Model Line Weight table control
    mlw.double_click()
    # move keyboard cursor to A1 cell
    mlw.send_keystrokes("^{HOME}^{RIGHT}")
    # count columns
    while "StaticWrapper" not in str(mlw.parent().get_focus()):
        n += 1
        mlw.parent().type_keys("+{TAB}")
    return n

def InputValues(table,vals):
    # activate Model Line Weight table control
    table.set_focus()
    # move keyboard cursor to A1 cell
    table.double_click()
    table.type_keys("^{HOME}")
    # input values
    for v in vals:
        table.type_keys(v)
        sleep(0.01)
        table.type_keys("\"")
        sleep(0.01)
        table.type_keys("{TAB}")
        sleep(0.01)

def Main(vals):
    revit_proc = GetRevitProc()
    revit_app = GetRevitApp(revit_proc)
    revit_wnd = GetRevitWnd(revit_app)
    lw_wnd = GetLineWeightsWnd(revit_app)
    dlg,hlp,tab,mlw,plw,alw = InitVars(lw_wnd)
    InputValues(mlw,vals)
    

if __name__ == "__main__":
    
    import contextlib, io
    with contextlib.redirect_stderr(io.StringIO()):
        try:
            import argparse
            parser = argparse.ArgumentParser(
                formatter_class=argparse.RawDescriptionHelpFormatter,
                description="""A script to programmatically fill out Revit Line Weights dialog."""
            )
            parser.add_argument(
                "--vals",
                dest="vals",
                nargs='+',
                help='<Required> Set flag',
                required=True
            )
            Main(parser.parse_args().vals)
        except BaseException:
            # import sys
            # print(sys.exc_info()[0])
            # import traceback
            # print(traceback.format_exc())
            pass
        finally:
            pass
