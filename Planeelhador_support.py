#! /usr/bin/env python3
#  -*- coding: utf-8 -*-
#
# Support module generated by PAGE version 8.0
#  in conjunction with Tcl version 8.6
#    Mar 30, 2024 08:32:50 PM -03  platform: Windows NT

import os
import tkinter as tk
from tkinter.constants import *

from planeelhador import TopLevelFormulario

def main(*args):
    savedir = os.getcwd()
    abspath = os.path.abspath(__file__)
    dname = os.path.dirname(abspath)
    os.chdir(dname)
    currdir = os.getcwd()
    print("Save to {}\nWork on {}\n".format(savedir, currdir))
    '''Main entry point for the application.'''
    global root
    root = tk.Tk()
    root.protocol( 'WM_DELETE_WINDOW' , root.destroy)
    # Creates a toplevel widget.
    global _top1, _w1
    top1 = root
    formulario = TopLevelFormulario(top1, savedir)
    root.mainloop()
    if (formulario.exitStatus):
        print("Success!")
        print(formulario.getFormInfo())
    
if __name__ == "__main__":
    main()