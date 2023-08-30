/**
 * 
 *  Solidworks API wrapper in C#
 * 
 *  Author: John Glatts
 *  Date: 8/29/23
 */
using System;
using System.Diagnostics;
using System.Linq;
using System.Reflection.Metadata.Ecma335;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using SolidWorks.Interop.sldworks;

/*
 *  Main class to edit solidworks parts and drawings
 *  Can modify sketches, features, notes, etc...
 */
class SldWrksApp {

    // SolidWorks application object
    private ISldWorks app;

    // Constructor 
    public SldWrksApp() {
        app = OpenSldWrksApp.getApp();
    }

    // Driver method to open, edit, and save a new part
    public void makePart(PartDimension part_dims) { 
        SldWrksZfillPart zfill = new SldWrksZfillPart(@"Z:\Manufacturing\SWAutomation\Zfill-default.SLDPRT");
        if (app == null)
            return;
        
        zfill.changePart(app, part_dims);
    }

    // Class main method
    public static void Main(string[] args) {
        new SldWrksApp().makePart(new PartDimension(60.0, 10.0, 6.0, .015, 4.0));
    }

}

