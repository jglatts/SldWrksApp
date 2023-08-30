/**
 * 
 *  Solidworks API wrapper in C#
 *  Working for opening a part\asm and modyifing dimensions
 *  More to come!
 *  
 *  See below link for saving DWG as PDF
 *  https://help.solidworks.com/2019/english/api/sldworksapi/save_file_as_pdf_example_csharp.htm
 * 
 *  Author: John Glatts
 *  Date: 8/29/23
 */
using System;
using System.Diagnostics;
using System.Linq;
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

    // Create a new part drawing
    private void createPartDrawing(String path) {
        ModelDoc2 default_dwg = app.OpenDoc6(path, (int)PartTypes.DWG, 256, "", 0, 0);
        
        if (default_dwg == null) {
            Console.WriteLine("error opening a drawing");
            return;
        }
        
    }

    // Driver method to open, edit, and make a drawing of a part
    public void run() {
        SldWrksZfillPart zfill = new SldWrksZfillPart(@"Z:\Manufacturing\SWAutomation\Zfill-Default.SLDPRT");
        if (app == null)  
            return;

        zfill.changePart(app);
        
        
    }

    // Class main method
    public static void Main(string[] args) {
        new SldWrksApp().run();
    }

}

