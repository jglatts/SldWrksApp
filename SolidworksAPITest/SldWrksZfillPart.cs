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

class SldWrksZfillPart {

    private String path;

    public SldWrksZfillPart(String path) {
        this.path = path;
    }

    // Change Zfill part 
    public void changePart(ISldWorks app)
    {
        ModelDoc2 default_dwg = null;
        Dimension swDim = null;

        default_dwg = app.OpenDoc6(path, (int)PartTypes.PART, 256, "", 0, 0);
        if (default_dwg == null)
        {
            Console.WriteLine("error opening part");
            return;
        }

        // get and change part length
        getPartSelection(default_dwg, ref swDim, "Boss-Extrude1", "BODYFEATURE", "D1");
        changePartDimensions(default_dwg, swDim, 105);


        // get and change part height
        getPartSelection(default_dwg, ref swDim, "Sketch1", "SKETCH", "D2");
        changePartDimensions(default_dwg, swDim, 10);
    }

    // Select a feature\sketch from the part
    private void getPartSelection(ModelDoc2 doc, ref Dimension swDim, String name, String swType, String dim)
    {
        Feature swFeature;
        SelectionMgr swSelectionManager;

        doc.Extension.SelectByID2(name, swType, 0, 0, 0, false, 0, null, 0);
        swSelectionManager = (SelectionMgr)doc.SelectionManager;
        swFeature = (Feature)swSelectionManager.GetSelectedObject6(1, -1);
        swDim = (Dimension)swFeature.Parameter(dim);
        double dim_mm = Convert.ToDouble(swDim.SystemValue.ToString());
        Console.WriteLine("the dim is " + (dim_mm * 1000) + "mm");
    }

    // Change part dimensnions after a call to getPartSelection
    private void changePartDimensions(ModelDoc2 doc, Dimension swDim, int new_val_int)
    {
        double new_val = new_val_int / 1000.0;
        swDim.SetSystemValue3(new_val, 1, null);
        doc.EditRebuild3();
        double dim_mm = Convert.ToDouble(swDim.SystemValue.ToString());
        Console.WriteLine("the dim is " + (dim_mm * 1000) + "mm after change");
    }

}
