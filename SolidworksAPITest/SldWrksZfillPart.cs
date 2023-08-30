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
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using SolidWorks.Interop.sldworks;

class SldWrksZfillPart {

    private String path;
    public SldWrksZfillPart(String path) {
        this.path = path;
    }

    // Change Zfill part 
    public void changePart(ISldWorks app, PartDimension part_dims)
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
        getPartSelection(default_dwg, ref swDim, "Boss-Extrude1", "BODYFEATURE", "D1", "length");
        changePartDimensions(default_dwg, swDim, part_dims.length / 1000);

        // get and change part height
        getPartSelection(default_dwg, ref swDim, "Sketch1", "SKETCH", "D2", "height");
        changePartDimensions(default_dwg, swDim, part_dims.height / 1000);

        // get and change part width
        part_dims.width = (part_dims.width / 25.4) - .002;
        getPartSelection(default_dwg, ref swDim, "Sketch1", "SKETCH", "D5", "width");
        changePartDimensions(default_dwg, swDim, part_dims.width / 39.37);

        // get and change keepoff, unit in mm
        getPartSelection(default_dwg, ref swDim, "Sketch1", "SKETCH", "D4", "keepoff");
        changePartDimensions(default_dwg, swDim, part_dims.keep_off / 1000);

        // get and change wire pitch
        getPartSelection(default_dwg, ref swDim, "Sketch3", "SKETCH", "D3", "pitch");
        changePartDimensions(default_dwg, swDim, part_dims.pitch / 39.37);

        // get and change number of wires
        int num_wires = Convert.ToInt32(((part_dims.length / 25.4) - .01) / part_dims.pitch);
        Console.WriteLine("Computed Num Wires = " + num_wires + "\n\n");
        getPartSelection(default_dwg, ref swDim, "Sketch3", "SKETCH", "D4", "num_wires");
        changePartDimensions(default_dwg, swDim, num_wires);

        // get and change wire span
        double wire_span = (num_wires-1) * part_dims.pitch;
        Console.WriteLine("Computed Wire Span = " + wire_span + "in\n\n");
        getPartSelection(default_dwg, ref swDim, "Sketch3", "SKETCH", "D1", "wire_span");
        changePartDimensions(default_dwg, swDim, wire_span / 39.37);

        // update the drawing
        SldWrksDrw.updatePartDrawing(app, @"Z:\Manufacturing\SWAutomation\Zfill-default.SLDDRW");
    }

    // Select a feature\sketch from the part
    private void getPartSelection(ModelDoc2 doc, ref Dimension swDim, String name, String swType, String dim, String valName)
    {
        Feature swFeature;
        SelectionMgr swSelectionManager;

        doc.Extension.SelectByID2(name, swType, 0, 0, 0, false, 0, null, 0);
        swSelectionManager = (SelectionMgr)doc.SelectionManager;
        swFeature = (Feature)swSelectionManager.GetSelectedObject6(1, -1);
        swDim = (Dimension)swFeature.Parameter(dim);
        double dim_mm = Convert.ToDouble(swDim.SystemValue.ToString());
        Console.WriteLine("the dim of " + valName + " is " + (dim_mm * 1000) + "mm");
    }

    // Change part dimensnions after a call to getPartSelection
    private void changePartDimensions(ModelDoc2 doc, Dimension swDim, double new_val)
    {
        swDim.SetSystemValue3(new_val, 1, null);
        doc.EditRebuild3();
        double dim_mm = Convert.ToDouble(swDim.SystemValue.ToString());
        Console.WriteLine("the dim is " + (dim_mm * 1000) + "mm after change\n");
    }

}
