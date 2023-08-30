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


class SldWrksDrw {

    // Create a new part drawing
    public static void updatePartDrawing(ISldWorks app, String path)
    {
        ModelDoc2 default_dwg = app.OpenDoc6(path, (int)PartTypes.DWG, 256, "", 0, 0);

        if (default_dwg == null)
        {
            Console.WriteLine("error opening a drawing");
            return;
        }

        savePDF(app, default_dwg);
    }

    // Save drawing as PDF
    private static void savePDF(ISldWorks app, ModelDoc2 model) {
        ExportPdfData swExportPDFData = default(ExportPdfData);
        ModelDocExtension swModExt = default(ModelDocExtension);
        String filename = "test.pdf";
        int errors = 0;
        int warnings = 0;

        swModExt = (ModelDocExtension)model.Extension;
        swExportPDFData = (ExportPdfData)app.GetExportFileData(1);
        swExportPDFData.ViewPdfAfterSaving = true;
        swModExt.SaveAs(filename, 0, 1, swExportPDFData, ref errors, ref warnings);
    }

}
