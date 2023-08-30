using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


// fix for OpenDoc6() method, using values
// see https://tinyurl.com/6222p97j for values
enum PartTypes
{
    PART = 1,
    ASM = 2,
    DWG = 3,
}

class PartDimension{
    public double length;
    public double width;
    public double height;
    public double pitch;
    public double keep_off;

    public PartDimension(double length, double width, double height, double pitch, double keep_off) 
    {
        this.length = length;
        this.width = width;
        this.height = height;
        this.pitch = pitch;
        this.keep_off = keep_off;
    }
    

}
