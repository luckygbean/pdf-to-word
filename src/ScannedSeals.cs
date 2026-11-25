using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

// A flattened scan has no stamp object. Split visible red ink from neutral
// paper/print only inside compact seal outlines. Never infer missing strokes.
internal sealed class ScannedSeals
{
    internal readonly int Width, Height;
    internal readonly byte[] Pixels;
    private readonly double scaleX, scaleY;

    internal ScannedSeals(Stream png, LayoutPage page)
    {
        png.Position=0;
        using(var image=new Bitmap(png))
        {
            Width=image.Width;
            Height=image.Height;
            scaleX=Width/page.Width;
            scaleY=Height/page.Height;
            var area=new Rectangle(0,0,Width,Height);
            using(var rgb=image.Clone(area,PixelFormat.Format32bppArgb))
            {
                var bits=rgb.LockBits(area,ImageLockMode.ReadOnly,PixelFormat.Format32bppArgb);
                try
                {
                    Pixels=new byte[checked(Width*Height*4)];
                    for(int y=0;y<Height;y++)
                        Marshal.Copy(IntPtr.Add(bits.Scan0,y*bits.Stride),Pixels,y*Width*4,Width*4);
                }
                finally { rgb.UnlockBits(bits); }
            }
        }
    }

    private int Redness(int i)
    {
        return Pixels[i+2]-Math.Max(Pixels[i],Pixels[i+1]);
    }

    internal List<Rectangle> Locate()
    {
        int step=Math.Max(3,(int)Math.Round(3*Math.Min(scaleX,scaleY)));
        int cols=(Width+step-1)/step, rows=(Height+step-1)/step;
        var occupied=new bool[cols*rows];
        for(int y=0;y<Height;y+=2)
            for(int x=0;x<Width;x+=2)
                if(Redness((y*Width+x)*4)>36) occupied[y/step*cols+x/step]=true;

        var found=new List<Rectangle>();
        var queue=new Queue<int>();
        for(int i=0;i<occupied.Length;i++)
        {
            if(!occupied[i]) continue;
            int left=cols,top=rows,right=0,bottom=0;
            occupied[i]=false;
            queue.Enqueue(i);
            while(queue.Count>0)
            {
                int current=queue.Dequeue(),cx=current%cols,cy=current/cols;
                left=Math.Min(left,cx); right=Math.Max(right,cx);
                top=Math.Min(top,cy); bottom=Math.Max(bottom,cy);
                for(int yy=Math.Max(0,cy-1);yy<=Math.Min(rows-1,cy+1);yy++)
                    for(int xx=Math.Max(0,cx-1);xx<=Math.Min(cols-1,cx+1);xx++)
                    {
                        int next=yy*cols+xx;
                        if(occupied[next]) { occupied[next]=false; queue.Enqueue(next); }
                    }
            }
            var box=Rectangle.FromLTRB(left*step,top*step,Math.Min(Width,(right+1)*step),Math.Min(Height,(bottom+1)*step));
            if(box.Width/scaleX<14 || box.Height/scaleY<14 || box.Width/scaleX>180 || box.Height/scaleY>180) continue;
            if(box.Width<box.Height*.64 || box.Width>box.Height*1.56) continue;
            box=TightRedBounds(box);
            if(box.Width>0 && HasOutline(box))
            {
                int padding=(int)Math.Ceiling(3*Math.Max(scaleX,scaleY));
                box.Inflate(padding,padding);
                found.Add(Rectangle.Intersect(box,new Rectangle(0,0,Width,Height)));
            }
        }
        // A seal can enclose disconnected characters or an emblem.
        return found.Where(box=>!found.Any(other=>other!=box && other.Contains(box))).Distinct().ToList();
    }

    private Rectangle TightRedBounds(Rectangle area)
    {
        int l=area.Right,t=area.Bottom,r=-1,b=-1;
        for(int y=area.Top;y<area.Bottom;y++)
            for(int x=area.Left;x<area.Right;x++)
                if(Redness((y*Width+x)*4)>20)
                {
                    l=Math.Min(l,x);r=Math.Max(r,x);t=Math.Min(t,y);b=y;
                }
        return r<l ? Rectangle.Empty : Rectangle.FromLTRB(l,t,r+1,b+1);
    }

    private bool HasOutline(Rectangle box)
    {
        var sectors=new bool[24];
        int colored=0, samples=0;
        for(int y=box.Top;y<box.Bottom;y+=2)
            for(int x=box.Left;x<box.Right;x+=2)
            {
                samples++;
                if(Redness((y*Width+x)*4)<=28) continue;
                colored++;
                double dx=2.0*(x-box.Left+.5)/box.Width-1,dy=2.0*(y-box.Top+.5)/box.Height-1;
                double squared=dx*dx+dy*dy;
                if(squared<.75 || squared>1.15) continue;
                int sector=(int)((Math.Atan2(dy,dx)+Math.PI)*sectors.Length/(2*Math.PI));
                sectors[Math.Min(sectors.Length-1,sector)]=true;
            }
        return colored>samples*.018 && colored<samples*.60 && sectors.Count(s=>s)>=21;
    }

    internal LayoutStamp Lift(Rectangle box)
    {
        var rgba=new byte[box.Width*box.Height*4];
        int stride=box.Width+1;
        var redCount=new int[stride*(box.Height+1)];
        int fringe=(int)Math.Ceiling(2.8*Math.Min(scaleX,scaleY));
        // JPEG chroma subsampling leaves a neutral fringe immediately beside
        // red strokes. Attribute pale fringe pixels only to nearby seal ink.
        for(int y=0;y<box.Height;y++)
        {
            int row=0;
            for(int x=0;x<box.Width;x++)
            {
                if(Redness(((y+box.Y)*Width+x+box.X)*4)>18) row++;
                redCount[(y+1)*stride+x+1]=redCount[y*stride+x+1]+row;
            }
        }
        for(int y=0;y<box.Height;y++)
            for(int x=0;x<box.Width;x++)
            {
                int i=((y+box.Y)*Width+x+box.X)*4, j=(y*box.Width+x)*4;
                int red=Pixels[i+2],low=Math.Min(Pixels[i+1],Pixels[i]);
                int l=Math.Max(0,x-fringe),r=Math.Min(box.Width,x+fringe+1);
                int t=Math.Max(0,y-fringe),b=Math.Min(box.Height,y+fringe+1);
                bool adjacent=redCount[b*stride+r]-redCount[t*stride+r]-redCount[b*stride+l]+redCount[t*stride+l]>0;
                bool paleFringe=adjacent && red>195;
                if(!paleFringe && Redness(i)<Math.Max(5,red/33)) continue;
                int paper=Math.Min(255,red+(red-low)/2);
                if(adjacent && paper>183) paper=255;
                if(paper<red || paper<low) continue;
                int alpha=(paper-low)*255/Math.Max(1,paper);
                if(alpha==0) continue;
                for(int c=0;c<3;c++)
                {
                    int light=Pixels[i+c]*255/Math.Max(1,paper);
                    rgba[j+c]=(byte)Math.Max(0,Math.Min(255,255-(255-light)*255/alpha));
                    Pixels[i+c]=(byte)paper;
                }
                rgba[j+3]=(byte)alpha;
            }
        using(var encoded=new MemoryStream())
        {
            SavePng(encoded,rgba,box.Width,box.Height);
            return new LayoutStamp{Left=box.X/scaleX,Top=box.Y/scaleY,Width=box.Width/scaleX,
                Height=box.Height/scaleY,Png=encoded.ToArray()};
        }
    }

    internal static void SavePng(Stream target,byte[] pixels,int width,int height)
    {
        var pin=GCHandle.Alloc(pixels,GCHandleType.Pinned);
        try
        {
            using(var bmp=new Bitmap(width,height,width*4,PixelFormat.Format32bppArgb,pin.AddrOfPinnedObject()))
                bmp.Save(target,ImageFormat.Png);
        }
        finally { pin.Free(); }
        target.Position=0;
    }
}
