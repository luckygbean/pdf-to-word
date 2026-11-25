using System;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Windows.Globalization;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;

internal static class StampedScanText
{
    // Chinese OCR can read a printed hyphen as the ideograph for "one".
    // Restrict that interpretation to grouped alphanumeric identifiers.
    private static readonly Regex Identifier=new Regex(@"[A-Za-z0-9]+(?:[-‐‑–—一][A-Za-z0-9]+){2,}");
    private static string Comparable(string s)
    {
        return Regex.Replace(s.ToUpperInvariant(),@"[‐‑–—一]","-").Replace('I','1').Replace('L','1').Replace('O','0');
    }

    internal static void CheckIdentifiers(LayoutPage cover,Func<string[]> nearby)
    {
        string[] references=null;
        foreach(var line in cover.Lines)
        {
            line.Text=Identifier.Replace(line.Text,match=>
            {
                if(match.Length<8 || !match.Value.Contains("一") || !match.Value.Any(Char.IsDigit)) return match.Value;
                if(references==null) references=nearby().SelectMany(s=>Identifier.Matches(s).Cast<Match>())
                    .Select(m=>m.Value).Distinct().ToArray();
                string same=Comparable(match.Value);
                var candidates=references.Where(s=>Comparable(s)==same).Take(2).ToArray();
                return candidates.Length==1?candidates[0]:match.Value;
            });
        }
    }

    private static string Join(OcrLine line)
    {
        var value=new StringBuilder();
        OcrWord previous=null;
        foreach(var word in line.Words)
        {
            if(String.IsNullOrWhiteSpace(word.Text)) continue;
            if(previous!=null && previous.Text.Last()<128 && word.Text[0]<128 &&
               word.BoundingRect.Left-previous.BoundingRect.Right>.36*word.BoundingRect.Height)
                value.Append(' ');
            value.Append(word.Text);
            previous=word;
        }
        return OcrText.CleanLine(value.ToString()).Trim();
    }

    private static Rectangle InkBounds(ScannedSeals scan, Rectangle area, out double density,out int inkGray)
    {
        int l=area.Right,r=-1,t=area.Bottom,b=-1,dark=0,paper=0,color=0;
        var histogram=new int[161];
        inkGray=0;
        for(int y=area.Top;y<area.Bottom;y++)
            for(int x=area.Left;x<area.Right;x++)
            {
                int i=(y*scan.Width+x)*4;
                int high=Math.Max(scan.Pixels[i],Math.Max(scan.Pixels[i+1],scan.Pixels[i+2]));
                int low=Math.Min(scan.Pixels[i],Math.Min(scan.Pixels[i+1],scan.Pixels[i+2]));
                if(high-low>45) { color++; continue; }
                if(low>203) paper++;
                if(high>160) continue;
                histogram[(scan.Pixels[i]+scan.Pixels[i+1]+scan.Pixels[i+2])/3]++;
                dark++;
                l=Math.Min(l,x); r=Math.Max(r,x); t=Math.Min(t,y); b=y;
            }
        int size=area.Width*area.Height;
        if(r<=l || b<=t || color>size*.015 || paper<size*.30)
        {
            density=0;
            return Rectangle.Empty;
        }
        density=(double)dark/((r-l+1)*(b-t+1));
        int total=0;
        while(inkGray<160 && (total+=histogram[inkGray])<dark*.16) inkGray++;
        return Rectangle.FromLTRB(l,t,r+1,b+1);
    }

    private static async Task<OcrResult> RecognizeAsync(SoftwareBitmap bitmap,OcrEngine engine)
    {
        if(engine==null) return null;
        try { return await engine.RecognizeAsync(bitmap); }
        catch(COMException) { return null; }
        catch(UnauthorizedAccessException) { return null; }
    }

    internal static async Task<int> ReadAsync(ScannedSeals scan, LayoutPage page)
    {
        var engine=OcrText.CreateEngine();
        if(engine==null || Math.Max(scan.Width,scan.Height)>OcrEngine.MaxImageDimension) return 0;
        var contrast=new byte[scan.Pixels.Length];
        for(int i=0;i<contrast.Length;i+=4)
        {
            int level=Math.Max(scan.Pixels[i],Math.Max(scan.Pixels[i+1],scan.Pixels[i+2]));
            byte gray=(byte)Math.Max(0,Math.Min(255,(level-48)*255/145));
            contrast[i]=contrast[i+1]=contrast[i+2]=gray;
            contrast[i+3]=255;
        }
        OcrResult result,latin=null;
        using(var bitmap=new SoftwareBitmap(BitmapPixelFormat.Bgra8,scan.Width,scan.Height,BitmapAlphaMode.Ignore))
        {
            bitmap.CopyFromBuffer(contrast.AsBuffer());
            result=await RecognizeAsync(bitmap,engine);
            if(result==null) return 0;
            if(result.Lines.Any(l=>Identifier.IsMatch(Join(l))) && !engine.RecognizerLanguage.LanguageTag.StartsWith("en"))
            {
                var english=OcrEngine.TryCreateFromLanguage(new Language("en-US"));
                latin=await RecognizeAsync(bitmap,english);
            }
        }
        if(Math.Abs(result.TextAngle.GetValueOrDefault())>3) return 0;
        double sx=page.Width/scan.Width, sy=page.Height/scan.Height;
        var frame=new Rectangle(0,0,scan.Width,scan.Height);
        int count=0;
        foreach(var line in result.Lines)
        {
            string value=Join(line);
            if(value.Length==0) continue;
            var words=line.Words.Select(w=>w.BoundingRect).ToArray();
            var area=Rectangle.FromLTRB((int)words.Min(w=>w.Left),(int)words.Min(w=>w.Top),
                (int)Math.Ceiling(words.Max(w=>w.Right)),(int)Math.Ceiling(words.Max(w=>w.Bottom)));
            area.Inflate(Math.Max(2,area.Height/15),Math.Max(2,area.Height/9));
            area=Rectangle.Intersect(frame,area);
            if(latin!=null && Identifier.IsMatch(value))
            {
                var matches=latin.Lines.Where(l=>l.Words.Any(w=>area.Contains(
                    (int)(w.BoundingRect.X+w.BoundingRect.Width/2),(int)(w.BoundingRect.Y+w.BoundingRect.Height/2))))
                    .SelectMany(l=>Identifier.Matches(Join(l)).Cast<Match>()).Select(m=>m.Value).Distinct().ToArray();
                value=Identifier.Replace(value,m=>
                {
                    var agree=matches.Where(s=>Comparable(s)==Comparable(m.Value)).Take(2).ToArray();
                    return agree.Length==1?agree[0]:m.Value;
                });
            }
            double density;
            int inkGray;
            var ink=InkBounds(scan,area,out density,out inkGray);
            if(ink.IsEmpty) continue;
            double height=ink.Height*sy, width=ink.Width*sx;
            double units=value.Sum(ch=>ch>255?1.0:.53);
            double font=Math.Max(height*.90,Math.Min(height*1.3,width/units));
            page.Lines.Add(new LayoutLine{Text=value,Left=ink.Left*sx,Right=ink.Right*sx,Top=ink.Top*sy,
                Bottom=ink.Bottom*sy,Baseline=ink.Bottom*sy-font*.04, FontSize=font,
                InkLeft=ink.Left*sx,InkRight=ink.Right*sx,FontName=height<13?"SimSun":"SimHei",
                Bold=density>.35 && font>22,InkGray=inkGray});
            for(int y=area.Top;y<area.Bottom;y++)
                for(int x=area.Left;x<area.Right;x++)
                {
                    int i=(y*scan.Width+x)*4;
                    scan.Pixels[i]=scan.Pixels[i+1]=scan.Pixels[i+2]=255;
                }
            count++;
        }
        return count;
    }
}
