using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Markup;
using System.Windows.Media;
using System.Xml.Linq;

internal sealed partial class RuledWord
{
    private sealed class FontUse
    {
        public readonly HashSet<int> Regular=new HashSet<int>();
        public readonly HashSet<int> Bold=new HashSet<int>();
    }

    private readonly Dictionary<string,FontUse> usedFonts=new Dictionary<string,FontUse>(StringComparer.OrdinalIgnoreCase);

    private void RememberFont(string family,string value,bool bold)
    {
        FontUse use;
        if(!usedFonts.TryGetValue(family,out use)) usedFonts.Add(family,use=new FontUse());
        for(int i=0;i<value.Length;i++)
        {
            int code=char.ConvertToUtf32(value,i);
            if(code>0xffff) i++;
            use.Regular.Add(code); // Also supplies glyphs used with synthetic bold.
            if(bold) use.Bold.Add(code);
        }
    }

    private static byte[] FontSubset(FontFamily family,HashSet<int> codes,bool bold)
    {
        if(codes.Count==0) return null;
        try
        {
            var face=new Typeface(family,FontStyles.Normal,bold?FontWeights.Bold:FontWeights.Normal,FontStretches.Normal);
            GlyphTypeface glyphs;
            if(!face.TryGetGlyphTypeface(out glyphs) || glyphs.StyleSimulations!=StyleSimulations.None) return null;
            if(glyphs.EmbeddingRights!=FontEmbeddingRight.Editable && glyphs.EmbeddingRights!=FontEmbeddingRight.Installable)
                return null;
            var indices=new HashSet<ushort>{0};
            ushort id;
            foreach(int code in codes)
                if(glyphs.CharacterToGlyphMap.TryGetValue(code,out id)) indices.Add(id);
            return indices.Count>1?glyphs.ComputeSubset(indices):null;
        }
        // A missing, inaccessible or unsupported local font leaves the font name
        // and its alternate name available to the reader for substitution.
        catch(IOException) { return null; }
        catch(FileFormatException) { return null; }
        catch(UnauthorizedAccessException) { return null; }
        catch(ArgumentException) { return null; }
        catch(NotSupportedException) { return null; }
        catch(InvalidOperationException) { return null; }
        catch(COMException) { return null; }
    }

    private void AddFontPart(XElement font,XElement rels,byte[] data,bool bold)
    {
        if(data==null || data.Length<32) return;
        var key=Guid.NewGuid();
        string hex=key.ToString("N");
        // ECMA-376 17.8.1 uses the reverse of the printed GUID byte sequence.
        for(int i=0;i<32;i++)
            data[i]^=Convert.ToByte(hex.Substring((15-i%16)*2,2),16);
        string id="font"+(rels.Elements().Count()+1), target="fonts/"+id+".odttf";
        using(var output=archive.CreateEntry("word/"+target).Open()) output.Write(data,0,data.Length);
        rels.Add(Relationship(id,"font",target));
        font.Add(new XElement(W+(bold?"embedBold":"embedRegular"),new XAttribute(R+"id",id),
            At("fontKey",key.ToString("B").ToUpperInvariant()),At("subsetted",1)));
    }

    private bool WriteFontTable()
    {
        // Paragraph marks and newly entered form values use the Normal style
        // even if the PDF contains no explicit run in its Latin font.
        string keyboard=new string(Enumerable.Range(32,95).Select(c=>(char)c).ToArray());
        RememberFont("Times New Roman",keyboard,false);
        RememberFont("SimSun",keyboard,false);
        // Resolve names only against installed system families. A name in an
        // input PDF must never be interpreted as a file path or remote URI.
        var installed=new Dictionary<string,FontFamily>(StringComparer.OrdinalIgnoreCase);
        foreach(var family in Fonts.SystemFontFamilies)
            foreach(string alias in family.FamilyNames.Values) installed[alias]=family;
        var fonts=new XElement(W+"fonts");
        var rels=new XElement(Rel+"Relationships");
        var chinese=XmlLanguage.GetLanguage("zh-cn");
        foreach(var entry in usedFonts)
        {
            var font=new XElement(W+"font",At("name",entry.Key));
            FontFamily family;
            if(installed.TryGetValue(entry.Key,out family))
            {
                string alias;
                if(family.FamilyNames.TryGetValue(chinese,out alias) && !String.Equals(alias,entry.Key,StringComparison.OrdinalIgnoreCase))
                    font.Add(Val("altName",alias));
                AddFontPart(font,rels,FontSubset(family,entry.Value.Regular,false),false);
                AddFontPart(font,rels,FontSubset(family,entry.Value.Bold,true),true);
            }
            fonts.Add(font);
        }
        Part(archive,"word/fontTable.xml",fonts);
        relationships.Add(Relationship("fonts","fontTable","fontTable.xml"));
        if(rels.HasElements) Part(archive,"word/_rels/fontTable.xml.rels",rels);
        return rels.HasElements;
    }
}
