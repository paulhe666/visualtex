using System.Text;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

internal static class MathTypeNativeWmfEvidence
{
    internal sealed record Font(string Name,int Weight,bool Italic);
    internal sealed record Run(Font Font,int[] Codes);

    internal static IReadOnlyList<Run> TextRuns(byte[] data)
    {
        var start=data.Length>=22 && BitConverter.ToUInt32(data,0)==0x9AC6CDD7 ? 22 : 0;
        if(data.Length<start+18) throw new InvalidDataException("Truncated WMF.");
        var cursor=start+BitConverter.ToUInt16(data,start+2)*2;
        var slots=new List<(bool Used,Font? Value)>();
        var stack=new Stack<Font?>();
        Font? selected=null;
        var runs=new List<Run>();
        while(cursor+6<=data.Length)
        {
            var size=checked((int)BitConverter.ToUInt32(data,cursor)*2);
            if(size<6 || cursor+size>data.Length) throw new InvalidDataException("Invalid WMF record bounds.");
            var kind=BitConverter.ToUInt16(data,cursor+4);
            if(kind is 0x02FB or 0x02FA or 0x02FC or 0x00F7 or 0x06FF or 0x01F9 or 0x0142)
            {
                Font? font=null;
                if(kind==0x02FB)
                {
                    if(size<24) throw new InvalidDataException("Truncated LOGFONT.");
                    var end=cursor+24; while(end<cursor+size && data[end]!=0) end++;
                    font=new Font(Encoding.ASCII.GetString(data,cursor+24,end-cursor-24),BitConverter.ToInt16(data,cursor+14),data[cursor+16]!=0);
                }
                var slot=slots.FindIndex(x=>!x.Used);
                if(slot<0) slots.Add((true,font)); else slots[slot]=(true,font);
            }
            else if(kind==0x012D)
            {
                var slot=BitConverter.ToUInt16(data,cursor+6);
                if((slot&0x8000)!=0)
                {
                    var stock=slot&0x7FFF;
                    if(stock is 10 or 11 or 12 or 13 or 14 or 16 or 17) selected=new Font("stock-"+stock,stock==13?700:400,false);
                }
                else
                {
                    if(slot>=slots.Count || !slots[slot].Used) throw new InvalidDataException("Selected an unallocated WMF object.");
                    if(slots[slot].Value is not null) selected=slots[slot].Value;
                }
            }
            else if(kind==0x01F0)
            {
                var slot=BitConverter.ToUInt16(data,cursor+6);
                if(slot<slots.Count) slots[slot]=(false,null);
            }
            else if(kind==0x001E) stack.Push(selected);
            else if(kind==0x0127)
            {
                var restore=BitConverter.ToInt16(data,cursor+6);
                var count=restore<0 ? -restore : stack.Count-restore+1;
                while(count-->0 && stack.Count>0) selected=stack.Pop();
            }
            else if(kind is 0x0A32 or 0x0521)
            {
                var count=BitConverter.ToUInt16(data,cursor+(kind==0x0A32?10:6));
                var text=cursor+(kind==0x0A32?14:8);
                if(kind==0x0A32 && (BitConverter.ToUInt16(data,cursor+12)&6)!=0) text+=8;
                if(text+count>cursor+size) throw new InvalidDataException("WMF text exceeds its record.");
                if(count>0) runs.Add(new Run(selected ?? throw new InvalidDataException("A glyph has no selected font."),data.Skip(text).Take(count).Select(b=>(int)b).ToArray()));
            }
            cursor+=size;
            if(kind==0) break;
        }
        return runs;
    }
}

public sealed class MathTypeNativeWmfEvidenceTests
{
    // A declared but unselected font must not satisfy a glyph-style assertion.
    [Fact]
    public void UnselectedFontCannotChangeTheDrawnGlyphStyle()
    {
        static byte[] Record(ushort kind,byte[] payload)
        {
            var bytes=new byte[6+payload.Length+(payload.Length&1)];
            BitConverter.GetBytes((uint)(bytes.Length/2)).CopyTo(bytes,0);
            BitConverter.GetBytes(kind).CopyTo(bytes,4); payload.CopyTo(bytes,6); return bytes;
        }
        static byte[] Font(string name,int weight,bool italic)
        {
            var payload=new byte[18+name.Length+1];
            BitConverter.GetBytes((short)weight).CopyTo(payload,8); payload[10]=(byte)(italic?1:0);
            Encoding.ASCII.GetBytes(name).CopyTo(payload,18); return Record(0x02FB,payload);
        }
        var header=new byte[18]; BitConverter.GetBytes((ushort)9).CopyTo(header,2);
        var bytes=header.Concat(Font("Base",400,false)).Concat(Font("Unused",700,true))
            .Concat(Record(0x012D,new byte[]{0,0}))
            .Concat(Record(0x0521,new byte[]{1,0,(byte)'x',0,0,0,0,0}))
            .Concat(Record(0,new byte[0])).ToArray();
        var run=Assert.Single(MathTypeNativeWmfEvidence.TextRuns(bytes));
        Assert.Equal("Base",run.Font.Name); Assert.Equal(400,run.Font.Weight); Assert.False(run.Font.Italic);
        Assert.Equal(new[]{(int)'x'},run.Codes);
    }
}
