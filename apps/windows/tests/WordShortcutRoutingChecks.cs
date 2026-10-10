using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using Word = Microsoft.Office.Interop.Word;

// Managed doubles and simulated routing calls. No Word process, hook install,
// registry write, template write or physical input is performed.
public static class WordShortcutRoutingChecks
{
    public sealed class Result { public string Case; public bool Passed; public string Detail; }
    private sealed class Double : RealProxy
    {
        private readonly Func<MethodInfo, object[], object> handler;
        internal Double(Type type, Func<MethodInfo, object[], object> invoke) : base(type) { handler = invoke; }
        public override IMessage Invoke(IMessage message)
        {
            var call = (IMethodCallMessage)message;
            try { return new ReturnMessage(handler((MethodInfo)call.MethodBase, call.Args),
                call.Args, call.ArgCount, call.LogicalCallContext, call); }
            catch (Exception error) { return new ReturnMessage(error, call); }
        }
    }
    private static Type shortcutType, commandType, hookType, policyType;
    private static int dispatched, pasted, expectedCode, lookups;
    private static bool accept, fail, previousSeen;
    private static Word.WdKeyCategory currentCategory, previousCategory;
    private static readonly List<string> unexpected = new List<string>();
    private static object Shortcut(uint key, bool ctrl, bool alt = false, bool shift = false)
    { return Activator.CreateInstance(shortcutType, BindingFlags.Instance | BindingFlags.NonPublic, null, new object[]{key,ctrl,alt,shift}, null); }
    public static bool Dispatch(object shortcut, object previous, object command)
    {
        dispatched++; previousSeen = previous != null;
        if (fail) throw new InvalidOperationException("Controlled rejection");
        return accept;
    }
    public static void Paste() { pasted++; }
    public static void DirectAction(object command) { dispatched++; }
    private static object Hook(bool actionCompatibility)
    {
        if (actionCompatibility) {
            var actionType=typeof(Action<>).MakeGenericType(commandType);
            var parameter=Expression.Parameter(commandType,"command");
            var action=Expression.Lambda(actionType,Expression.Call(typeof(WordShortcutRoutingChecks).GetMethod("DirectAction"),
                Expression.Convert(parameter,typeof(object))),parameter).Compile();
            var host=hookType.Assembly.GetType("VisualTeX.WindowsOffice.VstoShared.OfficeFormulaShortcutHost");
            return hookType.GetConstructor(BindingFlags.Instance|BindingFlags.NonPublic,null,
                new[]{host,actionType,typeof(Action),typeof(Action<string>)},null).Invoke(new object[]{Enum.ToObject(host,1),action,null,null});
        }
        var nullable = typeof(Nullable<>).MakeGenericType(shortcutType);
        Type delegateType = typeof(Func<,,,>).MakeGenericType(shortcutType, nullable, commandType, typeof(bool));
        var a = Expression.Parameter(shortcutType, "a");
        var b = Expression.Parameter(nullable, "b");
        var c = Expression.Parameter(commandType, "c");
        var callback = Expression.Lambda(delegateType, Expression.Call(typeof(WordShortcutRoutingChecks).GetMethod("Dispatch"),
            Expression.Convert(a,typeof(object)),Expression.Convert(b,typeof(object)),Expression.Convert(c,typeof(object))),a,b,c).Compile();
        var hostType = hookType.Assembly.GetType("VisualTeX.WindowsOffice.VstoShared.OfficeFormulaShortcutHost");
        var constructor = hookType.GetConstructor(BindingFlags.Instance|BindingFlags.NonPublic,null,
            new[]{hostType, delegateType, typeof(Action), typeof(Action<string>)},null);
        return constructor.Invoke(new object[]{Enum.ToObject(hostType,0),callback,(Action)Paste,null});
    }
    private static object Command()
    {
        var kind = hookType.Assembly.GetType("VisualTeX.WindowsOffice.VstoShared.OfficeFormulaShortcutKind");
        var format = hookType.Assembly.GetType("VisualTeX.WindowsOffice.VstoShared.OfficeFormulaShortcutFormat");
        return Activator.CreateInstance(commandType, BindingFlags.Instance|BindingFlags.NonPublic, null,
            new[]{Enum.ToObject(kind,0),Enum.ToObject(format,0)},null);
    }
    private static bool Route(object hook, uint key, bool ctrl, bool alt, bool shift, bool win, bool altGr, int focus, bool match)
    { return (bool)hookType.GetMethod("ProcessKeyDown",BindingFlags.Instance|BindingFlags.NonPublic).Invoke(hook,
        new object[]{key,ctrl,alt,shift,win,altGr,new IntPtr(focus),match?Command():null}); }
    private static bool Policy(uint key, bool ctrl, bool alt, bool shift, bool focus, object previous)
    {
        lookups=0; unexpected.Clear();
        Word.Application app=(Word.Application)new Double(typeof(Word.Application),delegate(MethodInfo method, object[] args) {
            if(method.Name=="BuildKeyCode") {
                int code=Convert.ToInt32(args[0]);
                for(int i=1;i<args.Length;i++) if(args[i]!=Type.Missing) code += Convert.ToInt32(args[i]);
                return code;
            }
            if(method.Name=="get_FindKey") {
                lookups++;
                if(fail) throw new InvalidOperationException("Controlled COM lookup failure");
                int code=Convert.ToInt32(args[0]);
                var category=lookups==1?currentCategory:previousCategory;
                if(lookups==1 && code!=expectedCode) unexpected.Add("Incorrect Word key code: "+code);
                return (Word.KeyBinding)new Double(typeof(Word.KeyBinding),delegate(MethodInfo m,object[] ignored){
                    if(m.Name=="get_KeyCategory")return category;
                    if(m.Name=="get_Command")return "TestCommand";
                    unexpected.Add(m.Name);throw new InvalidOperationException(m.Name);
                }).GetTransparentProxy();
            }
            unexpected.Add(method.Name);throw new InvalidOperationException(method.Name);
        }).GetTransparentProxy();
        lastConflict=null;
        var arguments=new[]{app,Shortcut(key,ctrl,alt,shift),previous,(object)focus,null};
        var withConflict=Array.Find(policyType.GetMethods(BindingFlags.Static|BindingFlags.NonPublic),
            m=>m.Name=="CanExecuteInContext"&&m.GetParameters().Length==5);
        var allowed=(bool)withConflict.Invoke(null,arguments);
        lastConflict=arguments[4];
        return allowed;
    }
    private static object lastConflict;
    private static object ConflictValue(string property)
    { return lastConflict==null?null:lastConflict.GetType().GetProperty(property,BindingFlags.Instance|BindingFlags.NonPublic).GetValue(lastConflict); }
    private static void Check(List<Result> results,string name,bool condition)
    { results.Add(new Result{Case=name,Passed=condition,Detail=condition?"":"Boundary assertion failed"}); }
    public static Result[] Run(string assemblyPath)
    {
        var assembly=Assembly.LoadFrom(assemblyPath);
        shortcutType=assembly.GetType("VisualTeX.WindowsOffice.VstoShared.OfficeFormulaShortcut");
        commandType=assembly.GetType("VisualTeX.WindowsOffice.VstoShared.OfficeFormulaShortcutCommand");
        hookType=assembly.GetType("VisualTeX.WindowsOffice.VstoShared.OfficeFormulaShortcutKeyboardHook");
        policyType=assembly.GetType("VisualTeX.WordVsto.WordFormulaShortcutPolicy");
        var results=new List<Result>();
        fail=false;currentCategory=Word.WdKeyCategory.wdKeyCategoryNil;previousCategory=Word.WdKeyCategory.wdKeyCategoryNil;
        expectedCode=0x44+512;
        Check(results,"unassigned Ctrl+D can launch",Policy(0x44,true,false,false,true,null)&&unexpected.Count==0);
        foreach(Word.WdKeyCategory category in Enum.GetValues(typeof(Word.WdKeyCategory))) {
            if(category==Word.WdKeyCategory.wdKeyCategoryNil)continue;
            currentCategory=category;
            Check(results,"Word category "+category+" has priority",!Policy(0x44,true,false,false,true,null)&&unexpected.Count==0);
            Check(results,"Word category "+category+" is reported",
                Equals(ConflictValue("Category"),category)&&(string)ConflictValue("Command")=="TestCommand");
        }
        currentCategory=Word.WdKeyCategory.wdKeyCategoryNil;
        Check(results,"unassigned chord reports no conflict",Policy(0x44,true,false,false,true,null)&&lastConflict==null);
        Check(results,"dialog focus reports no conflict",!Policy(0x44,true,false,false,false,null)&&lastConflict==null);
        currentCategory=Word.WdKeyCategory.wdKeyCategoryNil;
        previousCategory=Word.WdKeyCategory.wdKeyCategoryPrefix;
        Check(results,"second stroke after Word prefix passes through",!Policy(0x44,true,false,false,true,Shortcut(0x4B,true))&&lookups==2);
        previousCategory=Word.WdKeyCategory.wdKeyCategoryNil;
        Check(results,"ordinary preceding key does not block",Policy(0x44,true,false,false,true,Shortcut(0x41,false)));
        Check(results,"dialog/search/editor focus does no COM lookup",!Policy(0x44,true,false,false,false,null)&&lookups==0);
        Check(results,"Alt-only menu accelerators pass through",!Policy(0x46,false,true,false,true,null)&&lookups==0);
        Check(results,"unmapped F24 passes through",!Policy(0x87,true,false,false,true,null)&&lookups==0);
        expectedCode=0x7B+512+1024+256;
        Check(results,"Ctrl+Alt+Shift+F12 maps Word modifiers",Policy(0x7B,true,true,true,true,null)&&unexpected.Count==0);
        fail=true;expectedCode=0x44+512;
        Check(results,"COM query exception passes through",!Policy(0x44,true,false,false,true,null));fail=false;
        accept=true;dispatched=0;pasted=0;
        var hook=Hook(false);
        Check(results,"accepted configured chord consumes once",Route(hook,0x44,true,false,false,false,false,1,true)&&dispatched==1);
        accept=false;
        Check(results,"host rejection passes through",!Route(hook,0x44,true,false,false,false,false,1,true));
        fail=true;
        Check(results,"callback exception passes through",!Route(hook,0x44,true,false,false,false,false,1,true));fail=false;
        int before=dispatched;accept=true;
        Check(results,"unconfigured ordinary key passes through",!Route(hook,0x41,false,false,false,false,false,1,false)&&dispatched==before);
        Check(results,"AltGr never launches",!Route(hook,0x44,true,true,false,false,true,1,true)&&dispatched==before);
        Check(results,"Windows chord never launches",!Route(hook,0x44,true,false,false,true,false,1,true)&&dispatched==before);
        accept=false;pasted=0;
        Check(results,"denied Ctrl+V retains native paste observation",!Route(hook,0x56,true,false,false,false,false,1,true)&&pasted==1);
        Check(results,"Ctrl+Alt+V is not ordinary paste",!Route(hook,0x56,true,true,false,false,false,1,false)&&pasted==1);
        Route(hook,0x4B,true,false,false,false,false,1,false);
        previousSeen=false;
        Route(hook,0x44,true,false,false,false,false,1,true);
        Check(results,"same-focus preceding chord reaches policy",previousSeen);
        previousSeen=true;
        Route(hook,0x44,true,false,false,false,false,2,true);
        Check(results,"focus change clears prefix context",!previousSeen);
        var eligible=hookType.GetMethod("IsActionKeyDown",BindingFlags.NonPublic|BindingFlags.Static);
        Check(results,"HC_ACTION initial down eligible",(bool)eligible.Invoke(null,new object[]{0,0L}));
        Check(results,"HC_NOREMOVE never executes",!(bool)eligible.Invoke(null,new object[]{3,0L}));
        Check(results,"negative hook code never executes",!(bool)eligible.Invoke(null,new object[]{-1,0L}));
        Check(results,"repeat never duplicates launch",!(bool)eligible.Invoke(null,new object[]{0,0x40000000L}));
        Check(results,"key-up never duplicates launch",!(bool)eligible.Invoke(null,new object[]{0,0x80000000L}));
        dispatched=0;
        Check(results,"PowerPoint Action constructor remains compatible",Route(Hook(true),0x44,true,false,false,false,false,1,true)&&dispatched==1);
        var addinType=assembly.GetType("VisualTeX.WordVsto.ThisAddIn");
        var addin=System.Runtime.Serialization.FormatterServices.GetUninitializedObject(addinType);
        var begin=addinType.GetMethod("BeginSession",BindingFlags.NonPublic|BindingFlags.Instance);
        var arguments=new object[]{"create","inline",null,null,false};
        Check(results,"missing add-in lifetime rejects launch",!(bool)begin.Invoke(addin,arguments));
        var lifetime=new System.Threading.CancellationTokenSource();
        addinType.GetField("_lifetime",BindingFlags.NonPublic|BindingFlags.Instance).SetValue(addin,lifetime);
        Check(results,"create-selection failure rejects launch",!(bool)begin.Invoke(addin,arguments));
        lifetime.Cancel();
        Check(results,"cancelled lifetime rejects launch",!(bool)begin.Invoke(addin,arguments));
        lifetime.Dispose();
        return results.ToArray();
    }
}
