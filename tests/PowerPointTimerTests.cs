using System;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using OpenSlideCharts;

// Original thread-timer integration checks. No Office, PIA, or HWND access.
internal static class PowerPointTimerTests
{
    static int checks;
    static void Check(bool value, string why) { checks++; if (!value) throw new Exception(why); }
    static void Pump(int ms) { var end = DateTime.UtcNow.AddMilliseconds(ms); while(DateTime.UtcNow < end) { Application.DoEvents(); Thread.Sleep(5); } }
    static object Field(object target,string name) { return target.GetType().GetField(name,BindingFlags.Instance|BindingFlags.NonPublic).GetValue(target); }
    [STAThread] static void Main()
    {
        object host = new object(); // Deliberately has no HWND/ActiveWindow properties.
        int ticks=0, owner=Thread.CurrentThread.ManagedThreadId;
        bool threadCorrect=true;
        using(var timer = new PowerPointTimer(host, delegate { ticks++; threadCorrect &= Thread.CurrentThread.ManagedThreadId == owner; }))
        {
            timer.Start(); object first=Field(timer,"id"); timer.Start();
            Check(first.Equals(Field(timer,"id")),"Start not idempotent");
            GC.Collect(); GC.WaitForPendingFinalizers(); Pump(550);
            Check(ticks>=2,"No native thread ticks after GC"); Check(threadCorrect,"Tick left owner thread");
            timer.Stop(); int stopped=ticks; Pump(300); Check(ticks==stopped,"Tick after Stop");
            var dispatch=typeof(PowerPointTimer).GetMethod("Dispatch",BindingFlags.Static|BindingFlags.NonPublic);
            dispatch.Invoke(null,new object[]{IntPtr.Zero,(uint)0x113,first,(uint)0});
            Check(ticks==stopped,"Retired callback delivered without active registration");
            timer.Stop(); Check(((UIntPtr)Field(timer,"id"))==UIntPtr.Zero,"Stop did not release id");
            timer.Start(); Pump(250); Check(ticks>stopped,"Restart did not tick");
            // Windows may reuse first ID. No assertion that restart IDs differ.
            Exception wrongThread=null;
            var thread=new Thread(delegate(){try{timer.Stop();}catch(Exception e){wrongThread=e;}});
            thread.Start(); thread.Join(); Check(wrongThread is InvalidOperationException,"Wrong thread not rejected");
            timer.Dispose(); bool rejected=false;try{timer.Start();}catch(ObjectDisposedException){rejected=true;}
            Check(rejected,"Disposed timer restarted");
        }
        int failures=0, failureTicks=0;
        using(var failing = new PowerPointTimer(host,delegate{failureTicks++;throw new Exception("tick failure");},delegate(Exception e){failures++;throw new Exception("observer failure");}))
        {
            failing.Start(); Pump(650); Check(failureTicks==1,"Failed tick repeated");Check(failures==1,"Observer not called once");
            Check(((UIntPtr)Field(failing,"id"))==UIntPtr.Zero,"Failure did not kill timer");
        }
        int one=0,two=0;
        using(var first=new PowerPointTimer(host,delegate{one++;}))
        using(var second=new PowerPointTimer(host,delegate{two++;}))
        {
            first.Start(); second.Start(); Check(!Field(first,"id").Equals(Field(second,"id")),"Concurrent timers collided");
            Pump(250); Check(one>0&&two>0,"Concurrent timers did not both dispatch");
            first.Stop(); int before=two; Pump(250); Check(two>before,"Stopping one timer stopped another");
        }
        Exception mtaError=null;
        var mta=new Thread(delegate(){try{using(var timer=new PowerPointTimer(host,delegate{})){} }catch(Exception e){mtaError=e;}});
        mta.SetApartmentState(ApartmentState.MTA);mta.Start();mta.Join();
        Check(mtaError is InvalidOperationException,"MTA timer construction accepted");
        Console.WriteLine("PowerPointTimer: " + checks + " thread-timer checks passed. No HWND or Office process used.");
    }
}
