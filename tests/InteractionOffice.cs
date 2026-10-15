using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using OpenSlideCharts;

public static class InteractionOffice
{
    private static int checks;
    private static void Check(bool ok,string message) { checks++; if(!ok) throw new Exception(message); }
    private static void Pump(int milliseconds) { var end=DateTime.UtcNow.AddMilliseconds(milliseconds);while(DateTime.UtcNow<end){Application.DoEvents();Thread.Sleep(20);} }
    private static object Root(dynamic slide) { for(int i=1;i<=(int)slide.Shapes.Count;i++){dynamic s=slide.Shapes.Item(i);if((string)s.Tags.Item("OSC_VERSION")=="1")return s;}throw new Exception("Chart group missing"); }
    private static object NativeConnector(dynamic slide,int index)
    {
        dynamic root=Root(slide);int n=0;
        for(int i=1;i<=(int)root.GroupItems.Count;i++){dynamic s=root.GroupItems.Item(i);if((string)s.Tags.Item("OSC_ROLE")=="WATERFALL_CONNECTOR"&&n++==index)return s;}
        throw new Exception("Native connector missing: "+index);
    }
    private static int Overlays(dynamic slide) {int n=0;for(int i=1;i<=(int)slide.Shapes.Count;i++)if((string)slide.Shapes.Item(i).Tags.Item("OSC_EDIT")!="")n++;return n;}
    private static void MoveEnd(dynamic shape,double x1,double y1,double x2,double y2)
    {
        shape.Left=(float)Math.Min(x1,x2);shape.Top=(float)Math.Min(y1,y2);shape.Width=(float)Math.Abs(x2-x1);shape.Height=(float)Math.Abs(y2-y1);
        bool flipX=(int)shape.HorizontalFlip==-1,flipY=(int)shape.VerticalFlip==-1;
        if(flipX!=(x2<x1))shape.Flip(0);if(flipY!=(y2<y1))shape.Flip(1);
    }
    private static ChartSpec LeftExample()
    {
        var s=new ChartSpec{Kind="Waterfall",Title="Edition du total à gauche",Waterfall=new WaterfallSpec()};
        double?[] values={null,3,2,4};for(int i=0;i<values.Length;i++){var c=new WaterfallColumn{Id="C"+i,Label="Colonne "+i};c.Segments.Add(new WaterfallSegment{Id="S"+i,SheetRow=0,Calculated=!values[i].HasValue,Value=values[i]});s.Waterfall.Columns.Add(c);}WaterfallProfiles.Arrange(s.Waterfall,"EndAtZero");return s;
    }
    [STAThread] public static int Main(string[] args)
    {
        dynamic app=null,deck=null;PowerPointEvents external=null;
        try
        {
            Directory.CreateDirectory(args[0]);app=Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application",true));app.Visible=-1;deck=app.Presentations.Add(-1);deck.PageSetup.SlideWidth=960f;deck.PageSetup.SlideHeight=540f;Pump(1000);
            app.COMAddIns.Update();dynamic addin=app.COMAddIns.Item("OpenSlideCharts.Connect");if(!(bool)addin.Connect)addin.Connect=true;dynamic api=addin.Object;
            string status=api.DiagnosticStatus();Console.WriteLine(status);Check(status.Contains("events=True"),"PowerPoint event registration failed.");
            int selected=0,saving=0,closing=0;
            external=new PowerPointEvents((object)app,delegate(object sel){selected++;},null,delegate(object pres,ref bool cancel){saving++;},delegate(object pres,ref bool cancel){closing++;});
            dynamic slide=deck.Slides.Add(1,12);var spec=LeftExample();dynamic group=OfficeRenderer.Render((object)slide,spec,30,35,900,470);group.Select();Pump(100);
            Check(selected>0,"Selection event was not dispatched over COM.");
            int connectorItem=-1;for(int i=1;i<=(int)group.GroupItems.Count;i++){dynamic child=group.GroupItems.Item(i);if((string)child.Tags.Item("OSC_ROLE")=="WATERFALL_CONNECTOR"){connectorItem=i;break;}}
            Check(connectorItem>0,"Missing primitive role tags.");group.GroupItems.Item(connectorItem).Select();Pump(100);
            var selection=OfficeChartSelection.Resolve((object)app.ActiveWindow.Selection);Check(selection!=null&&selection.Primitive!=null&&selection.Primitive.EntityId==spec.Waterfall.Connectors[0].Id,"Child selection did not resolve its owning chart/connector.");
            string file=Path.Combine(Path.GetFullPath(args[0]),"InteractiveWaterfall.pptx");deck.SaveAs(file,24);api.EditConnectors(null);Check(Overlays(slide)==0&&(int)slide.Shapes.Count==1,"Entering native editing added document helper shapes.");
            dynamic line=NativeConnector(slide,0);ConnectorEndpoints points=ConnectorEndpoints.Read((object)line);var scene=WaterfallScene.Build(spec.Waterfall,900,470);
            var target=scene.Targets.Find(delegate(WaterfallTarget t){return t.ColumnId=="C3"&&t.Edge=="Output"&&t.Side=="Left";});
            MoveEnd(line,points.X1,points.Y1,30+target.X,35+target.Y);api.ApplyConnectorMove(null);Pump(200);
            ChartSpec changed=ChartStorage.Read(Root(slide));Check(WaterfallSolver.Solve(changed.Waterfall).Columns[0].Output==9,"Dragging endpoint did not change initial e from -9 to +9.");
            Check(changed.Waterfall.Columns[0].Segments[0].Calculated&&!changed.Waterfall.Columns[0].Segments[0].Value.HasValue,"Interactive edit flattened e into a cached number.");
            Check(changed.Waterfall.Pins.Exists(delegate(WaterfallPin p){return p.AutoGenerated&&p.Anchor.ColumnId=="C1";}),"Detached numeric sum was not grounded.");
            Check(!changed.Waterfall.Pins.Exists(delegate(WaterfallPin p){return p.AutoGenerated&&p.Anchor.ColumnId=="C3";}),"Obsolete automatic output pin was retained.");
            line=NativeConnector(slide,0);points=ConnectorEndpoints.Read((object)line);scene=WaterfallScene.Build(changed.Waterfall,900,470);
            target=scene.Targets.Find(delegate(WaterfallTarget t){return t.ColumnId=="C2"&&t.Edge=="Output"&&t.Side=="Left";});
            MoveEnd(line,points.X1,points.Y1,30+target.X,35+target.Y);Pump(1800);
            changed=ChartStorage.Read(Root(slide));Console.WriteLine(api.DiagnosticStatus());Check(WaterfallSolver.Solve(changed.Waterfall).Columns[0].Output==5,"Stable endpoint change was not applied automatically by the in-process timer.");
            line=NativeConnector(slide,0);points=ConnectorEndpoints.Read((object)line);scene=WaterfallScene.Build(changed.Waterfall,900,470);
            target=scene.Targets.Find(delegate(WaterfallTarget t){return t.ColumnId=="C3"&&t.Edge=="Output"&&t.Side=="Left";});
            MoveEnd(line,points.X1,points.Y1,30+target.X,35+target.Y);api.ApplyConnectorMove(null);Pump(100);
            Check(WaterfallSolver.Solve(ChartStorage.Read(Root(slide)).Waterfall).Columns[0].Output==9,"Moving endpoint back lost the explicit graph.");
            string copy=Path.Combine(Path.GetFullPath(args[0]),"InteractiveWaterfall-copy.pptx");deck.SaveCopyAs(copy,24);Check((int)slide.Shapes.Count==1&&Overlays(slide)==0,"SaveCopyAs editing introduced helper shapes.");
            // Save while editing: the application event flushes the current model.
            deck.Save();Pump(100);Check(saving>0,"BeforeSave event missing.");Check(Overlays(slide)==0,"Temporary editing lines survived BeforeSave.");Check(((string)api.DiagnosticStatus()).Contains("connector-editing=False"),"Editor did not end before saving.");
            slide.Export(Path.Combine(Path.GetFullPath(args[0]),"InteractiveWaterfall.png"),"PNG",1440,810);
            deck.Close();Marshal.FinalReleaseComObject((object)deck);deck=null;Check(closing>0,"BeforeClose event missing.");
            deck=app.Presentations.Open(file,0,0,-1);slide=deck.Slides.Item(1);Check((int)slide.Shapes.Count==1&&Overlays(slide)==0,"Saved deck contains helper shapes.");Check(WaterfallSolver.Solve(ChartStorage.Read(Root(slide)).Waterfall).Columns[0].Output==9,"Reattached constraint did not survive reopen.");
            // Native delete of an editing line starts an independent sum.
            spec=WaterfallProfiles.Sample();dynamic second=deck.Slides.Add(2,12);group=OfficeRenderer.Render((object)second,spec,30,35,900,470);app.ActiveWindow.View.GotoSlide(2);group.Select();api.EditConnectors(null);dynamic removed=NativeConnector(second,0);removed.Select();removed.Delete();Pump(1200);api.ApplyConnectorMove(null);Pump(100);
            changed=ChartStorage.Read(Root(second));Check(changed.Waterfall.Connectors.Count==4,"Delete did not remove the model connector.");Check(WaterfallSolver.Solve(changed.Waterfall).Columns[3].Output==6,"Delete did not start a new summation.");api.EndConnectorEditing(null);Check(Overlays(second)==0,"Ending edit retained helpers.");
            // Navigation must flush the old slide without stealing selection back.
            app.ActiveWindow.View.GotoSlide(1);group=Root(slide);group.Select();api.EditConnectors(null);
            changed=ChartStorage.Read(Root(slide));line=NativeConnector(slide,0);points=ConnectorEndpoints.Read((object)line);scene=WaterfallScene.Build(changed.Waterfall,900,470);
            target=scene.Targets.Find(delegate(WaterfallTarget t){return t.ColumnId=="C2"&&t.Edge=="Output"&&t.Side=="Left";});
            MoveEnd(line,points.X1,points.Y1,30+target.X,35+target.Y);app.ActiveWindow.View.GotoSlide(2);Pump(1000);
            Check(WaterfallSolver.Solve(ChartStorage.Read(Root(slide)).Waterfall).Columns[0].Output==5,"Navigating before debounce lost the endpoint edit.");
            Check((int)app.ActiveWindow.View.Slide.SlideID==(int)second.SlideID,"Flushing the old slide changed the active slide.");
            Check(((string)api.DiagnosticStatus()).Contains("connector-editing=False"),"Navigation left a stale editing session.");
            // Native deletion of the complete group or slide must release the session.
            group=Root(second);group.Select();api.EditConnectors(null);group.Delete();Pump(800);
            Check(((string)api.DiagnosticStatus()).Contains("connector-editing=False"),"Deleting a chart left the editor locked.");
            group=OfficeRenderer.Render((object)second,LeftExample(),30,35,900,470);group.Select();api.EditConnectors(null);
            Check(((string)api.DiagnosticStatus()).Contains("connector-editing=True"),"A fresh chart could not start editing after deletion.");
            second.Delete();Pump(800);
            Check(((string)api.DiagnosticStatus()).Contains("connector-editing=False"),"Deleting a slide left the editor locked.");
            // Charts saved by older builds have no semantic child tags.
            second=deck.Slides.Add(2,12);app.ActiveWindow.View.GotoSlide(2);group=OfficeRenderer.Render((object)second,LeftExample(),30,35,900,470);
            for(int i=1;i<=(int)group.GroupItems.Count;i++){group.GroupItems.Item(i).Tags.Delete("OSC_ROLE");group.GroupItems.Item(i).Tags.Delete("OSC_ITEM");}
            group.Select();api.EditConnectors(null);Pump(300);
            Check(((string)api.DiagnosticStatus()).Contains("connector-editing=True"),"Upgrading an old chart reentered and cleared the session.");
            Check(NativeConnector(second,0)!=null&&(int)second.Shapes.Count==1,"Upgrade did not regenerate native connectors transactionally.");
            api.EndConnectorEditing(null);
            Console.WriteLine("PASS: "+checks+" Office interaction assertions; COM events, child selection, endpoint change, automatic grounding, save/reopen, native deletion without helper shapes.");return 0;
        }
        catch(Exception error){Console.Error.WriteLine(error);return 1;}
        finally{if(external!=null)try{external.Dispose();}catch{}if((object)deck!=null)try{deck.Saved=-1;deck.Close();}catch{}if((object)app!=null)try{if((int)app.Presentations.Count==0)app.Quit();}catch{}if((object)deck!=null)try{Marshal.FinalReleaseComObject((object)deck);}catch{}if((object)app!=null)try{Marshal.FinalReleaseComObject((object)app);}catch{}}
    }
}
