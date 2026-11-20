using System.Diagnostics;
using System.Runtime.InteropServices;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class OfficeDoubleClickGestureTests
{
    private static readonly IntPtr Page = new(0x1000);
    private static readonly IntPtr OtherPage = new(0x2000);
    private static readonly IntPtr Dialog = new(0x3000);

    private static bool Press(
        OfficeDoubleClickPairing pairing,
        IntPtr window,
        bool onPage,
        int x,
        int y,
        long milliseconds) =>
        pairing.Register(
            window,
            onPage,
            x,
            y,
            milliseconds,
            doubleClickMilliseconds: 500,
            maxDeltaX: 4,
            maxDeltaY: 4);

    [Fact]
    public void TwoQuickPressesOnThePageFormADoubleClick()
    {
        var pairing = new OfficeDoubleClickPairing();

        Assert.False(Press(pairing, Page, true, 100, 100, 1000));
        Assert.True(Press(pairing, Page, true, 102, 101, 1300));
    }

    [Fact]
    public void QuickPressesInADialogOverAFormulaNeverPair()
    {
        // Reported case: fast clicks while setting paragraph styles, in a Word
        // dialog that happened to lie over a formula.
        var pairing = new OfficeDoubleClickPairing();

        Assert.False(Press(pairing, Dialog, false, 100, 100, 1000));
        Assert.False(Press(pairing, Dialog, false, 100, 100, 1100));
    }

    [Fact]
    public void ADialogPressThenAPagePressIsNotADoubleClick()
    {
        // The first press closes the dialog (e.g. OK); the second lands on the
        // formula that was underneath. Windows does not pair these either.
        var pairing = new OfficeDoubleClickPairing();

        Assert.False(Press(pairing, Dialog, false, 100, 100, 1000));
        Assert.False(Press(pairing, Page, true, 100, 100, 1150));
        // A real double-click on the page right afterwards still works.
        Assert.True(Press(pairing, Page, true, 100, 100, 1300));
    }

    [Fact]
    public void APagePressThenAnOverlayPressIsNotADoubleClick()
    {
        var pairing = new OfficeDoubleClickPairing();

        Assert.False(Press(pairing, Page, true, 100, 100, 1000));
        Assert.False(Press(pairing, Dialog, false, 100, 100, 1150));
        Assert.False(Press(pairing, Page, true, 100, 100, 1300));
    }

    [Fact]
    public void PressesOnTwoDifferentPageWindowsNeverPair()
    {
        var pairing = new OfficeDoubleClickPairing();

        Assert.False(Press(pairing, Page, true, 100, 100, 1000));
        Assert.False(Press(pairing, OtherPage, true, 100, 100, 1100));
    }

    [Fact]
    public void SlowOrDistantPressesDoNotPair()
    {
        var slow = new OfficeDoubleClickPairing();
        Assert.False(Press(slow, Page, true, 100, 100, 1000));
        Assert.False(Press(slow, Page, true, 100, 100, 1501));

        var distant = new OfficeDoubleClickPairing();
        Assert.False(Press(distant, Page, true, 100, 100, 1000));
        Assert.False(Press(distant, Page, true, 105, 100, 1100));
    }

    [Fact]
    public void AThirdQuickPressStartsANewGesture()
    {
        var pairing = new OfficeDoubleClickPairing();

        Assert.False(Press(pairing, Page, true, 100, 100, 1000));
        Assert.True(Press(pairing, Page, true, 100, 100, 1100));
        Assert.False(Press(pairing, Page, true, 100, 100, 1200));
        Assert.True(Press(pairing, Page, true, 100, 100, 1300));
    }

    [Fact]
    public void ResetForgetsThePreviousPress()
    {
        var pairing = new OfficeDoubleClickPairing();

        Assert.False(Press(pairing, Page, true, 100, 100, 1000));
        pairing.Reset();
        Assert.False(Press(pairing, Page, true, 100, 100, 1100));
    }

    [Fact]
    public void ExtremeCoordinatesDoNotOverflow()
    {
        var pairing = new OfficeDoubleClickPairing();

        Assert.False(Press(pairing, Page, true, int.MinValue, int.MinValue, 1000));
        Assert.False(Press(pairing, Page, true, int.MaxValue, int.MaxValue, 1100));
    }

    [Fact]
    public void NowMillisecondsAdvancesInMilliseconds()
    {
        var first = OfficeDoubleClickPairing.NowMilliseconds();
        Thread.Sleep(30);
        var second = OfficeDoubleClickPairing.NowMilliseconds();

        Assert.InRange(second - first, 20, 5000);
    }

    [Fact]
    public void OnlyAVisibleEnabledPageWindowOfTheOwnerProcessIsThePage()
    {
        // Real windows, created far off-screen and never activated.
        const int left = -20000;
        const int top = -20000;
        var frame = CreateWindowEx(
            WsExToolWindow | WsExNoActivate,
            "STATIC",
            "VisualTeX double-click test frame",
            WsPopup,
            left,
            top,
            400,
            300,
            IntPtr.Zero,
            IntPtr.Zero,
            IntPtr.Zero,
            IntPtr.Zero);
        Assert.NotEqual(IntPtr.Zero, frame);
        try
        {
            var page = CreateWindowEx(
                0,
                "STATIC",
                string.Empty,
                WsChild | WsVisible,
                50,
                40,
                200,
                100,
                frame,
                IntPtr.Zero,
                IntPtr.Zero,
                IntPtr.Zero);
            Assert.NotEqual(IntPtr.Zero, page);
            ShowWindow(frame, SwShowNoActivate);

            var owner = Process.GetCurrentProcess().Id;
            var pageClasses = new[] { "Static" };

            Assert.True(OfficeDocumentSurface.IsPage(page, pageClasses, owner));
            // Dialogs, galleries and panes have other window classes.
            Assert.False(OfficeDocumentSurface.IsPage(page, OfficeDocumentSurface.WordPageClasses, owner));
            Assert.False(OfficeDocumentSurface.IsPage(page, OfficeDocumentSurface.PowerPointPageClasses, owner));
            // Another process's window.
            Assert.False(OfficeDocumentSurface.IsPage(page, pageClasses, owner + 1));
            Assert.False(OfficeDocumentSurface.IsPage(IntPtr.Zero, pageClasses, owner));

            // A modal dialog disables the document's frame.
            EnableWindow(frame, false);
            Assert.False(OfficeDocumentSurface.IsPage(page, pageClasses, owner));
            EnableWindow(frame, true);
            Assert.True(OfficeDocumentSurface.IsPage(page, pageClasses, owner));

            ShowWindow(page, SwHide);
            Assert.False(OfficeDocumentSurface.IsPage(page, pageClasses, owner));
        }
        finally
        {
            DestroyWindow(frame);
        }
    }

    [Fact]
    public void APressReachesThePageDirectlyOrThroughAClickThroughWindow()
    {
        // Real windows, far off-screen and never activated: a frame (Button)
        // holding a page (Static), with popups (Edit) lying over it.
        const int left = -24000;
        const int top = -24000;
        var frame = CreateWindowEx(WsExToolWindow | WsExNoActivate, "BUTTON", "VisualTeX press test frame",
            WsPopup, left, top, 400, 300, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
        Assert.NotEqual(IntPtr.Zero, frame);
        var dialog = IntPtr.Zero;
        var clickThrough = IntPtr.Zero;
        try
        {
            var page = CreateWindowEx(0, "STATIC", string.Empty, WsChild | WsVisible,
                50, 40, 200, 100, frame, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
            Assert.NotEqual(IntPtr.Zero, page);
            ShowWindow(frame, SwShowNoActivate);
            dialog = CreateWindowEx(WsExToolWindow | WsExNoActivate, "EDIT", "dialog",
                WsPopup, left, top, 400, 300, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
            clickThrough = CreateWindowEx(
                WsExToolWindow | WsExNoActivate | WsExLayered | WsExTransparent, "EDIT", "overlay",
                WsPopup, left, top, 400, 300, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
            Assert.NotEqual(IntPtr.Zero, dialog);
            Assert.NotEqual(IntPtr.Zero, clickThrough);

            var owner = Process.GetCurrentProcess().Id;
            var pageClasses = new[] { "Static" };
            const int onPageX = left + 60;
            const int onPageY = top + 50;
            const int offPageX = left + 5;
            const int offPageY = top + 5;

            var direct = OfficeDocumentSurface.ResolvePress(page, pageClasses, owner, onPageX, onPageY);
            Assert.True(direct.OnPage, direct.Detail);
            Assert.Equal(page, direct.Page);

            // A dialog over the page takes the press itself.
            Raise(dialog);
            var onDialog = OfficeDocumentSurface.ResolvePress(dialog, pageClasses, owner, onPageX, onPageY);
            Assert.False(onDialog.OnPage);
            Assert.Contains("class=Edit", onDialog.Detail);
            ShowWindow(dialog, SwHide);

            // A click-through window passes the press to the page below it.
            Raise(clickThrough);
            var through = OfficeDocumentSurface.ResolvePress(clickThrough, pageClasses, owner, onPageX, onPageY);
            Assert.True(through.OnPage, through.Detail);
            Assert.Equal(page, through.Page);
            Assert.Contains("through=Edit", through.Detail);
            // ...and to the frame where there is no page.
            var throughToFrame = OfficeDocumentSurface.ResolvePress(clickThrough, pageClasses, owner, offPageX, offPageY);
            Assert.False(throughToFrame.OnPage);
            Assert.Contains("class=Button", throughToFrame.Detail);
            // The dialog stacked between them still takes the press.
            Raise(dialog);
            Raise(clickThrough);
            Assert.False(OfficeDocumentSurface.ResolvePress(clickThrough, pageClasses, owner, onPageX, onPageY).OnPage);
            ShowWindow(dialog, SwHide);

            // Another program's window takes the press itself.
            var foreign = OfficeDocumentSurface.ResolvePress(page, pageClasses, owner + 1, onPageX, onPageY);
            Assert.False(foreign.OnPage);
            Assert.Contains("other-process", foreign.Detail);
            Assert.False(OfficeDocumentSurface.ResolvePress(IntPtr.Zero, pageClasses, owner, onPageX, onPageY).OnPage);

            // A modal dialog disables the frame.
            EnableWindow(frame, false);
            var disabled = OfficeDocumentSurface.ResolvePress(page, pageClasses, owner, onPageX, onPageY);
            Assert.False(disabled.OnPage);
            Assert.Contains("page-unavailable", disabled.Detail);
            EnableWindow(frame, true);
        }
        finally
        {
            if (clickThrough != IntPtr.Zero) DestroyWindow(clickThrough);
            if (dialog != IntPtr.Zero) DestroyWindow(dialog);
            DestroyWindow(frame);
        }
    }

    [Fact]
    public void TheHookThreadUsesPhysicalCoordinates()
    {
        var aware = false;
        var thread = new Thread(() => aware = OfficeDocumentSurface.UsePhysicalCoordinatesOnThisThread());
        thread.Start();
        thread.Join();
        Assert.True(aware);
    }

    private static void Raise(IntPtr window)
    {
        ShowWindow(window, SwShowNoActivate);
        SetWindowPos(window, IntPtr.Zero, 0, 0, 0, 0, SwpNoMove | SwpNoSize | SwpNoActivate);
    }

    private const uint WsExLayered = 0x00080000;
    private const uint WsExTransparent = 0x00000020;
    private const uint SwpNoSize = 0x0001;
    private const uint SwpNoMove = 0x0002;
    private const uint SwpNoActivate = 0x0010;

    [DllImport("user32.dll")]
    private static extern bool SetWindowPos(IntPtr window, IntPtr insertAfter, int x, int y, int cx, int cy, uint flags);

    private const uint WsPopup = 0x80000000;
    private const uint WsChild = 0x40000000;
    private const uint WsVisible = 0x10000000;
    private const uint WsExToolWindow = 0x00000080;
    private const uint WsExNoActivate = 0x08000000;
    private const int SwHide = 0;
    private const int SwShowNoActivate = 4;

    [DllImport("user32.dll", EntryPoint = "CreateWindowExW", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr CreateWindowEx(
        uint exStyle,
        string className,
        string windowName,
        uint style,
        int x,
        int y,
        int width,
        int height,
        IntPtr parent,
        IntPtr menu,
        IntPtr instance,
        IntPtr parameter);

    [DllImport("user32.dll")]
    private static extern bool DestroyWindow(IntPtr window);

    [DllImport("user32.dll")]
    private static extern bool EnableWindow(IntPtr window, bool enable);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr window, int command);
}
