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
            var insideX = left + 60;
            var insideY = top + 50;

            Assert.True(OfficeDocumentSurface.IsPage(page, pageClasses, owner, insideX, insideY));
            // A point outside the page's client area.
            Assert.False(OfficeDocumentSurface.IsPage(page, pageClasses, owner, left + 5, top + 5));
            // Dialogs, galleries and panes have other window classes.
            Assert.False(OfficeDocumentSurface.IsPage(
                page,
                OfficeDocumentSurface.WordPageClasses,
                owner,
                insideX,
                insideY));
            Assert.False(OfficeDocumentSurface.IsPage(
                page,
                OfficeDocumentSurface.PowerPointPageClasses,
                owner,
                insideX,
                insideY));
            // Another process's window.
            Assert.False(OfficeDocumentSurface.IsPage(page, pageClasses, owner + 1, insideX, insideY));
            Assert.False(OfficeDocumentSurface.IsPage(IntPtr.Zero, pageClasses, owner, insideX, insideY));

            // A modal dialog disables the document's frame.
            EnableWindow(frame, false);
            Assert.False(OfficeDocumentSurface.IsPage(page, pageClasses, owner, insideX, insideY));
            EnableWindow(frame, true);
            Assert.True(OfficeDocumentSurface.IsPage(page, pageClasses, owner, insideX, insideY));

            ShowWindow(page, SwHide);
            Assert.False(OfficeDocumentSurface.IsPage(page, pageClasses, owner, insideX, insideY));
        }
        finally
        {
            DestroyWindow(frame);
        }
    }

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
