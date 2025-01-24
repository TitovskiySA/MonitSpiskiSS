import wx
from wx.adv import SplashScreen as SplashScreen
import os

# HelloFrame
class HelloFrame(SplashScreen):
    def __init__(self, parent = None, path_to_png = os.getcwd() + "\\images"):
        super(HelloFrame, self).__init__(
            bitmap = wx.Bitmap(name = path_to_png, type = wx.BITMAP_TYPE_PNG),
            splashStyle = wx.adv.SPLASH_CENTRE_ON_SCREEN | wx.adv.SPLASH_TIMEOUT,
            milliseconds = 1500,
            parent = None,
            id = -1,
            pos = wx.DefaultPosition,
            size = (1000,1000),
            #wx.DefaultSize,
            style = wx.STAY_ON_TOP | wx.BORDER_NONE)
        self.Show(True)
        self.Bind(wx.EVT_CLOSE, self.OnClose)

    def OnClose(self, event):
        event.Skip()
        self.Hide()
