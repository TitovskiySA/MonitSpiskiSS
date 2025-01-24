import wx
import os
from Logging import LogThread, ToLog

def except_method_brief(method):
    def wrapper(self, *args, **kwargs):
        #ToLog(f"{method.__name__} started")
        try:
            method(self, *args, **kwargs)
        except Exception as Err:
            ToLog(f"Error in {method.__name__}, *args = {args}, **kwargs = {kwargs}, Error code = {Err}")
            #raise Exception
        #else:
        #    ToLog(f"{method.__name__} finished successfully")
    return wrapper

#checkframe
class ChFrame(wx.Frame):
    def __init__(
        self, parent = None, label = " ", data = {"some_Key": "some_value"}, path_to_png = os.getcwd()):

        self.data = data
    
        #wx.Frame.__init__(self, None, -1, label)
        wx.Frame.__init__(
            self, parent, -1, label,
            style =
            wx.MINIMIZE_BOX|wx.CAPTION|wx.SYSTEM_MENU|wx.CLOSE_BOX|
            wx.CLIP_CHILDREN|wx.MAXIMIZE_BOX|wx.RESIZE_BORDER)

        frameIcon = wx.Icon(path_to_png)
        self.SetIcon(frameIcon)

        #self.sizer = sizer = wx.BoxSizer(wx.VERTICAL)
        self.sizer = sizer = wx.FlexGridSizer(rows = len(data.keys()), cols = 2, hgap = 6, vgap = 6)
        #for i in range (0, len(data.keys())):
            #sizer.AddGrowableRow(i, 0)
        #sizer.AddGrowableCol(0, 1)
        sizer.AddGrowableCol(1, 1)
        self.values = {}
        
        for key in data.keys():
            text = wx.StaticText(self, wx.ID_ANY, key)
            text.SetFont(wx.Font(12, wx.ROMAN, wx.NORMAL, wx.NORMAL))
            sizer.Add(text, -1, wx.ALL | wx.EXPAND , border = 5)
          
            temp = wx.TextCtrl(self, wx.ID_ANY, data[key], style = wx.TE_READONLY | wx.TE_CENTRE)
            temp.SetFont(wx.Font(12, wx.ROMAN, wx.NORMAL, wx.NORMAL))
            temp.SetValue(data[key])
            self.values.update({key: temp})
            sizer.Add(temp, -1, wx.ALL | wx.EXPAND , border = 5)
         

        colour = wx.SystemSettings.GetColour(wx.SYS_COLOUR_MENU)
        self.SetBackgroundColour(colour)
        
        self.SetSizer(sizer)
        
        self.SetMaxSize((1000, len(data.keys()) * 60))
        self.Layout()
        self.Fit()
        self.SetClientSize(500, len(data.keys()) * 60)
        self.Show(True)

    @except_method_brief
    def UpdateData(self, upd_data = {"some_Key": "some_value"}):
        for key in self.data.keys():
            for upd_key in upd_data.keys():
                if key == upd_key:
                    self.values[key].SetValue(upd_data[key])
