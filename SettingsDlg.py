import wx


#settings dialog
class SettingsDlg(wx.Dialog):
    def __init__(
        self, settings, label = "Settings"):
        self.def_dict = {
            "num_days": ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10'], # numdays
            "time_renew": ['30', '60', '90', '120', '150', '180', '300', '600'], # time renew16
            "time_ftp": ['300', '600', '900', '1200', '1500'],
            "ftp_addr": "None", "ftp_login":"None", "ftp_password":"None"}

        choices = {}
        self.Data = []
        
        for key in self.def_dict.keys():
            if settings[key] in self.def_dict[key]:
                for i in range (0, len(self.def_dict[key])):
                    if self.def_dict[key][i] == settings[key]:
                        choices.update({key: i})
            else:
                choices.update({key: 0})
        
        wx.Dialog.__init__(self, None, -1, label)
 
        labels = [
            "Number of days", "Scanning interval, seconds",
            "Sending to FTP interval, seconds",
            "FTP address",
            "FTP login",
            "FTP password"]
        #posSText = []
        for i in range (0, len(labels)):
            text = wx.StaticText(self, wx.ID_ANY, labels[i], pos = (10, 10 + 60*i))
            text.SetFont(wx.Font(12, wx.ROMAN, wx.NORMAL, wx.NORMAL))

        i = 0
        for key in self.def_dict.keys():
            if isinstance (self.def_dict[key], list):
                temp = wx.Choice(self, wx.ID_ANY, pos = (10, 35 + 60*i), size = (380, 30),
                                 choices = [str(some) for some in self.def_dict[key]])
                temp.SetFont(wx.Font(12, wx.ROMAN, wx.NORMAL, wx.NORMAL))
                temp.SetSelection(choices[key])
                self.Data.append(temp)
            else:
                if key != "ftp_password":
                    temp = wx.TextCtrl(
                        self, wx.ID_ANY, settings[key], pos = (10, 35 + 60*i),
                        size = (330, 30))
                else:
                    temp = wx.TextCtrl(
                        self, wx.ID_ANY, settings[key], pos = (10, 35 + 60*i),
                        size = (330, 30), style = wx.TE_PASSWORD)
                
                temp.SetFont(wx.Font(12, wx.ROMAN, wx.NORMAL, wx.NORMAL))
                temp.SetValue(settings[key])
                self.Data.append(temp)

            i = i + 1
                
        OKButton = wx.Button(self, wx.ID_OK, "OK", pos = (140, 20 + 60*len(labels)), size = (120, 30))
        OKButton.SetDefault()
        OKButton.SetFont(wx.Font(12, wx.ROMAN, wx.NORMAL, wx.NORMAL))
        self.SetClientSize((400, 60 + len(labels)*60))
        self.Bind(wx.EVT_CLOSE, self.NoClose)
    
    def NoClose(self, evt):
        print("No Close")
        
