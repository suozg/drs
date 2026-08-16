import wx

class PasswordDialog(wx.Dialog):
    def __init__(self, parent, message, title):
        super(PasswordDialog, self).__init__(parent, title=title)
        panel = wx.Panel(self)
        sizer = wx.BoxSizer(wx.VERTICAL)

        self.password_entry = wx.TextCtrl(panel, style=wx.TE_PASSWORD | wx.TE_PROCESS_ENTER)
        sizer.Add(self.password_entry, 0, wx.ALL | wx.EXPAND, 10)
        self.password_entry.Bind(wx.EVT_TEXT_ENTER, self.on_ok_button)

        button_sizer = wx.StdDialogButtonSizer()

        ok_button = wx.Button(panel, wx.ID_OK, "ОК")
        ok_button.SetDefault()
        button_sizer.AddButton(ok_button)
        self.Bind(wx.EVT_BUTTON, self.on_ok_button, id=wx.ID_OK)

        cancel_button = wx.Button(panel, wx.ID_CANCEL, "Скасувати")
        button_sizer.AddButton(cancel_button)
        self.Bind(wx.EVT_BUTTON, self.on_cancel_button, id=wx.ID_CANCEL)

        button_sizer.Realize()
        sizer.Add(button_sizer, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 10)

        panel.SetSizer(sizer)
        sizer.Fit(self) 
        self.Centre()
        self.password_entry.SetFocus()

    def on_ok_button(self, _event):
        self.EndModal(wx.ID_OK)

    def on_cancel_button(self, _event):
        self.EndModal(wx.ID_CANCEL)

    def GetValue(self):
        return self.password_entry.GetValue()
