import wx

class BasePasswordDialog(wx.Dialog):
    """Базовий клас для діалогів введення пароля, щоб уникнути дублювання."""
    def __init__(self, parent, title):
        super().__init__(parent, title=title)
        self.panel = wx.Panel(self)
        self.sizer = wx.BoxSizer(wx.VERTICAL)

    def _add_buttons(self):
        button_sizer = wx.StdDialogButtonSizer()
        
        ok_button = wx.Button(self.panel, wx.ID_OK, "ОК")
        ok_button.SetDefault()
        button_sizer.AddButton(ok_button)
        self.Bind(wx.EVT_BUTTON, self.on_ok_button, id=wx.ID_OK)

        cancel_button = wx.Button(self.panel, wx.ID_CANCEL, "Скасувати")
        button_sizer.AddButton(cancel_button)
        self.Bind(wx.EVT_BUTTON, self.on_cancel_button, id=wx.ID_CANCEL)

        button_sizer.Realize()
        self.sizer.Add(button_sizer, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 10)

    def _finalize(self, target_ctrl):
        self.panel.SetSizer(self.sizer)
        self.sizer.Fit(self)
        self.Centre()
        target_ctrl.SetFocus()

    def on_cancel_button(self, _event):
        self.EndModal(wx.ID_CANCEL)

    def on_ok_button(self, _event):
        # Перевизначається у спадкоємцях за потреби
        self.EndModal(wx.ID_OK)


class ConfirmPasswordDialog(BasePasswordDialog):
    def __init__(self, parent, title="Встановлення пароля"):
        super().__init__(parent, title)
        
        # Підписи та поля вводу
        self.sizer.Add(wx.StaticText(self.panel, label="Введіть пароль:"), 0, wx.ALL | wx.LEFT, 10)
        self.pass_ctrl1 = wx.TextCtrl(self.panel, style=wx.TE_PASSWORD | wx.TE_PROCESS_ENTER)
        self.sizer.Add(self.pass_ctrl1, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 10)
        
        self.sizer.Add(wx.StaticText(self.panel, label="Повторіть пароль:"), 0, wx.TOP | wx.LEFT, 10)
        self.pass_ctrl2 = wx.TextCtrl(self.panel, style=wx.TE_PASSWORD | wx.TE_PROCESS_ENTER)
        self.sizer.Add(self.pass_ctrl2, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 10)
        
        # Обробка натискання Enter у будь-якому полі
        self.pass_ctrl1.Bind(wx.EVT_TEXT_ENTER, self.on_ok_button)
        self.pass_ctrl2.Bind(wx.EVT_TEXT_ENTER, self.on_ok_button)

        self._add_buttons()
        self._finalize(self.pass_ctrl1)

    def on_ok_button(self, _event):
        password = self.pass_ctrl1.GetValue()
        confirm = self.pass_ctrl2.GetValue()
        
        if not password:
            wx.MessageBox("Пароль не може бути порожнім!", "Помилка", wx.OK | wx.ICON_ERROR, self)
            return
            
        if password != confirm:
            wx.MessageBox("Введені паролі не збігаються. Спробуйте ще раз.", "Помилка", wx.OK | wx.ICON_ERROR, self)
            self.pass_ctrl1.Clear()
            self.pass_ctrl2.Clear()
            self.pass_ctrl1.SetFocus()
            return
            
        self.EndModal(wx.ID_OK)

    def GetPassword(self):
        return self.pass_ctrl1.GetValue()


class PasswordDialog(BasePasswordDialog):
    def __init__(self, parent, message, title):
        super().__init__(parent, title)
        
        if message:
            self.sizer.Add(wx.StaticText(self.panel, label=message), 0, wx.ALL, 10)

        self.password_entry = wx.TextCtrl(self.panel, style=wx.TE_PASSWORD | wx.TE_PROCESS_ENTER)
        self.sizer.Add(self.password_entry, 0, wx.ALL | wx.EXPAND, 10)
        self.password_entry.Bind(wx.EVT_TEXT_ENTER, self.on_ok_button)

        self._add_buttons()
        self._finalize(self.password_entry)

    def GetValue(self):
        return self.password_entry.GetValue()
