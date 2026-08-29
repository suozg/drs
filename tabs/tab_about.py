import wx

class AboutTab(wx.Panel):
    def __init__(self, parent, main_frame):
        super().__init__(parent)
        self.main_frame = main_frame
        self.init_ui()

    def init_ui(self):
        sizer = wx.BoxSizer(wx.VERTICAL)

        self.content_text = wx.TextCtrl(self, style=wx.TE_MULTILINE | wx.TE_READONLY | wx.HSCROLL | wx.VSCROLL)
        sizer.Add(self.content_text, 1, wx.EXPAND | wx.ALL, 5)
        self.SetSizer(sizer)

        self.show_how_to()

    def show_how_to(self):
        help_search = (
            "1. Пошук одного слова\n"
            "\tслово\n"
            "Знайде всі записи, що містять вказане слово.\n\n"
            "2. Пошук кількох слів (AND)\n"
            "\tслово1 слово2\n"
            "Знайде записи, що містять обидва слова (слово1 та слово2).\n\n"
            "3. Пошук з OR\n"
            "\tслово1 OR слово2\n"
            "Знайде записи, які містять хоча б одне зі слів.\n\n"
            "4. Пошук фрази (використовуємо лапки)\n"
            "\t'точна фраза'\n"
            "Знайде точний збіг заданої фрази, зберігаючи порядок слів.\n\n"
            "5. Виключення слів (NOT, -)\n"
            "\tслово1 -слово2\n"
            "Знайде записи, які містять слово1, але не містять слово2.\n\n"
            "6. Пошук за близькістю слів (NEAR)\n"
            "\tслово1 NEAR/5 слово2\n"
            "Знайде записи, в яких слово1 та слово2 знаходяться не далі ніж за 5 слів одне від одного.\n\n"
            "7. Пошук за префіксом (частина слова)\n"
            "\tсло*\n"
            "Знайде всі слова, які починаються зі сло (наприклад, слово, слон)."
        )

        self.content_text.SetEditable(True)
        self.content_text.Clear()

        for line in help_search.split("\n"):
            if line.strip().isdigit() or line.endswith("пошуку") or "." in line[:3]:
                self.content_text.SetDefaultStyle(wx.TextAttr(wx.NullColour, wx.NullColour, wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)))
                self.content_text.AppendText(line + "\n")
            elif line.startswith("\t"):
                self.content_text.SetDefaultStyle(wx.TextAttr(wx.NullColour, wx.NullColour, wx.Font(8, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)))
                self.content_text.AppendText(line.strip() + "\n")
            else:
                self.content_text.SetDefaultStyle(wx.TextAttr(wx.NullColour, wx.NullColour, wx.Font(8, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL)))
                self.content_text.AppendText(line + "\n")
        self.content_text.SetEditable(False)
        self.content_text.SetDefaultStyle(wx.TextAttr())

        # Повертаємо курсор на початок і прокручуємо до першого рядка
        self.content_text.SetInsertionPoint(0)
        self.content_text.ShowPosition(0)

