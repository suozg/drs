# tabs/tab_about.py
import wx

FLOOR = 0
WALL = 1
PLAYER = 2
BOX = 3
TARGET = 4
BOXTARGET = 5

LEVELS = [
    [
        "..111...",
        "..141...",
        "..1 1111",
        "1113 341",
        "14 32111",
        "111131..",
        "...141..",
        "...111..",
    ],
    [
        "11111111",
        "1 4   21",
        "1   43 1",
        "111 3111",
        "..1  1..",
        "..1  1..",
        "..1111..",
        "........",
    ],
    [
        "....1111",
        "..111  1",
        "111    1",
        "14  3121",
        "1443 3 1",
        "1114 3 1",
        "..111  1",
        "....1111",
    ],
    [
        "..1111..",
        "111..1..",
        "123.41..",
        "1111.11.",
        "..14.31.",
        "..1.1.1.",
        "..143.1.",
        "..11111.",
    ],
    [
        "11111...",
        "1...11111",
        "1.1..441",
        "1.1.3141",
        "1.23.41.",
        "1.4..4.1",
        "11111111",
    ],
    [
        "111111..",
        "1....111",
        "1.13...1",
        "1.4231.1",
        "1.41.4.1",
        "1..11111",
        "1111....",
    ],
    [
        "..111111",
        "111....1",
        "1.41.1.3",
        "1.43.3.1",
        "1.423..1",
        "11111111",
    ],
    [
        "..111111",
        "111....1",
        "1....3.1",
        "1.1.43111",
        "1.424..1",
        "11111111",
    ],
    [
        "11111...",
        "1...11111",
        "1.1..441",
        "1.1.3141",
        "1.23.41.",
        "1.4..4.1",
        "11111111",
    ],
    [
        "..111111",
        "111....1",
        "1....1.1",
        "1.13.3.1",
        "1.4241.1",
        "1.4..111",
        "111111..",
    ],
    [
        "11111111",
        "1......1",
        "1.1311.1",
        "1.1441.1",
        "1.1321.1",
        "1......1",
        "11111111",
    ],
    [
        "..111111",
        "111....1",
        "1..3.1.1",
        "1.1411.1",
        "1.14.2.1",
        "1......1",
        "11111111",
    ],
    [
        "..111111",
        "111....1",
        "1..441.1",
        "1.1321.1",
        "1.14.1.1",
        "1......1",
        "11111111",
    ],
    [
        "11111...",
        "1...11111",
        "1.1...341",
        "1.1.42141",
        "1...3.1.1",
        "11111...",
    ],
    [
        "..111111",
        "111....1",
        "1..31..1",
        "1.1.4.111",
        "1.123..1",
        "1......1",
        "11111111",
    ],
    [
        "11111111",
        "1......1",
        "1.4131.1",
        "1.4231.1",
        "1.4..1.1",
        "11111111",
    ],
    [
        "..111111",
        "111....1",
        "1..3...1",
        "1.1.4.311",
        "1.124.1.1",
        "11111111",
    ],
    [
        "..111111",
        "111....1",
        "1..1.4.1",
        "1.13.3.1",
        "1.1244.1",
        "11111111",
    ],
    [
        "11111111",
        "1....2.1",
        "1.1311.1",
        "1.1441.1",
        "1.1.31.1",
        "1......1",
        "11111111",
    ],
    [
        "..111111",
        "111....1",
        "1..3.1.1",
        "1.1.11.1",
        "1.1.24.1",
        "1......1",
        "11111111",
    ],
    [
        "11111111",
        "1......1",
        "1.1.43.1",
        "1.1231.1",
        "1.14...1",
        "11111111",
    ],
    [
        "..111111",
        "111....1",
        "1..3...1",
        "1.1.4.311",
        "1.124.1.1",
        "11111111",
    ],
    [
        "11111111",
        "1....2.1",
        "1.1.43.1",
        "1.1.43.1",
        "1......1",
        "11111111",
    ],
]


class SokobanPanel(wx.Panel):
    def __init__(self, parent, update_callback=None):
        super().__init__(parent, size=(160, 160))
        self.SetBackgroundColour(wx.Colour(30, 30, 30))
        self.update_callback = update_callback
        
        self.level_idx = 0
        self.max_x = 8
        self.max_y = 8
        self.map2 = []
        self.game_map = []
        self.px, self.py = 0, 0
        
        self.load_level()
        
        self.Bind(wx.EVT_PAINT, self.on_paint)
        self.Bind(wx.EVT_KEY_DOWN, self.on_key_down)
        self.Bind(wx.EVT_LEFT_DOWN, lambda evt: self.SetFocus())

    def load_level(self):
        if self.level_idx >= len(LEVELS):
            self.level_idx = 0
        elif self.level_idx < 0:
            self.level_idx = len(LEVELS) - 1
            
        raw_map = LEVELS[self.level_idx]
        self.max_y = len(raw_map)
        self.max_x = len(raw_map[0])

        self.map2 = [[0] * self.max_x for _ in range(self.max_y)]
        self.game_map = [[0] * self.max_x for _ in range(self.max_y)]
        self.px, self.py = 0, 0

        for y in range(self.max_y):
            for x in range(self.max_x):
                char = raw_map[y][x]
                val = FLOOR if char in ". " else int(char)
                self.map2[y][x] = val
                if val == PLAYER:
                    self.px, self.py = x, y

        self.reset_level()
        if self.update_callback:
            self.update_callback(self.level_idx + 1)

    def set_level(self, idx):
        self.level_idx = idx
        self.load_level()
        self.Refresh()
        self.SetFocus()

    def reset_level(self):
        for y in range(self.max_y):
            for x in range(self.max_x):
                d = self.map2[y][x]
                if d == TARGET:
                    d = FLOOR
                if d == PLAYER:
                    self.px, self.py = x, y
                    d = FLOOR
                self.game_map[y][x] = d
        self.game_map[self.py][self.px] = PLAYER

    def on_paint(self, event):
        dc = wx.PaintDC(self)
        width, height = self.GetClientSize()
        
        if self.max_x == 0 or self.max_y == 0:
            return

        cell_w = width / self.max_x
        cell_h = height / self.max_y

        for y in range(self.max_y):
            for x in range(self.max_x):
                c = self.game_map[y][x]
                if self.map2[y][x] == TARGET:
                    if c == FLOOR:
                        dc.SetBrush(wx.Brush(wx.Colour(60, 60, 60)))
                        dc.SetPen(wx.Pen(wx.Colour(60, 60, 60)))
                        dc.DrawRectangle(int(x * cell_w), int(y * cell_h), int(cell_w) + 1, int(cell_h) + 1)
                        dc.SetTextForeground(wx.Colour(255, 215, 0))
                        dc.DrawText("·", int(x * cell_w + cell_w / 3), int(y * cell_h + cell_h / 6))
                        continue
                    elif c == BOX:
                        c = BOXTARGET

                if c == FLOOR:
                    dc.SetBrush(wx.Brush(wx.Colour(40, 40, 40)))
                elif c == WALL:
                    dc.SetBrush(wx.Brush(wx.Colour(120, 120, 120)))
                elif c == PLAYER:
                    dc.SetBrush(wx.Brush(wx.Colour(0, 150, 255)))
                elif c == BOX:
                    dc.SetBrush(wx.Brush(wx.Colour(200, 120, 50)))
                elif c == BOXTARGET:
                    dc.SetBrush(wx.Brush(wx.Colour(50, 200, 50)))

                dc.SetPen(wx.Pen(wx.Colour(20, 20, 20)))
                dc.DrawRectangle(int(x * cell_w), int(y * cell_h), int(cell_w) + 1, int(cell_h) + 1)

    def on_key_down(self, event):
        key = event.GetKeyCode()

        if key == ord('Q'):
            wx.GetApp().ExitMainLoop()
            return
        elif key == ord(' '):
            self.reset_level()
            self.Refresh()
            return

        dx, dy = 0, 0
        if key == wx.WXK_UP:
            dy = -1
        elif key == wx.WXK_DOWN:
            dy = 1
        elif key == wx.WXK_LEFT:
            dx = -1
        elif key == wx.WXK_RIGHT:
            dx = 1
        else:
            event.Skip()
            return

        px1, py1 = self.px + dx, self.py + dy

        if not (0 <= px1 < self.max_x and 0 <= py1 < self.max_y):
            return

        if self.game_map[py1][px1] == BOX:
            px2 = px1 + dx
            py2 = py1 + dy
            if 0 <= px2 < self.max_x and 0 <= py2 < self.max_y:
                if self.game_map[py2][px2] == FLOOR:
                    self.game_map[py2][px2] = BOX
                    self.game_map[py1][px1] = FLOOR

        if self.game_map[py1][px1] == FLOOR:
            self.game_map[self.py][self.px] = FLOOR
            self.px, self.py = px1, py1
            self.game_map[self.py][self.px] = PLAYER

        self.Refresh()

        won = all(
            self.map2[y][x] != TARGET or self.game_map[y][x] == BOX
            for y in range(self.max_y)
            for x in range(self.max_x)
        )

        if won:
            wx.MessageBox("Рівень пройдено!", "Перемога!", wx.OK | wx.ICON_INFORMATION)
            self.level_idx += 1
            self.load_level()
            self.Refresh()


class AboutTab(wx.Panel):
    def __init__(self, parent, main_frame):
        super().__init__(parent)
        self.main_frame = main_frame
        self.init_ui()

    def init_ui(self):
        main_sizer = wx.BoxSizer(wx.VERTICAL)
        top_sizer = wx.BoxSizer(wx.HORIZONTAL)

        self.content_text = wx.TextCtrl(self, style=wx.TE_MULTILINE | wx.TE_READONLY | wx.HSCROLL | wx.VSCROLL)
        top_sizer.Add(self.content_text, 1, wx.EXPAND | wx.ALL, 5)

        game_container = wx.BoxSizer(wx.VERTICAL)
        
        game_label = wx.StaticText(self, label="Сокобан")
        game_label.SetFont(wx.Font(8, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        game_container.Add(game_label, 0, wx.ALIGN_CENTER | wx.BOTTOM, 2)

        ctrl_sizer = wx.BoxSizer(wx.HORIZONTAL)
        
        self.btn_prev = wx.Button(self, label="◄", size=(30, 22))
        self.btn_prev.Bind(wx.EVT_BUTTON, self.on_prev_level)
        
        self.lbl_level = wx.StaticText(self, label=f"1/{len(LEVELS)}")
        self.lbl_level.SetFont(wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        
        self.btn_next = wx.Button(self, label="►", size=(30, 22))
        self.btn_next.Bind(wx.EVT_BUTTON, self.on_next_level)

        ctrl_sizer.Add(self.btn_prev, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        ctrl_sizer.Add(self.lbl_level, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT | wx.LEFT, 5)
        ctrl_sizer.Add(self.btn_next, 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 5)

        self.sokoban = SokobanPanel(self, update_callback=self.update_level_display)
        game_container.Add(self.sokoban, 0, wx.ALIGN_CENTER | wx.ALL, 0)
        game_container.Add(ctrl_sizer, 0, wx.ALIGN_CENTER | wx.TOP, 5)

        top_sizer.Add(game_container, 0, wx.ALIGN_TOP | wx.ALL, 5)

        main_sizer.Add(top_sizer, 1, wx.EXPAND)
        self.SetSizer(main_sizer)

        self.show_how_to()

    def update_level_display(self, current_level):
        if hasattr(self, 'lbl_level'):
            self.lbl_level.SetLabel(f"{current_level}/{len(LEVELS)}")

    def on_prev_level(self, event):
        new_idx = (self.sokoban.level_idx - 1) % len(LEVELS)
        self.sokoban.set_level(new_idx)

    def on_next_level(self, event):
        new_idx = (self.sokoban.level_idx + 1) % len(LEVELS)
        self.sokoban.set_level(new_idx)

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

        self.content_text.SetInsertionPoint(0)
        self.content_text.ShowPosition(0)
