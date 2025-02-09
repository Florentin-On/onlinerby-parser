from app.common.constants import heading_font, create_font
import wx


class MultiparsePanel(wx.Panel):
    def __init__(self, parent, size):
        wx.Panel.__init__(self, parent, size=size)
        self.SetFont(create_font(heading_font))
        font = wx.Font(8, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_ITALIC, wx.FONTWEIGHT_NORMAL)

        main_sizer = wx.BoxSizer(orient=wx.HORIZONTAL)

        # LEFT SIDE
        self.parent = parent
        left_sizer = wx.BoxSizer(orient=wx.VERTICAL)

        self.highest_categories_label = wx.StaticText(self, label='Раздел')
        self.highest_categories_combobox = wx.ComboBox(self, -1, name='highest_categories_combobox')
        left_sizer_line_1 = wx.BoxSizer(orient=wx.VERTICAL)
        left_sizer_line_1_1 = wx.BoxSizer(orient=wx.VERTICAL)
        left_sizer_line_1.Add(self.highest_categories_label, flag=wx.LEFT | wx.ALIGN_CENTER,
                              border=10)
        left_sizer_line_1_1.Add(self.highest_categories_combobox, flag=wx.ALL | wx.EXPAND,
                                border=5)

        left_sizer_line_2 = wx.BoxSizer(orient=wx.VERTICAL)
        left_sizer_line_2_1 = wx.BoxSizer(orient=wx.VERTICAL)
        self.categories_label = wx.StaticText(self, label='Категория')
        self.categories_combobox = wx.ComboBox(self, -1, name='categories_combobox')
        left_sizer_line_2.Add(self.categories_label, flag=wx.ALIGN_CENTER | wx.UP, border=10)
        left_sizer_line_2_1.Add(self.categories_combobox, flag=wx.ALL | wx.EXPAND,
                                border=5)

        left_sizer_line_3 = wx.BoxSizer(orient=wx.VERTICAL)
        left_sizer_line_3_1 = wx.BoxSizer(orient=wx.VERTICAL)
        self.sub_categories_label = wx.StaticText(self, label='Податегория')
        self.sub_categories_combobox = wx.ComboBox(self, -1, name='sub_categories_combobox')
        left_sizer_line_3.Add(self.sub_categories_label, flag=wx.ALIGN_CENTER | wx.UP, border=10)
        left_sizer_line_3_1.Add(self.sub_categories_combobox, flag=wx.ALL | wx.EXPAND,
                                border=5)

        left_sizer_line_4 = wx.BoxSizer(orient=wx.VERTICAL)
        self.select_search_general_params_button = wx.Button(self, label='Задать основные параметры фильтра', name='general')
        self.select_search_general_params_button.Disable()
        left_sizer_line_4.Add(self.select_search_general_params_button, flag=wx.ALL | wx.ALIGN_CENTER, border=10)

        left_sizer_line_5 = wx.BoxSizer(orient=wx.VERTICAL)
        self.select_search_add_params_button = wx.Button(self, label='Задать дополнительные параметры фильтра', name='additional')
        self.select_search_add_params_button.Disable()
        left_sizer_line_5.Add(self.select_search_add_params_button,
                              flag=wx.ALL | wx.ALIGN_CENTER_VERTICAL | wx.ALIGN_CENTER, border=10)

        left_sizer_line_6 = wx.BoxSizer(orient=wx.VERTICAL)
        self.select_needed_params_button = wx.Button(self, label='Выбрать параметры для формирования отчета',
                                                         name='select_needed_params_button')
        self.select_needed_params_button.Disable()
        left_sizer_line_6.Add(self.select_needed_params_button,
                              flag=wx.ALL | wx.ALIGN_CENTER_VERTICAL | wx.ALIGN_CENTER, border=10)

        left_sizer_line_7 = wx.BoxSizer(orient=wx.VERTICAL)
        self.generate_report_button = wx.Button(self, label='Сформировать отчет с заданными параметрами',
                                                         name='generate_report_button')
        self.generate_report_button.Disable()
        left_sizer_line_7.Add(self.generate_report_button,
                              flag=wx.ALL | wx.ALIGN_CENTER_VERTICAL | wx.ALIGN_CENTER, border=10)

        left_sizer_line_8 = wx.BoxSizer(orient=wx.VERTICAL)
        self.dev_button = wx.Button(self, label='DEV', name='dev_button')
        left_sizer_line_8.Add(self.dev_button,
                              flag=wx.ALL | wx.ALIGN_CENTER_VERTICAL | wx.ALIGN_CENTER, border=10)

        left_sizer.Add((0, 0), proportion=1, flag=wx.ALL | wx.EXPAND, border=5)
        left_sizer.Add(left_sizer_line_1, proportion=1, flag=wx.ALL | wx.EXPAND, border=0)
        left_sizer.Add(left_sizer_line_1_1, proportion=1, flag=wx.ALL | wx.EXPAND, border=0)
        left_sizer.Add(left_sizer_line_2, proportion=1, flag=wx.ALL | wx.EXPAND, border=0)
        left_sizer.Add(left_sizer_line_2_1, proportion=1, flag=wx.ALL | wx.EXPAND, border=0)
        left_sizer.Add(left_sizer_line_3, proportion=1, flag=wx.ALL | wx.EXPAND, border=0)
        left_sizer.Add(left_sizer_line_3_1, proportion=1, flag=wx.ALL | wx.EXPAND, border=0)
        left_sizer.Add(left_sizer_line_4, proportion=1, flag=wx.ALL | wx.EXPAND, border=0)
        left_sizer.Add(left_sizer_line_5, proportion=1, flag=wx.ALL | wx.EXPAND, border=0)
        left_sizer.Add(left_sizer_line_6, proportion=1, flag=wx.ALL | wx.EXPAND, border=0)
        left_sizer.Add(left_sizer_line_7, proportion=1, flag=wx.ALL | wx.EXPAND, border=0)
        left_sizer.Add(left_sizer_line_8, proportion=1, flag=wx.ALL | wx.EXPAND, border=0)
        left_sizer.Add((0, 0), proportion=1, flag=wx.ALL | wx.EXPAND, border=5)

        # MAIN
        main_sizer.Add(left_sizer, proportion=1, flag=wx.LEFT | wx.RIGHT | wx.EXPAND, border=100)
        self.SetSizer(main_sizer)
