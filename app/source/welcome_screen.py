from app.common.constants import heading_font, create_font
import wx


class WelcomePanel(wx.Panel):
    def __init__(self, parent, size):
        wx.Panel.__init__(self, parent, size=size)
        self.SetFont(create_font(heading_font))
        main_sizer = wx.BoxSizer(orient=wx.VERTICAL)

        line_1 = wx.BoxSizer(orient=wx.HORIZONTAL)
        line_1.Add((0, 100), flag=wx.ALL | wx.ALIGN_CENTER_VERTICAL)

        line_2 = wx.BoxSizer(orient=wx.VERTICAL)
        main_text = (
            'Добро пожаловать! Для начала работы:\n'
            '\n1. Выберите вкладку \"Обработать каталог Onliner\"'
            '\n2. Выберите раздел, категорию и подкатегорию товаров'
            '\n3. Задайте необходимые условия фильтрации товаров'
            '\n4. Выберите параметры, которые необходимо выгрузить в отчет по каждому из товаров'
            '\n5. Сформируйте отчет!'
            '\n\nP.S. При возникновении сбоев при выгрузке Вы можете продолжить выгрузку в тот же файл, '
            'выбрав его в соотвествующем окне\n\n'
        )
        self.main_label = wx.StaticText(self, label=main_text)
        line_2.Add(self.main_label, flag=wx.ALL, border=10)
        line_2.Add((0, 100), flag=wx.ALL | wx.ALIGN_CENTER_HORIZONTAL)

        main_sizer.Add(line_1, flag=wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, border=0)
        main_sizer.Add(line_2, flag=wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, border=0)

        self.SetSizer(main_sizer)
