import time

import wx
import wx.lib.scrolledpanel as scrolled
from app.common.constants import small_heading_font, default_font, create_font


class TemplateMultiparseDialog(wx.Dialog):

    def __init__(self, parameters, group_name, controller):
        """Constructor"""

        # DIALOG_NO_PARENT - to prevent been on top of the app
        wx.Dialog.__init__(self, parent=None, title='Выбрать параметры для фильтра',
                           style=wx.CAPTION | wx.DIALOG_NO_PARENT)

        self.parameters = parameters
        self.current_panel_parameters = controller.filters_parameters
        self.current_panel_product_parameters = controller.filtered_product_parameters
        self.current_panel_main_parameters = controller.main_product_parameters
        self.group_name = group_name
        self.controller = controller
        self._show_template_dlg_layout()

    def _show_template_dlg_layout(self):

        main_sizer = wx.BoxSizer(wx.VERTICAL)

        up_sizer = wx.BoxSizer(wx.VERTICAL)
        if self.group_name == 'general':
            label = 'Основные'
        elif self.group_name == 'additional':
            label = 'Дополнительные'
        else:
            self.SetTitle('Выбрать параметры для отчета')
            label = 'Параметры товаров для отчета'
        group_name_label = wx.StaticText(self, label=label)
        group_name_label.SetFont(create_font(small_heading_font))
        up_sizer.Add(group_name_label, flag=wx.ALIGN_CENTER | wx.ALL, border=5)
        self.up_sizer_scroll = None
        self.main_parameters_scroll = None
        if self.group_name in ('general', 'additional'):
            self.up_sizer_scroll = ScrolledPanel(self, self.parameters, self.current_panel_parameters, self.group_name)
        else:

            self.main_parameters_scroll = ScrolledPanel(self, None, self.current_panel_main_parameters,
                                                 self.group_name)
            self.up_sizer_scroll = ScrolledPanel(self, self.parameters, self.current_panel_product_parameters,
                                                 self.group_name)
        if self.main_parameters_scroll:
            up_sizer.Add(self.main_parameters_scroll, flag=wx.ALL | wx.EXPAND, border=0)
        up_sizer.Add(self.up_sizer_scroll, flag=wx.ALL | wx.EXPAND, border=0)
        down_sizer = wx.BoxSizer(wx.VERTICAL)

        button_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.accept_button = wx.Button(self, label='Применить', name='accept_' + self.group_name)
        self.cancel_button = wx.Button(self, label='Отменить', name='cancel')
        self.accept_button.SetFont(create_font(small_heading_font))
        self.cancel_button.SetFont(create_font(small_heading_font))
        button_sizer.Add(self.accept_button, flag=wx.ALL | wx.ALIGN_CENTER, border=5)
        button_sizer.Add(self.cancel_button, flag=wx.ALL | wx.ALIGN_CENTER, border=5)

        down_sizer.Add(button_sizer, flag=wx.ALL | wx.ALIGN_CENTER, border=10)

        main_sizer.Add(up_sizer, flag=wx.EXPAND)
        main_sizer.Add(down_sizer, flag=wx.EXPAND)
        self.SetSizerAndFit(main_sizer)

        self.accept_button.Bind(wx.EVT_BUTTON, self._close_dialog)
        self.cancel_button.Bind(wx.EVT_BUTTON, self._close_dialog)

    def _close_dialog(self, event):
        button_name = event.GetEventObject().GetName()
        if button_name in ('accept_general', 'accept_additional'):
            for key, value in self.up_sizer_scroll.current_panel_parameters.items():
                if key == 'parameters_dict':
                    for parameter_id, control in value.items():
                        self.controller.filters_parameters[self.group_name][key][parameter_id] = \
                            control.GetCheckedStrings()
                if key == 'parameters_dict_from' or key == 'parameters_dict_to':
                    for parameter_id, control in value.items():
                        self.controller.filters_parameters[self.group_name][key][parameter_id] = \
                            control.GetStringSelection()
                if key == 'parameters_number_range_from' or key == 'parameters_number_range_to' or \
                        key == 'parameters_checkbox':
                    for parameter_id, control in value.items():
                        self.controller.filters_parameters[self.group_name][key][parameter_id] = \
                            control.GetValue()
            self.controller.filterSpecified = True
            self.Destroy()
        elif button_name == 'cancel':
            self.Destroy()
        else:
            for parameter_id, control in self.main_parameters_scroll.current_product_panel_parameters.items():
                self.controller.main_product_parameters[parameter_id] = control.GetValue()
            for parameter_group, parameter in self.up_sizer_scroll.current_product_panel_parameters.items():
                self.controller.filtered_product_parameters[parameter_group] = {}
                for parameter_id, control in parameter.items():
                    self.controller.filtered_product_parameters[parameter_group][parameter_id] = control.GetValue()
            self.Destroy()


class ScrolledPanel(scrolled.ScrolledPanel):
    def __init__(self, parent, parameters, panel_parameters, group_name):
        if parameters:
            scrolled.ScrolledPanel.__init__(self, parent, -1, size=(480, 480))
            self.parameters = parameters
            self.group_name = group_name
            self.panel_parameters = panel_parameters
            self.current_panel_parameters = {
                'parameters_dict': {},
                'parameters_dict_from': {},
                'parameters_dict_to': {},
                'parameters_number_range_from': {},
                'parameters_number_range_to': {},
                'parameters_checkbox': {},
            }
            self.current_product_panel_parameters = {}
            if self.group_name in ('general', 'additional'):
                self._show_search_scroll()
            elif self.parameters:
                self._show_product_scroll()
        else:
            scrolled.ScrolledPanel.__init__(self, parent, -1, size=(480, 240))
            self.panel_parameters = panel_parameters
            self.current_product_panel_parameters = {}
            self._show_main_parameters_scroll()

    def _show_search_scroll(self):
        up_sizer_line = wx.BoxSizer(wx.VERTICAL)
        bool_values_list = []
        for value in self.parameters[1]['facets'][self.group_name]['items']:
            unit = ''
            if 'unit' in value.keys():
                if value['unit'] != '':
                    unit = ', ' + value['unit']
            up_sizer_parameter_label = wx.StaticText(self, label=value['name'] + unit)
            parameter_id = value['parameter_id']
            stored_parameters = self.panel_parameters[self.group_name]
            if value['type'] == 'boolean':
                parameter_checkbox = wx.CheckBox(self, -1, name=parameter_id)
                self.current_panel_parameters['parameters_checkbox'][parameter_id] = parameter_checkbox
                if parameter_id in stored_parameters['parameters_checkbox'].keys():
                    parameter_checkbox.SetValue(stored_parameters['parameters_checkbox'][parameter_id])
                bool_values_list.append((parameter_checkbox, up_sizer_parameter_label))
                continue

            up_sizer_line.Add(up_sizer_parameter_label, flag=wx.ALIGN_CENTER | wx.UP, border=3)
            if value['type'] == 'dictionary':

                if parameter_id in self.parameters[1]['dictionaries'].keys():
                    parameter_dict = wx.CheckListBox(self, -1, name=parameter_id)
                    self.current_panel_parameters['parameters_dict'][parameter_id] = parameter_dict
                    parameter_dict.Append([i['name'] for i in self.parameters[1]['dictionaries'][parameter_id]])
                    if parameter_id in stored_parameters['parameters_dict'].keys():
                        parameter_dict.SetCheckedStrings(stored_parameters['parameters_dict'][parameter_id])
                    up_sizer_line.Add(parameter_dict, flag=wx.ALIGN_CENTER | wx.DOWN, border=3)

            elif value['type'] == 'dictionary_range':
                if parameter_id in self.parameters[1]['dictionaries'].keys():
                    parameter_dict_from = wx.ComboBox(self, -1, name=parameter_id + '_from')
                    parameter_dict_to = wx.ComboBox(self, -1, name=parameter_id + '_to')
                    self.current_panel_parameters['parameters_dict_from'][parameter_id] = parameter_dict_from
                    self.current_panel_parameters['parameters_dict_to'][parameter_id] = parameter_dict_to
                    insert_list = [i['name'] for i in self.parameters[1]['dictionaries'][parameter_id]]
                    parameter_dict_from.Append(insert_list)
                    parameter_dict_to.Append(insert_list)

                    if parameter_id in stored_parameters['parameters_dict_from'].keys():
                        parameter_dict_from.SetStringSelection(stored_parameters['parameters_dict_from'][parameter_id])
                    if parameter_id in stored_parameters['parameters_dict_to'].keys():
                        parameter_dict_to.SetStringSelection(stored_parameters['parameters_dict_to'][parameter_id])

                    up_sizer_line_range = wx.BoxSizer(wx.HORIZONTAL)
                    up_sizer_line_range.Add(wx.StaticText(self, label='От'), flag=wx.LEFT | wx.ALIGN_CENTER, border=5)
                    up_sizer_line_range.Add(parameter_dict_from, flag=wx.ALL | wx.ALIGN_CENTER, border=5)
                    up_sizer_line_range.Add(wx.StaticText(self, label='До'), flag=wx.LEFT | wx.ALIGN_CENTER, border=5)
                    up_sizer_line_range.Add(parameter_dict_to, flag=wx.ALL | wx.ALIGN_CENTER, border=5)

                    up_sizer_line.Add(up_sizer_line_range, flag=wx.ALIGN_CENTER)

            elif value['type'] == 'number_range':
                parameter_number_range_from = wx.TextCtrl(self, -1, '', name=parameter_id + '_from')
                parameter_number_range_to = wx.TextCtrl(self, -1, '', name=parameter_id + '_to')
                self.current_panel_parameters['parameters_number_range_from'][parameter_id] = parameter_number_range_from
                self.current_panel_parameters['parameters_number_range_to'][parameter_id] = parameter_number_range_to

                if parameter_id in stored_parameters['parameters_number_range_from'].keys():
                    parameter_number_range_from.SetValue(stored_parameters['parameters_number_range_from'][parameter_id])
                if parameter_id in stored_parameters['parameters_number_range_to'].keys():
                    parameter_number_range_to.SetValue(stored_parameters['parameters_number_range_to'][parameter_id])

                if parameter_id in self.parameters[1]['placeholders'].keys():
                    parameter_number_range_from.SetHint(str(self.parameters[1]['placeholders'][parameter_id]['from']))
                    parameter_number_range_to.SetHint(str(self.parameters[1]['placeholders'][parameter_id]['to']))

                up_sizer_line_range = wx.BoxSizer(wx.HORIZONTAL)
                up_sizer_line_range.Add(wx.StaticText(self, label='От'), flag=wx.LEFT | wx.ALIGN_CENTER, border=5)
                up_sizer_line_range.Add(parameter_number_range_from, flag=wx.ALL | wx.ALIGN_CENTER, border=5)
                up_sizer_line_range.Add(wx.StaticText(self, label='До'), flag=wx.LEFT | wx.ALIGN_CENTER, border=5)
                up_sizer_line_range.Add(parameter_number_range_to, flag=wx.ALL | wx.ALIGN_CENTER, border=5)
                up_sizer_line.Add(up_sizer_line_range, flag=wx.ALIGN_CENTER)

        if bool_values_list:
            bool_ver_sizer = wx.BoxSizer(wx.VERTICAL)
            for checkbox, text in bool_values_list:
                bool_hor_sizer = wx.BoxSizer(wx.HORIZONTAL)
                bool_hor_sizer.Add(checkbox, flag=wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
                bool_hor_sizer.AddSpacer(10)
                bool_hor_sizer.AddStretchSpacer()
                bool_hor_sizer.Add(text, flag=wx.ALIGN_CENTER)
                bool_hor_sizer.AddStretchSpacer()
                bool_ver_sizer.Add(bool_hor_sizer, flag=wx.EXPAND | wx.ALL, border=5)
            up_sizer_line.Add(bool_ver_sizer, flag=wx.ALIGN_CENTER_HORIZONTAL | wx.ALL, border=3)
        self.SetSizer(up_sizer_line)
        self.SetupScrolling()

    def _show_main_parameters_scroll(self):
        up_sizer = wx.BoxSizer(wx.VERTICAL)
        stored_parameters = self.panel_parameters
        up_sizer_parameter_label = wx.StaticText(self, label='Основные параметры')
        up_sizer_parameter_label.SetFont(create_font(default_font))
        up_sizer.Add(up_sizer_parameter_label, flag=wx.ALIGN_CENTER | wx.UP, border=5)
        for key, value in stored_parameters.items():
            up_sizer_line = wx.BoxSizer(wx.HORIZONTAL)
            parameter_checkbox = wx.CheckBox(self, -1, name=key)
            self.current_product_panel_parameters[key] = parameter_checkbox
            parameter_checkbox.SetValue(value)
            parameter_name = wx.StaticText(self, label=key)
            up_sizer_line.Add(parameter_checkbox, flag=wx.ALL | wx.ALIGN_CENTER, border=5)
            up_sizer_line.Add(parameter_name, flag=wx.ALL | wx.ALIGN_CENTER, border=5)
            up_sizer.Add(up_sizer_line, flag=wx.ALIGN_CENTER)
        self.SetSizer(up_sizer)
        self.SetupScrolling()

    def _show_product_scroll(self):
        up_sizer = wx.BoxSizer(wx.VERTICAL)
        stored_parameters = self.panel_parameters
        select_all_button = wx.Button(self, label='Выбрать все', name='select_all')
        up_sizer.Add(select_all_button, flag=wx.ALIGN_CENTER | wx.UP, border=5)
        select_all_button.Bind(wx.EVT_BUTTON, self._select_all_parameters)
        for key, value in self.parameters.items():
            self.current_product_panel_parameters[key] = {}
            up_sizer_parameter_label = wx.StaticText(self, label=key)
            up_sizer_parameter_label.SetFont(create_font(default_font))
            up_sizer.Add(up_sizer_parameter_label, flag=wx.ALIGN_CENTER | wx.UP, border=5)
            bool_ver_sizer = wx.BoxSizer(wx.VERTICAL)
            for parameter_id in value:
                bool_hor_sizer = wx.BoxSizer(wx.HORIZONTAL)
                parameter_checkbox = wx.CheckBox(self, -1, name=parameter_id)
                parameter_name = wx.StaticText(self, label=parameter_id)
                bool_hor_sizer.Add(parameter_checkbox, flag=wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
                bool_hor_sizer.AddSpacer(10)
                bool_hor_sizer.AddStretchSpacer()
                bool_hor_sizer.Add(parameter_name, flag=wx.ALIGN_CENTER)
                bool_hor_sizer.AddStretchSpacer()
                bool_ver_sizer.Add(bool_hor_sizer, flag=wx.EXPAND | wx.ALL, border=5)
                self.current_product_panel_parameters[key][parameter_id] = parameter_checkbox
                if key in stored_parameters.keys():
                    if parameter_id in stored_parameters[key].keys():
                        parameter_checkbox.SetValue(stored_parameters[key][parameter_id])

            up_sizer.Add(bool_ver_sizer, flag=wx.ALIGN_CENTER_HORIZONTAL | wx.ALL, border=3)
        self.SetSizer(up_sizer)
        self.SetupScrolling()

    def _select_all_parameters(self, event):
        button_current_status = event.GetEventObject().GetLabel()
        if button_current_status == 'Выбрать все':
            event.GetEventObject().SetLabel('Убрать все')
            for element in self.GetChildren():
                if isinstance(element, wx.CheckBox):
                    element.SetValue(True)
        else:
            event.GetEventObject().SetLabel('Выбрать все')
            for element in self.GetChildren():
                if isinstance(element, wx.CheckBox):
                    element.SetValue(False)
