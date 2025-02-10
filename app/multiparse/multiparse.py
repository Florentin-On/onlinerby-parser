import io
import json
import os
import time
import traceback
from threading import Thread

import openpyxl
import requests
import wx
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import PatternFill

from app.common.constants import get_panel_parameters, heading_excel_style, \
    heading_simple_excel_style, text_excel_style, link_excel_style, bool_true_excel_style, bool_false_excel_style, \
    get_main_parameters, heading_font, create_font
from app.common.safe_requesters import safe_get_requester
from app.common_ui.dialogs import dialog, confirmation_dialog, confirmation_with_cancel_dialog
from app.multiparse.multiparse_dialogs import TemplateMultiparseDialog


class Multiparse(wx.Panel):
    def __init__(self, parent, size):
        wx.Panel.__init__(self, parent=parent, size=size)
        self.filterNotSaved = False
        self.productNotSaved = False
        self.categories = self.load_categories()
        self.categories_parameters = {}
        self.panel_parameters = get_panel_parameters()
        self.product_parameters = {}
        self.main_product_parameters = get_main_parameters()
        self.panel_product_parameters = {}

        self.SetFont(create_font(heading_font))
        main_sizer = wx.BoxSizer(orient=wx.HORIZONTAL)
        left_sizer = wx.BoxSizer(orient=wx.VERTICAL)
        self.highest_categories_label = wx.StaticText(self, label='Раздел')
        self.product_category_combobox = wx.ComboBox(self, -1, name='product_category_combobox')
        left_sizer_line_1 = wx.BoxSizer(orient=wx.VERTICAL)
        left_sizer_line_1_1 = wx.BoxSizer(orient=wx.VERTICAL)
        left_sizer_line_1.Add(self.highest_categories_label, flag=wx.LEFT | wx.ALIGN_CENTER,
                              border=10)
        left_sizer_line_1_1.Add(self.product_category_combobox, flag=wx.ALL | wx.EXPAND,
                                border=5)

        left_sizer_line_2 = wx.BoxSizer(orient=wx.VERTICAL)
        left_sizer_line_2_1 = wx.BoxSizer(orient=wx.VERTICAL)
        self.categories_label = wx.StaticText(self, label='Категория')
        self.product_group_combobox = wx.ComboBox(self, -1, name='product_group_combobox')
        left_sizer_line_2.Add(self.categories_label, flag=wx.ALIGN_CENTER | wx.UP, border=10)
        left_sizer_line_2_1.Add(self.product_group_combobox, flag=wx.ALL | wx.EXPAND,
                                border=5)

        left_sizer_line_3 = wx.BoxSizer(orient=wx.VERTICAL)
        left_sizer_line_3_1 = wx.BoxSizer(orient=wx.VERTICAL)
        self.sub_categories_label = wx.StaticText(self, label='Податегория')
        self.product_section_combobox = wx.ComboBox(self, -1, name='product_section_combobox')
        left_sizer_line_3.Add(self.sub_categories_label, flag=wx.ALIGN_CENTER | wx.UP, border=10)
        left_sizer_line_3_1.Add(self.product_section_combobox, flag=wx.ALL | wx.EXPAND,
                                border=5)

        left_sizer_line_4 = wx.BoxSizer(orient=wx.VERTICAL)
        self.select_search_general_params_button = wx.Button(self, label='Задать основные параметры фильтра',
                                                             name='general')
        self.select_search_general_params_button.Disable()
        left_sizer_line_4.Add(self.select_search_general_params_button, flag=wx.ALL | wx.ALIGN_CENTER, border=10)

        left_sizer_line_5 = wx.BoxSizer(orient=wx.VERTICAL)
        self.select_search_add_params_button = wx.Button(self, label='Задать дополнительные параметры фильтра',
                                                         name='additional')
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
        left_sizer.Add((0, 0), proportion=1, flag=wx.ALL | wx.EXPAND, border=5)

        main_sizer.Add(left_sizer, proportion=1, flag=wx.LEFT | wx.RIGHT | wx.EXPAND, border=100)
        self.SetSizer(main_sizer)

        self.update_categories(self.product_category_combobox)

        self.product_category_combobox.Bind(wx.EVT_TEXT, self.categories_changes)
        self.product_group_combobox.Bind(wx.EVT_TEXT, self.categories_changes)
        self.product_section_combobox.Bind(wx.EVT_TEXT, self.categories_changes)

        self.select_search_general_params_button.Bind(wx.EVT_BUTTON, self.open_search_params)
        self.select_search_add_params_button.Bind(wx.EVT_BUTTON, self.open_search_params)
        self.select_needed_params_button.Bind(wx.EVT_BUTTON, self.open_search_params)
        self.generate_report_button.Bind(wx.EVT_BUTTON, self.generate_report)

    @staticmethod
    def load_categories() -> dict:
        """
        Загрузка всех категорий из каталога Онлайнера и формирование трехуровневого словаря вида
        Категория -> Группа -> Раздел
        """
        category_dict: dict = {}
        categories_info = safe_get_requester('https://catalog.api.onliner.by/navigation/elements', [])
        for category in categories_info:
            category_title = category.get('title')
            if category_title and category_title not in category_dict and category.get('slug').lower() != 'prime':
                category_dict[category_title]: dict = {}
                category_groups_info = safe_get_requester(category.get('groups_url', ''), [])
                for group in category_groups_info:
                    group_title = group.get('title')
                    if group_title and group_title not in category_dict[category_title]:
                        category_dict[category_title][group_title]: dict = {}
                        group_sections_info = group.get('links', [])
                        for section in group_sections_info:
                            section_title = section.get('title')
                            if section_title and section_title not in category_dict[category_title][group_title]:
                                category_dict[category_title][group_title][section_title]: dict = section.get(
                                    'source_urls', {})
        return category_dict

    def update_categories(self, product_combobox: wx.ComboBox) -> None:
        """
        Обновление ComboBox в зависимости от того, кто вызвал данный метод
        :param product_combobox: wx.ComboBox, вызывающий метод
        """
        categories_keys = []
        combobox_name = product_combobox.GetName()
        if combobox_name == 'product_category_combobox':
            categories_keys += list(self.categories.keys())
        if combobox_name == 'product_group_combobox':
            categories_keys += list(self.categories[self.product_category_combobox.GetValue()].keys())
        if combobox_name == 'product_section_combobox':
            categories_keys += list(self.categories[self.product_category_combobox.GetValue()][
                                        self.product_group_combobox.GetValue()].keys())
        product_combobox.Append(categories_keys)

    def categories_changes(self, event: wx.Event) -> None:
        """
        Метод смены выбора категории товаров. При выборе Раздела - подгружает информацию по его параметрам
        :param event: wx.Event
        """
        combobox_value = event.GetEventObject().GetValue()
        if combobox_value == '':  # Этот же метод вызывается при вызове Clear() у wx.ComboBox
            return None
        key = event.GetEventObject().GetName()

        if key == 'product_category_combobox':
            self.product_group_combobox.Clear()
            self.product_section_combobox.Clear()

            self.update_categories(self.product_group_combobox)
        elif key == 'product_group_combobox':
            self.product_section_combobox.Clear()

            self.update_categories(self.product_section_combobox)
        elif key == 'product_section_combobox':
            category = self.product_category_combobox.GetValue()
            group = self.product_group_combobox.GetValue()
            section_url = self.categories[category][group][combobox_value]
            if category not in self.categories_parameters.keys():
                self.get_all_panel_parameters(combobox_value, section_url['catalog.schema.facets'])

            button_disabled = not (self.select_search_general_params_button.IsEnabled() and
                                   self.select_search_add_params_button.IsEnabled() and
                                   self.select_needed_params_button.IsEnabled() and
                                   self.generate_report_button.IsEnabled())
            if button_disabled:
                self.select_search_general_params_button.Enable()
                self.select_search_add_params_button.Enable()
                self.select_needed_params_button.Enable()
                self.generate_report_button.Enable()

        button_enabled = self.select_search_general_params_button.IsEnabled() or \
                         self.select_search_add_params_button.IsEnabled() or \
                         self.select_needed_params_button.IsEnabled() or \
                         self.generate_report_button.IsEnabled()
        if self.product_section_combobox.GetValue() == '' and button_enabled:
            self.select_search_general_params_button.Disable()
            self.select_search_add_params_button.Disable()
            self.select_needed_params_button.Disable()
            self.generate_report_button.Disable()
        if self.productNotSaved:
            self.panel_parameters = get_panel_parameters()
            self.panel_product_parameters = {}
        self.productNotSaved = False

    def get_all_panel_parameters(self, category: str, category_url: str):
        section_facets_data = safe_get_requester(category_url, [])
        self.categories_parameters[category] = (category_url, section_facets_data)

    def open_search_params(self, event):
        category = self.product_section_combobox.GetValue()
        group_name = event.GetEventObject().GetName()
        if self.product_section_combobox.GetValue() != '':
            self.show_template_dialog(category, group_name)

    def show_template_dialog(self, category, group_name):
        if group_name in ('general', 'additional'):
            dlg = TemplateMultiparseDialog(self.categories_parameters[category], group_name, self)
        else:
            link = self.get_link_for_report()
            if link not in self.product_parameters.keys():
                try:
                    wait = wx.BusyInfo('Подгружаем параметры товаров...')
                    products_with_filter = safe_get_requester(link.lower(), {})
                    if products_with_filter.get('products'):
                        self.get_all_product_parameters(link, products_with_filter['products'][0])
                        wait = None
                    else:
                        wait = None
                        dialog('Не найдено', 'Товары с заданным фильтром не найдены!')
                        return
                except Exception as err:
                    wait = None
                    msg = 'Ошибка:\n' + str(err) + '\n\n' + 'Подробности в логах программы.'
                    traceback.print_exc()
                    dialog(caption='Ошибка', message=msg, style=wx.ICON_ERROR)
                    return
            if self.filterNotSaved:
                self.panel_product_parameters = {}
                self.filterNotSaved = False
            dlg = TemplateMultiparseDialog(self.product_parameters[link], group_name, self)
        dlg.Center()
        dlg.ShowModal()
        dlg.Destroy()
        return dlg

    def generate_report(self, event):
        link = self.get_link_for_report()
        url_get_dict = safe_get_requester(link.lower(), {})
        pages_count = url_get_dict['page']['last']
        total_product_count = url_get_dict['total']
        if not total_product_count:
            dialog('Не найдено', 'Товары с указанным фильтром не найдены!')
            return
        result = confirmation_with_cancel_dialog('Основные параметры', 'Выгрузить в отчет только основные параметры?')
        if result == wx.YES:
            main_parameters = True
        elif result == wx.NO:
            main_parameters = False
        else:
            return
        with wx.FileDialog(self, "Выберите место и имя для сохранения отчета",
                           wildcard="Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                           style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as fileDialog:
            if fileDialog.ShowModal() == wx.ID_CANCEL:
                return

            new_file = True
            pathname = fileDialog.GetPath()
            progress_window = None
            try:
                if os.path.isfile(pathname):
                    wb = openpyxl.load_workbook(pathname)
                    if 'DEV_ONLINER_PARSER' in wb.sheetnames:
                        stored_link = wb['DEV_ONLINER_PARSER']['A2'].value
                        if stored_link == link:
                            result = confirmation_dialog("Внимание", "Файл, который вы выбрали, уже содержит товары с "
                                                                     "заданным фильтром. Продолжить выгрузку в данный файл?",
                                                         style=wx.YES_NO | wx.ICON_EXCLAMATION)
                            if result == wx.YES:
                                new_file = False
                            else:
                                return
                        else:
                            result = confirmation_dialog("Внимание",
                                                         "Файл, который вы выбрали, уже содержит товары другого "
                                                         "фильтра. Вы утеряете все данные из данного файла. "
                                                         "Продолжить?",
                                                         style=wx.YES_NO | wx.ICON_EXCLAMATION)
                            if result == wx.NO:
                                return
                    else:
                        result = confirmation_dialog("Внимание", "Вы утеряете все данные из выбранного файла. "
                                                                 "Продолжить?",
                                                     style=wx.YES_NO | wx.ICON_EXCLAMATION)
                        if result == wx.NO:
                            return
                selected_parameters = {}
                if not main_parameters:
                    selected_parameters = self.get_selected_parameters(link)
                    if not selected_parameters:
                        return
                if new_file:
                    wait = wx.BusyInfo("Создается файл отчета...", None)
                    self.create_empty_excel_table(pathname, selected_parameters, link, main_parameters)
                wait = None
                wb = openpyxl.load_workbook(pathname)
                goods_amount = wb['DEV_ONLINER_PARSER']['A1'].value
                progress_window = wx.GenericProgressDialog('Товары выгружаются', 'Прогресс\nВыгружено {} из {}'
                                                           .format('0', str(total_product_count - goods_amount)),
                                                           maximum=total_product_count - goods_amount,
                                                           parent=self.parent,
                                                           style=wx.PD_APP_MODAL | wx.PD_ELAPSED_TIME |
                                                                 wx.PD_REMAINING_TIME | wx.PD_ESTIMATED_TIME |
                                                                 wx.PD_AUTO_HIDE | wx.PD_SMOOTH | wx.PD_CAN_ABORT)

                Thread(target=self.process_report, args=(pathname, link, url_get_dict, pages_count,
                                                         selected_parameters, wb,
                                                         progress_window),
                       kwargs={'main_parameters': main_parameters}).start()
            except Exception as err:
                msg = 'Ошибка создания отчета:\n' + str(err) + '\n\n' + 'Подробности в логах программы.'
                traceback.print_exc()
                dialog(caption='Ошибка', message=msg, style=wx.ICON_ERROR)
            wait = None

    def get_link_for_report(self):
        link = self.categories[self.product_category_combobox.GetValue()][self.product_group_combobox.GetValue()][
            self.product_section_combobox.GetValue()]['catalog.schema.products']
        if '?' in link:
            link += '&'
        else:
            link += '?'
        to_add = ''
        for _ in self.panel_parameters.values():
            for parameter_id, value in _.items():
                if value:
                    if parameter_id == 'parameters_dict':
                        for parameter_dict_id, parameters_dict_value in value.items():
                            if parameters_dict_value:
                                all_ids = []
                                for i in self.categories_parameters[self.product_section_combobox.GetValue()][1][
                                    'dictionaries'][parameter_dict_id]:
                                    if i['name'] in parameters_dict_value:
                                        all_ids.append(str(i['id']))
                                to_add += ''.join(
                                    [parameter_dict_id + '[' + str(id) + ']=' + parameters_dict_value_i + '&' for
                                     id, parameters_dict_value_i in enumerate(all_ids)])
                    if parameter_id == 'parameters_dict_from':
                        for parameter_dict_from_id, parameters_dict_from_value in value.items():
                            if parameters_dict_from_value:
                                all_ids = []
                                for i in self.categories_parameters[self.product_section_combobox.GetValue()][1][
                                    'dictionaries'][parameter_dict_from_id]:
                                    if i['name'] in parameters_dict_from_value:
                                        all_ids.append(str(i['id']))
                                to_add += ''.join(
                                    [parameter_dict_from_id + '[from]=' + parameters_dict_value_i + '&' for
                                     parameters_dict_value_i in all_ids])
                    if parameter_id == 'parameters_dict_to':
                        for parameter_dict_to_id, parameters_dict_to_value in value.items():
                            if parameters_dict_to_value:
                                all_ids = []
                                for i in self.categories_parameters[self.product_section_combobox.GetValue()][1][
                                    'dictionaries'][parameter_dict_to_id]:
                                    if i['name'] in parameters_dict_to_value:
                                        all_ids.append(str(i['id']))
                                to_add += ''.join([parameter_dict_to_id + '[to]=' + parameters_dict_value_i + '&' for
                                                   parameters_dict_value_i in all_ids])
                    if parameter_id == 'parameters_number_range_from':
                        for parameter_number_from_id, parameters_number_from_value in value.items():
                            if parameters_number_from_value:
                                to_add += parameter_number_from_id + '[from]=' + parameters_number_from_value + '&'
                    if parameter_id == 'parameters_number_range_to':
                        for parameter_number_to_id, parameters_number_to_value in value.items():
                            if parameters_number_to_value:
                                to_add += parameter_number_to_id + '[to]=' + parameters_number_to_value + '&'
                    if parameter_id == 'parameters_checkbox':
                        for parameter_checkbox_id, parameters_checkbox_value in value.items():
                            if parameters_checkbox_value:
                                to_add += parameter_checkbox_id + '=' + str(int(parameters_checkbox_value)) + '&'
        link += to_add
        return link

    def get_selected_parameters(self, link):
        all_headings = {}
        if self.panel_product_parameters:
            for key, value in self.panel_product_parameters.items():
                for flag_name, flag in value.items():
                    if flag:
                        if key not in all_headings.keys():
                            all_headings[key] = []
                        all_headings[key].append(flag_name)
        if not all_headings:
            if link not in self.product_parameters.keys():
                try:
                    wait = wx.BusyInfo('Подгружаем параметры товаров...')
                    url_get_dict = safe_get_requester(link.lower(), {})
                    self.get_all_product_parameters(link, url_get_dict['products'][0])
                    wait = None
                except Exception as err:
                    wait = None
                    msg = 'Ошибка:\n' + str(err) + '\n\n' + 'Подробности в логах программы.'
                    traceback.print_exc()
                    dialog(caption='Ошибка', message=msg, style=wx.ICON_ERROR)
                    return {}
            for key, value in self.product_parameters[link].items():
                all_headings[key] = value
        return all_headings

    def get_all_product_parameters(self, link, url_get_dict):
        b = safe_get_requester(url_get_dict['html_url'], raw_response=True).content
        soup = BeautifulSoup(b, 'lxml')
        specs_table = soup.find("table", class_="product-specs__table")
        specs_table_groups = specs_table.findAll('tbody')
        group_headings = []
        self.product_parameters[link] = {}
        for group in specs_table_groups:
            settings_block = group.find('tr', class_='product-specs__table-title')
            if settings_block is not None:
                group_headings.append(group.find('tr', class_='product-specs__table-title').text.strip())
                headings = []
                for finded_params in group.findAll('tr', class_=lambda x: x is None):
                    finded_td = finded_params.find('td', class_=lambda x: x is None)
                    if finded_td is not None:
                        headings.append(finded_td.text.strip())
                self.product_parameters[link][group_headings[-1]] = headings

    def get_selected_product_parameters(self, product_link, selected_parameters):
        # TODO: переделать выгрузку со страницы продукта, возможно дергать API
        b = safe_get_requester(product_link, raw_response=True).content
        soup = BeautifulSoup(b, 'lxml')
        specs_table = soup.find("table", class_="product-specs__table")
        specs_table_groups = specs_table.findAll('tbody')
        parameters = {}
        for group in specs_table_groups:
            parameter_group = group.find('tr', class_='product-specs__table-title').text.strip()
            parameters[parameter_group] = {}
            headings = ''
            for finded_params in group.findAll('tr', class_=''):
                finded_td = finded_params.find('td', class_='')
                finded_tip = finded_td.find('p', class_='product-tip__term')
                if finded_tip:
                    headings = finded_tip.text.strip()
                else:
                    headings = finded_td.text.strip()
                parameters[parameter_group][headings] = {}
                finded_text = finded_params.find('span', class_='value__text')
                if finded_text:
                    parameters[parameter_group][headings] = self.normilize_title(finded_text.text.strip())
                    continue
                finded_plus = finded_params.find('span', class_='i-tip')
                if finded_plus:
                    parameters[parameter_group][headings] = True
                    continue
                finded_minus = finded_params.find('span', class_='i-x')
                if finded_minus:
                    parameters[parameter_group][headings] = False
                    continue
                # здесь нужно FindAll
                finded_link = finded_params.findAll('span', class_='value__link')
                if finded_link:
                    parameters[parameter_group][headings] = [(self.normilize_title(i.text.strip()),
                                                              i.contents[0].attrs['href']) for i in finded_link]
                    continue
                parameters[parameter_group][headings] = 'НЕ НАЙДЕНО'
        return_parameters = {}
        for sel_key, sel_value in selected_parameters.items():
            return_parameters[sel_key] = {}
            for sel_key_value in sel_value:
                return_parameters[sel_key][sel_key_value] = '-'
        for key, value in parameters.items():
            if key in selected_parameters.keys():
                for param in value.keys():
                    if param in selected_parameters[key]:
                        return_parameters[key][param] = parameters[key][param]
        return return_parameters

    def create_empty_excel_table(self, pathname, all_headings, link, main_parameters):
        wb = Workbook()
        ws = wb.active
        ws.title = 'Report'
        id = 1
        wb.add_named_style(heading_excel_style)
        wb.add_named_style(text_excel_style)
        wb.add_named_style(link_excel_style)
        wb.add_named_style(bool_true_excel_style)
        wb.add_named_style(bool_false_excel_style)
        wb.add_named_style(heading_simple_excel_style)
        process_headings = [heading for heading in self.main_product_parameters.keys() if
                            self.main_product_parameters[heading]]
        for title in process_headings + ['        ']:
            ws.cell(column=id, row=1, value=title).style = 'Heading Style'
            # ws.column_dimensions[openpyxl.utils.get_column_letter(id)].width = len(title)

            id += 1

        if not main_parameters:
            for group in all_headings:
                ws.cell(column=id, row=1, value=group).style = 'Heading Style'
                # ws.column_dimensions[openpyxl.utils.get_column_letter(id)].width = len(group)
                id += 1
                for heading in all_headings[group]:
                    # ws.column_dimensions[openpyxl.utils.get_column_letter(id)].width = len(heading)
                    ws.cell(column=id, row=1, value=heading).style = 'Heading Text Style'

                    id += 1
        ws.freeze_panes = 'A2'
        ws.auto_filter.ref = ws.dimensions
        ws = wb.create_sheet('DEV_ONLINER_PARSER')
        ws.sheet_state = 'hidden'
        ws['A1'] = 0
        ws['A2'] = link
        wb.save(pathname)
        wb.close()

    def process_report(self, pathname, link, url_get_dict, pages_count, selected_parameters, wb: Workbook,
                       progress_window: wx.GenericProgressDialog, main_parameters=False):
        goods_amount = wb['DEV_ONLINER_PARSER']['A1'].value
        start_index = 1
        if goods_amount != 0:
            start_index = goods_amount // 30 + 1
        sleep = 0.2
        progress_window_bar = 0
        delta_iterate_value = 2
        for i in range(start_index, pages_count + 1):
            break_flag = False
            if i != 1:
                proceeded_link = link
                time.sleep(sleep)
                proceeded_link += 'page=' + str(i)
                url_get = requests.get(proceeded_link.lower()).text
                url_get_dict = json.loads(url_get)
            ws = wb.active
            start_page_index = goods_amount % 30
            for product in url_get_dict['products'][start_page_index:]:
                product_html = product['html_url']
                id = 1
                for main_header, value in self.main_product_parameters.items():
                    if value:
                        ws.cell(column=id, row=goods_amount + delta_iterate_value).style = 'Text Style'
                        if main_header == 'Картинка':
                            if product['images']['header']:
                                img_link = product['images']['header']
                                r = safe_get_requester(img_link, raw_response=True)
                                if r is not None:
                                    image_file = io.BytesIO(r.content)
                                    img = Image(image_file)
                                    img.anchor = ws.cell(column=1, row=goods_amount + delta_iterate_value).coordinate
                                    img_width = (img.width - 5) / 7.0
                                    if img_width > ws.column_dimensions['A'].width:
                                        ws.column_dimensions['A'].width = img_width
                                    ws.row_dimensions[goods_amount + delta_iterate_value].height = img.height / 1.33
                                    ws.add_image(img)
                            else:
                                ws.cell(column=id, row=goods_amount + delta_iterate_value, value='-')
                        if main_header == 'Бренд':
                            brand = product['full_name'].replace(product['name'], '')
                            ws.cell(column=id, row=goods_amount + delta_iterate_value, value=brand)
                        if main_header == 'Модель и ссылка на Onliner':
                            ws.cell(column=id, row=goods_amount + delta_iterate_value,
                                    value='=HYPERLINK("{}", "{}")'.format(product_html, product['name'])).style = \
                                'Link Style'
                        if main_header == 'Тип':
                            ws.cell(column=id, row=goods_amount + delta_iterate_value, value=product['name_prefix'])
                        if main_header == 'Цена минимальная':
                            ws.cell(column=id, row=goods_amount + delta_iterate_value,
                                    value=float(product['prices']['price_min']['amount']) if product['prices'] else '-')
                        if main_header == 'Цена максимальная':
                            ws.cell(column=id, row=goods_amount + delta_iterate_value,
                                    value=float(product['prices']['price_max']['amount']) if product['prices'] else '-')
                        if main_header == 'Количество предложений':
                            ws.cell(column=id, row=goods_amount + delta_iterate_value,
                                    value=product['prices']['offers']['count'] if product['prices'] else '-')
                        id += 1
                id += 1
                time.sleep(sleep)
                if not main_parameters:
                    selected_product_parameters = self.get_selected_product_parameters(product_html,
                                                                                       selected_parameters)
                    need_row_to_increase = 0
                    for group_id, value in enumerate(selected_product_parameters.values()):
                        ws.cell(column=id, row=goods_amount + delta_iterate_value, value=' ').fill = \
                            PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
                        id += 1
                        for parameter_name, parameter_value in value.items():
                            if isinstance(parameter_value, list):
                                need_row_to_increase = len(parameter_value) - 1
                                if need_row_to_increase > 1:
                                    for merge_id in range(1, id + sum(
                                            [len(i.values()) + 1 for i in selected_product_parameters.values()])):
                                        if merge_id != id:
                                            ws.merge_cells(start_column=merge_id,
                                                           start_row=goods_amount + delta_iterate_value,
                                                           end_column=merge_id,
                                                           end_row=goods_amount + delta_iterate_value +
                                                                   need_row_to_increase)
                                for list_value in parameter_value:
                                    ws.cell(column=id,
                                            row=goods_amount + delta_iterate_value + parameter_value.index(list_value),
                                            value='=HYPERLINK("{}", "{}")'.format(list_value[1], list_value[0])).style = \
                                        'Link Style'

                            elif isinstance(parameter_value, bool):
                                ws.cell(column=id, row=goods_amount + delta_iterate_value, value=str(parameter_value))
                                if parameter_value:
                                    ws.cell(column=id, row=goods_amount + delta_iterate_value).style = 'Bool True Style'
                                else:
                                    ws.cell(column=id,
                                            row=goods_amount + delta_iterate_value).style = 'Bool False Style'
                            else:
                                ws.cell(column=id, row=goods_amount + delta_iterate_value).style = 'Text Style'
                                ws.cell(column=id, row=goods_amount + delta_iterate_value, value=str(parameter_value))
                            id += 1
                    delta_iterate_value += need_row_to_increase
                goods_amount += 1
                wb['DEV_ONLINER_PARSER']['A1'].value = goods_amount
                progress_window_bar += 1
                if (progress_window.WasCancelled() or
                        not progress_window.Update(progress_window_bar,
                                                   newmsg='Прогресс\nВыгружено {} из {}. Выгружаю продукт: {}'.format(
                                                       progress_window_bar, str(progress_window.GetRange()),
                                                       product['full_name']))[0]):
                    break_flag = True
                    break
            if break_flag:
                break

        caption, message, icon = 'Завершено', 'Выгрузка успешно завершена!', wx.OK
        try:
            wb.save(pathname)
            if progress_window.WasCancelled():
                caption, message, icon = 'Прервано', 'Выгрузка прервана пользователем', wx.ICON_WARNING
        except PermissionError as err:
            msg = 'Ошибка создания отчета - указанный файл был открыт во время выгрузки:\n' + str(
                err) + '\n\nПодробности в логах программы.'
            traceback.print_exc()
            caption, message, icon = 'Ошибка', msg, wx.ICON_ERROR
        dialog(caption, message, icon)
        progress_window.Destroy()
        if os.path.isfile(pathname):
            os.startfile(pathname)
        return None
