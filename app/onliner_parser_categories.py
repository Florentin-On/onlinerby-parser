import requests
import json
import time
from datetime import date
import os
import errno
import sys
from tqdm.auto import tqdm
from bs4 import BeautifulSoup


def json_loads_byteified(json_text):
    return _byteify(
        json.loads(json_text, object_hook=_byteify),
        ignore_dicts=True
    )


def _byteify(data, ignore_dicts = False):
    # if this is a unicode string, return its string representation
    if isinstance(data, unicode):
        return data.encode('utf-8')
    # if this is a list of values, return list of byteified values
    if isinstance(data, list):
        return [ _byteify(item, ignore_dicts=True) for item in data ]
    # if this is a dictionary, return dictionary of byteified keys and values
    # but only if we haven't already byteified it
    if isinstance(data, dict) and not ignore_dicts:
        return {
            _byteify(key, ignore_dicts=True): _byteify(value, ignore_dicts=True)
            for key, value in data.iteritems()
        }
    # if it anything else, return it in its original form
    return data


def get_category_list():
    b = requests.get('https://catalog.onliner.by/').text
    soup = BeautifulSoup(b, 'lxml')
    content = soup.find_all("a", class_="catalog-navigation-classifier__item")
    cat_list = []
    for i in content:
        f = i
        flag = False
        while True:
            if f is None:
                break
            f = f.parent
            if f is not None and 'class' in f.attrs:
                if 'catalog-navigation-list__category' in f.attrs['class']:
                    if f.attrs['data-id'].isdigit() and int(f.attrs['data-id']) != 16:
                        flag = True
                        break
        k = i.attrs['href'].split('/')[-1]
        cat_list.append(k) if k not in cat_list else cat_list
    cat_list.sort()
    return cat_list


def process_categories(category_list, compare_with=None):
    loop = tqdm(total=len(category_list), position=0, leave=False)
    today = date.today()
    sleep = 0.2
    if compare_with is None:
        dir = os.getcwd() + '\\' + str(today) + '_report' + '\\'
        if not os.path.exists(os.path.dirname(dir)):
            try:
                os.makedirs(os.path.dirname(dir))
            except OSError as exc:  # Guard against race condition
                if exc.errno != errno.EEXIST:
                    raise
        print('\nReport in ', dir)
        for i in category_list:
            try:
                time.sleep(sleep)
                a = requests.get('https://catalog.onliner.by/sdapi/catalog.api/facets/' + i).text
                obj = json.loads(a)
                mfr = obj['dictionaries']['mfr']
                shops = obj['dictionaries']['shops']
                time.sleep(sleep)
                b = requests.get('https://catalog.onliner.by/sdapi/catalog.api/search/' + i).text
                obj2 = json.loads(b)
                count = obj2['total_ungrouped']
                i_norm = i.replace('?', '__').replace(':', '__')
                with open(dir + i_norm + '_mfr.txt', 'w+', encoding="utf-8") as f:
                    json.dump(mfr, f)
                    f.close()
                with open(dir + i_norm + '_shops.txt', 'w+', encoding="utf-8") as f:
                    json.dump(shops, f)
                    f.close()
                with open(dir + i_norm + '_count.txt', 'w+', encoding="utf-8") as f:
                    json.dump(count, f)
                    f.close()
            except Exception as err:
                print('Exception!', sys.exc_info()[0], err)
            loop.set_description('Loading...'.format(category_list.index(i)))
            loop.update(1)
        loop.close()
    else:
        dir = os.getcwd() + '\\' + str(today) + '_report_compared_with_' + compare_with + '\\'
        if not os.path.exists(os.path.dirname(dir)):
            try:
                os.makedirs(os.path.dirname(dir))
            except OSError as exc:  # Guard against race condition
                if exc.errno != errno.EEXIST:
                    raise
        dir_compare = os.getcwd() + '\\' + str(compare_with) + '\\'
        if not os.path.exists(os.path.dirname(dir_compare)):
            raise FileNotFoundError('Directory %s not found!' % dir_compare)
        print('\nReport in ', dir)
        for i in category_list:
            try:
                time.sleep(sleep)
                a = requests.get('https://catalog.onliner.by/sdapi/catalog.api/facets/' + i).text
                obj = json.loads(a)
                mfr = obj['dictionaries']['mfr']
                shops = obj['dictionaries']['shops']
                time.sleep(sleep)
                b = requests.get('https://catalog.onliner.by/sdapi/catalog.api/search/' + i).text
                obj2 = json.loads(b)
                count = obj2['total_ungrouped']
                i_norm = i.replace('?', '__').replace(':', '__')
                with open(dir + i_norm + '_mfr.txt', 'w+', encoding="utf-8") as f:
                    json.dump(mfr, f)
                    f.close()
                with open(dir + i_norm + '_shops.txt', 'w+', encoding="utf-8") as f:
                    json.dump(shops, f)
                    f.close()
                with open(dir + i_norm + '_count.txt', 'w+', encoding="utf-8") as f:
                    json.dump(count, f)
                    f.close()
                with open(dir_compare + i_norm + '_mfr.txt', 'r', encoding="utf-8") as f:
                    mfr_string_line = json.load(f)
                    if mfr_string_line != mfr:
                        print('\n', i, 'mfr diff (added/deleted)', '\n', [x for x in mfr_string_line if x not in mfr],
                              '\n', [x for x in mfr if x not in mfr_string_line])
                    f.close()
                with open(dir_compare + i_norm + '_shops.txt', 'r', encoding="utf-8") as f:
                    shops_string_line = json.load(f)
                    if shops_string_line != shops:
                        print('\n', i, 'shops diff (added/deleted)', '\n', [x for x in shops_string_line if x not in shops],
                              '\n', [x for x in shops if x not in shops_string_line])
                    f.close()
                with open(dir_compare + i_norm + '_count.txt', 'r', encoding="utf-8") as f:
                    count_string_line = json.load(f)
                    if count_string_line != count:
                        print('\n', i, 'count diff (was/now)', '\n', count_string_line, '\n', repr(count))
                    f.close()
            except Exception as err:
                print('Exception!', sys.exc_info()[0], err)
            loop.set_description('Loading...'.format(category_list.index(i)))
            loop.update(1)
        loop.close()


category_list = get_category_list()
try:
    process_categories(category_list)
    # process_categories(category_list, '2021-01-16_report')
except FileNotFoundError as err:
    print('!!!', err)






