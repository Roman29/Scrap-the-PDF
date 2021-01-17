from pdfminer.layout import LAParams, LTTextBox
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator

import re
import openpyxl


def write_to_xlsx(data: dict):
    file = 'Data Entry - 5th World Psoriasis & Psoriatic Arthritis Conference 2018 - Case format (2).xlsx'
    ref_workbook = openpyxl.load_workbook(file)

    sheet = ref_workbook.active
    max = sheet.max_row

    for row in data:
        for name in row['authors']['Names']:
            to_write = [name, str(row['authors']['Locations']), " ", row['title_id'], row['title'], row['abstract']]
            sheet.append(to_write)

    ref_workbook.save(file)


def get_title_id(source_strings):
    title_id = None
    regular_expression = re.compile('P\d+')

    title_id = regular_expression.match(source_strings).group(0)

    return title_id


def get_title(source_strings):
    intermediate_string = source_strings.replace('\n', ' ')
    intermediate_string = intermediate_string.replace('  ', ' ')
    reg_expression = re.compile("P\d*\s([A-Z0-9()\*: ­,-]+\s)")

    return reg_expression.match(intermediate_string).group(1)


def get_author(source_strings, title):
    names = []
    locations = []
    intermediate_string = source_strings.replace('\n', ' ')
    intermediate_string = intermediate_string.replace('  ', ' ')

    title_last_word = title.split()[-1]

    start = source_strings.find(title_last_word) + len(title_last_word) + 1
    end = source_strings.find("Introduction:")

    result = source_strings[start:end]

    result = re.sub(r'\d', '', result)
    result = result.replace('  ', ' ')

    list = result.split(',')

    if len(result.splitlines()) > 2:
        locations = []
        for item in list:
            matches = ['of', 'Hospital', 'University']
            if any(x in item for x in matches):
                locations.append(item)

        names = [x for x in list if x not in locations]

        names.append(locations[0].split('\n')[0])

        for index, item in enumerate(names):
            names[index] = item.replace("\n", " ").replace('  ', ' ').strip()

        locations[0] = locations[0][locations[0].find('\n') + 1:]

    else:
        names_string = result.split('\n')[0]
        locations_string = result.split('\n')[1]

        names = names_string.split(',')
        locations = locations_string.split(',')
        print(result)


    locations_string = ','.join(locations)
    locations_string = locations_string.replace('\n', '')

    return {"Names": names, "Locations": locations_string}


def get_abstract(source_strings):
    start = source_strings.find("Introduction:")

    return source_strings[start:]


if __name__ == '__main__':
    fp = open('Abstract Book from the 5th World Psoriasis and Psoriatic Arthritis Conference 2018.pdf', 'rb')
    rsrcmgr = PDFResourceManager()
    laparams = LAParams()
    device = PDFPageAggregator(rsrcmgr, laparams=laparams)
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    pages = PDFPage.get_pages(fp)

    cards = []
    need_exit = False
    for n, page in enumerate(pages):
        if n >= 43 and not need_exit:
            print('Processing next page...', n)
            interpreter.process_page(page)
            layout = device.get_result()
            for lobj in layout:
                if isinstance(lobj, LTTextBox):
                    x, y, text = lobj.bbox[0], lobj.bbox[3], lobj.get_text()
                    regex = re.compile('P\d+').match(text)
                    if regex:
                        identificator = regex.group()
                        cards.append(text)
                        if identificator == 'P106':
                            need_exit = True

        else:
            pass

    results = []
    for item in cards:
        title_id = get_title_id(item)
        title = get_title(item)
        authors = get_author(item, title)
        abstract = get_abstract(item)

        results.append({"title_id": title_id, "authors": authors, "title": title, "abstract": abstract})

    print(results)
    write_to_xlsx(results)
