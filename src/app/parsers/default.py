from src.app.models.esipdata import EsipData, LockoutRow
import re


def docx_parser(doc, data: EsipData) -> EsipData:
    """default parser for docx objects using python-docx

    :param doc: docx.Document object
    :param data: EsipData object
    :return: an EsipData object with 'data' from 'doc'
    """

    # parsing logic

    first_header_table = doc.sections[0].first_page_header.tables[0]

    document_no = ';'.join([x.strip() for x in first_header_table.cell(1, 1).text.strip().splitlines()])
    doc_no_regex = re.compile(r'8LWI[.]?')
    if len(re.split(doc_no_regex, document_no)) > 1:
        data.doc_no = re.split(doc_no_regex, document_no)[1].strip()
    else:
        data.doc_no = document_no

    date_list = []
    for header_r in first_header_table.rows:
        for header_cell in header_r.cells:
            header_cell_text = ";".join([x.strip() for x in header_cell.text.splitlines()])
            if re.search(r"\d\d?/\d\d?/\d\d\d?\d?|\d\d? [a-zA-Z]+ \d\d\d\d|[a-zA-Z]+ \d\d?, ?\d\d\d?\d?|\d\d?-\d\d?-\d\d\d?\d?|\d\d?/?\d\d?/\d\d\d?\d?|\d\d? ?/?\d\d? ?/ ?\d\d\d?\d?", header_cell_text):
                date_list.append(header_cell_text)

    date_eff = ""
    date_set = list(set(date_list))
    if len(date_set) == 1:
        find = re.search(r"\d\d?/\d\d?/\d\d\d?\d?|\d\d? [a-zA-Z]+ \d\d\d\d|[a-zA-Z]+ \d\d?, ?\d\d\d?\d?|\d\d?-\d\d?-\d\d\d?\d?|\d\d?/?\d\d?/\d\d\d?\d?|\d\d? ?/?\d\d? ?/ ?\d\d\d?\d?", date_set[0])
        date_eff = find.group()
    else:
        for date in date_set:
            find = re.search(r"\d\d?/\d\d?/\d\d\d?\d?|\d\d? [a-zA-Z]+ \d\d\d\d|[a-zA-Z]+ \d\d?, ?\d\d\d?\d?|\d\d?-\d\d?-\d\d\d?\d?|\d\d?/?\d\d?/\d\d\d?\d?|\d\d? ?/?\d\d? ?/ ?\d\d\d?\d?", date)
            find_effect = re.search(r"ffect", date)
            find_rev = re.search(r"evis", date)
            if find_rev:
                date_eff = find.group()
            elif find_effect and date_eff == "":
                date_eff = find.group()
        if date_eff == "":
            try:
                date_else = re.search(r"\d\d?/\d\d?/\d\d\d?\d?|\d\d? [a-zA-Z]+ \d\d\d\d|[a-zA-Z]+ \d\d?, ?\d\d\d?\d?|\d\d?-\d\d?-\d\d\d?\d?|\d\d?/?\d\d?/\d\d\d?\d?|\d\d? ?/?\d\d? ?/ ?\d\d\d?\d?", date_set[0])
                if date_else:
                    date_eff = date_else.group()
            except IndexError:
                pass

    # the date_eff takes the revised date first, then the effective date, and then any other date if found
    data.date_eff = date_eff

    equipment_table = doc.tables[0]

    try:
        department = ';'.join([x.strip() for x in equipment_table.cell(0, 1).text.strip().splitlines()])
        data.dept = department

        equipment_desc = equipment_table.cell(1, 1).text.strip()
        data.equip_desc = equipment_desc

        equipment_numb = ';'.join([x.strip() for x in equipment_table.cell(2, 1).text.strip().splitlines()])
        data.equip_id = equipment_numb
    except IndexError:
        # in the case where this table only has one column
        dept_regex = re.compile(r':')
        dept_string = ';'.join([x.strip() for x in equipment_table.cell(0, 0).text.strip().splitlines()])
        dept_split = re.split(dept_regex, dept_string)
        data.dept = dept_split[1].strip()

        equipment_desc_regex = re.compile(r':')
        equipment_desc_string = ';'.join([x.strip() for x in equipment_table.cell(1, 0).text.strip().splitlines()])
        equipment_desc_split = re.split(equipment_desc_regex, equipment_desc_string)
        data.equip_desc = equipment_desc_split[1].strip()

        equipment_numb_regex = re.compile(r':')
        equipment_numb_string = ';'.join([x.strip() for x in equipment_table.cell(2, 0).text.strip().splitlines()])
        equipment_numb_split = re.split(equipment_numb_regex, equipment_numb_string)
        data.equip_id = equipment_numb_split[1].strip()

    lockout_table = doc.tables[1]

    lockout_rows = []


    for i, row in enumerate(lockout_table.rows):
        lockout_row = []
        lr = LockoutRow()
        if i == 0:
            continue
        if row.cells[0].text.strip().lower() == "isolation date:":
            # isolation table logic
            for iso_row in lockout_table.rows[i:]:
                for iso_cell in iso_row.cells:
                    if re.search(r"special", iso_cell.text.strip().lower()):
                        data.special_precautions = re.search(r"[s|S]pecial ?(?:[p|P]recautions)?:?([\S\s]*)", iso_cell.text.strip()).group().strip()
            break
        else:
            # regular logic
            cell0 = row.cells[0].text.strip()
            cell1 = row.cells[1].text.strip()
            cell2 = row.cells[2].text.strip()
            cell3 = row.cells[3].text.strip()
            cell4 = row.cells[4].text.strip()

            if cell1 == "":
                continue
            elif cell1 == cell2 == cell3:
                lr.context = True
                lr.context_text = cell1
                lockout_rows.append(lr)
                del lr
                continue
            else:
                lr.num = cell0
                lr.point = cell1
                lr.tag_no = cell2
                lr.energy_src = cell3
                lr.isolating_means = cell4
                lockout_rows.append(lr)
                del lr
                continue

    data.lockout = lockout_rows

    try:
        # in case there is a separate isolation table
        isolation_table = doc.tables[2]
        for iso_row in isolation_table.rows:
            for iso_cell in iso_row.cells:
                if re.search(r"special", iso_cell.text.strip().lower()):
                    data.special_precautions = re.search(r"[s|S]pecial ?(?:[p|P]recautions)?:?([\S\s]*)", iso_cell.text.strip()).group().strip()
    except IndexError:
        # print(f'Must contain a joined table: {data.file_name}')
        data.joined = True

    return data


def txt_parser(txt, data) -> EsipData:
    """default text file parser

    :param data: EsipData object
    :param txt: text file object
    :return: an EsipData object with 'data' from 'txt'
    """
    txt_file_string = txt.read()
    # parsing logic

    approved_by_regex = re.compile(r"Approved By:?")
    prepared_by_regex = re.compile(r"Prepared By:?")
    date_revised_regex = re.compile(r"Date Revised:?")

    search_block = re.split(approved_by_regex, txt_file_string, maxsplit=1)[1]
    approved_by_block = re.split(date_revised_regex, search_block)
    approved_by_value = ';'.join(x.strip() for x in approved_by_block[0].strip().splitlines())
    approved_by_value = f"{approved_by_value}"

    approved_by = re.split(r"Rev", approved_by_value)[0]
    approved_by2 = re.split(r"8LWI", approved_by)[0]
    approved_by2a = re.split(r"Rev", approved_by2)[0]
    approved_by3 = re.sub(r';+$', "", approved_by2a)
    approved_by4 = re.sub(r'^;+', "", approved_by3)

    app_by_name = re.search(r"^(.+);(.+)", approved_by4)
    if app_by_name:
        data.app_by = app_by_name.group(1)
        data.app_by_title = app_by_name.group(2)
    else:
        data.app_by = approved_by4

    search_block_for_prepared = re.split(prepared_by_regex, txt_file_string)[1]
    prepared_by_text = re.split(r"Approved By:|Revision No\.:", search_block_for_prepared)[0]
    prepared_by_value = ";".join([x.strip() for x in prepared_by_text.splitlines()])
    prepared_by_value2 = re.sub(r';+$', "", prepared_by_value)
    prepared_by_value3 = re.sub(r'^;+', "", prepared_by_value)

    prepared_by_name = re.search(r"^(.+);(.+)", prepared_by_value3)
    if prepared_by_name:
        data.prep_by = prepared_by_name.group(1)
        data.prep_by_title = prepared_by_name.group(2)
    else:
        data.prep_by = prepared_by_value3

    requirements = re.compile(r"(3\.0 +Requirements)(.+?)(Lockout +Point)", flags=re.DOTALL)
    extra_section = re.search(requirements, txt_file_string)
    if extra_section:
        if re.search(r"3\.1|3\.2", extra_section.group(2)):
            extra = extra_section.group(2).strip()
            extra2 = re.sub(r"3\.\d +", "", extra)
            extra3 = extra2.replace("\t", "").splitlines()
            data.extra = extra3
    return data
