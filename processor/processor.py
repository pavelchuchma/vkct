# coding: utf-8
#  from openpyxl import Workbook
from openpyxl import load_workbook
from collections import namedtuple
from datetime import date

from ResultWriter import ResultWriter


class Person:
    def __init__(self, name, team, birth_year):
        self.name = name
        self.team = team
        self.birth_year = birth_year

    def get_key(self):
        return ("%s@%s" % (self.name, self.birth_year)).encode('utf-8')


class ResultLine:
    def __init__(self, person, position):
        self.person = person
        self.position = position


class RaceResult:
    def __init__(self):
        self.position = None
        self.points = None
        self.sum_points = None
        self.sum_position = None


class PersonalResults:
    def __init__(self, person, race_count):
        self.person = person
        self.race_results = []
        for i in range(race_count):
            self.race_results.append(RaceResult())


class CategorySummaryResults:
    def __init__(self, category):
        self.category = category
        self.personal_results = []


ParsePosition = namedtuple("ParsePosition", "file sheet row")
CategoryInput = namedtuple("CategoryInput", "file_name sheet_name first_row name_col team_col birth_year_col pos_col")
Category = namedtuple("Category", "name min_age max_age inputs results")
Config = namedtuple("Config", "year categories")

parsePosition = ParsePosition


def main():
    config = load_config('config2019.xlsx')

    read_results(config)
    category_sum_results = extract_summary_results(config)
    complete_summary_results(category_sum_results)

    writer = ResultWriter(config, category_sum_results)
    writer.write()
    pass


def read_results(config):
    for cat in config.categories:
        for i in cat.inputs:
            res = read(i.file_name, i.sheet_name, i.first_row, i.name_col, i.team_col,
                       i.birth_year_col, i.pos_col)
            cat.results.append(res)


def extract_summary_results(config):
    category_sum_results = []
    for cat in config.categories:
        cat_res = CategorySummaryResults(cat)
        res_map = {}
        category_sum_results.append(cat_res)
        race_idx = 0
        for race_results in cat.results:
            for r in race_results:
                key = r.person.get_key()
                pr = res_map.get(key)
                if pr is None:
                    res_map[key] = pr = PersonalResults(r.person, len(cat.results))
                elif pr.person.team is None and pr.person.team is not None:
                    pr.person = Person(pr.person.name, pr.person.team, pr.person.birth_year)

                pr.race_results[race_idx].position = r.position
                pr.race_results[race_idx].points = compute_points(len(race_results), r.position)

            race_idx = race_idx + 1
        cat_res.personal_results = res_map.values()
    return category_sum_results


point_table = [50, 45, 40, 37, 34, 31, 28, 26, 24, 22, 20, 19, 18, 17, 16, 15, 14, 13, 12, 11, 10, 9, 8, 7, 6, 5, 4, 3,
               2, 1]


def compute_points(people_count, position):
    if position == 'DNF':
        return 0

    bonus = min(people_count, 30)
    p = max(bonus - (position - 1), 0)
    if position - 1 < len(point_table):
        p = p + point_table[position - 1]
    return p


#
# def get_sort_key(pr):
#     return pr.race_results[2].sum_points


def complete_summary_results(category_sum_results):
    for cat_results in category_sum_results:
        # compute sum_points
        if len(cat_results.personal_results) == 0:
            continue

        for pr in cat_results.personal_results:
            sum = None
            for rr in pr.race_results:
                if rr.points is not None:
                    sum = rr.points if sum is None else sum + rr.points
                if sum is not None:
                    rr.sum_points = sum

        # complete sum_position
        race_count = len(cat_results.personal_results[0].race_results)
        for i in range(1, race_count):
            cat_results.personal_results.sort(key=lambda pr: pr.race_results[i].sum_points, reverse=True)
            sum_pos = 1
            last_sum_points = None
            last_sum_position = None
            for pr in cat_results.personal_results:
                if pr.race_results[i].sum_points == last_sum_points:
                    pr.race_results[i].sum_position = last_sum_position
                else:
                    last_sum_points = pr.race_results[i].sum_points
                    pr.race_results[i].sum_position = last_sum_position = sum_pos
                sum_pos = sum_pos + 1


def load_config(config_file):
    wb = load_workbook(config_file)
    ws = wb.get_sheet_by_name('Kategorie')

    current_year = ws['B1'].value
    categories = []
    row = 4
    while ws.cell(row=row, column=1).value is not None:
        cat_name = ws.cell(row=row, column=1).value
        category = Category(cat_name,
                            ws.cell(row=row, column=2).value,
                            ws.cell(row=row, column=3).value,
                            load_category_input_config(wb, cat_name),
                            [])
        categories.append(category)
        row = row + 1

    return Config(current_year, categories)


def load_category_input_config(wb, category_name):
    ws = wb.get_sheet_by_name(category_name)

    input_configs = []
    row = 2
    while ws.cell(row=row, column=1).value is not None:
        file_name = ws.cell(row=row, column=1).value
        sheet_name = ws.cell(row=row, column=2).value
        first_row = ws.cell(row=row, column=3).value
        name_col = get_column_index(ws.cell(row=row, column=4).value)
        team_col = get_column_index(ws.cell(row=row, column=5).value)
        birth_year_col = get_column_index(ws.cell(row=row, column=6).value)
        pos_col = get_column_index(ws.cell(row=row, column=7).value)

        input_configs.append(CategoryInput(
            file_name, sheet_name, first_row, name_col, team_col, birth_year_col, pos_col
        ))
        row = row + 1
    return input_configs


def get_column_index(col_name):
    return ord(col_name[0]) - ord('A') + 1


def read(file_name, sheet_name, first_row, name_col, team_col, birth_year_col, pos_col):
    wb = load_workbook(file_name)
    ws = wb.get_sheet_by_name(sheet_name)

    parsePosition.file = file_name
    parsePosition.sheet = sheet_name

    lines = []
    row = first_row
    while ws.cell(row=row, column=name_col).value is not None:
        parsePosition.row = row
        line = create_normalized_result_line(
            ws.cell(row=row, column=name_col).value,
            ws.cell(row=row, column=team_col).value,
            ws.cell(row=row, column=birth_year_col).value,
            ws.cell(row=row, column=pos_col).value)

        if line is not None:
            lines.append(line)
        row = row + 1
    return lines


def create_normalized_result_line(name, team, birtyear, pos):
    n_name = normalize_name(name)
    # if n_name != name:
    #     info("Normalized name: '%s' -> '%s'" % (name, n_name))

    if not isinstance(birtyear, long):
        warning("Birth year '%s' of %s is not a number" % (birtyear, name))
    elif birtyear < date.today().year - 80 or birtyear > date.today().year - 1:
        warning("Birth year %d is strange" % birtyear)

    n_pos = pos
    if isinstance(n_pos, unicode) and n_pos[-1] == '.' and n_pos[:-1].isdecimal():
        # replace u"1." -> L1
        n_pos = long(n_pos[:-1])
    if not isinstance(n_pos, long):
        info("Position '%s' is not a number! DNF '%s'!" % (n_pos, n_name))
        n_pos = 'DNF'

    return ResultLine(Person(n_name, team, birtyear), n_pos)


def normalize_name(src):
    parts = src.replace('Mgr.', '').replace('Ing.', '').replace('ml.', '').replace('st.', '').strip().split()
    if len(parts) == 2:
        n1 = parts[0].capitalize()
        n2 = parts[1].capitalize()
        if n1 in first_names and n2 not in first_names:
            return "%s %s" % (n1, n2)
        if n2 in first_names and n1 not in first_names:
            return "%s %s" % (n2, n1)
        warning("Unable to detect first and second name: '%s'" % src)
        return src
    else:
        error("Unexpected name format, exactly 2 parts expected: '%s'" % src)
        return src


def info(msg):
    message("INFO: ", msg)


def warning(msg):
    message("WARNING: ", msg)


def error(msg):
    message("ERROR: ", msg)


def message(prefix, msg):
    print("%s%s @%s[%s]:%d" % (prefix, msg, parsePosition.file, parsePosition.sheet, parsePosition.row))


def load_first_names():
    res = {""}
    for l in open('firstNames.txt').read().split():
        res.add(l.decode('utf-8'))
    return res


first_names = load_first_names()

main()
