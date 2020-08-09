# coding: utf-8
from collections import namedtuple

from openpyxl import load_workbook

from ResultWriter import ResultWriter


class Person:
    def __init__(self, name, team, birth_year):
        self.name = name
        self.team = team
        self.birth_year = birth_year

    def get_key(self):
        return ("%s@%s" % (self.name, self.birth_year)).encode('utf-8')


class ResultLine:
    def __init__(self, person, position, is_alternative):
        self.person = person
        self.position = position
        self.is_alternative = is_alternative


class RaceResult:
    def __init__(self):
        self.position = None
        self.points = None
        self.sum_points = None
        self.sum_position = None
        self.half_points = False


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


class Category:
    def __init__(self, name, min_age, max_age, count_positions, inputs, current_year):
        self.name = name
        self.min_age = min_age
        self.max_age = max_age
        self.count_positions = count_positions
        self.inputs = inputs
        self.results = []  # ResultLine[]
        self.min_year = current_year - max_age
        self.max_year = current_year - min_age

    def get_race_count(self):
        i = 0
        for ipt in self.inputs:
            if not ipt.is_alternative:
                i = i + 1
        return i

    def get_title(self):
        if self.min_age == 0:
            return u'%s (%s a mladší)' % (self.name, self.min_year)
        if self.max_age == 100:
            return u'%s (%s a starší)' % (self.name, self.max_year)
        return u'%s (%s a %s)' % (self.name, self.min_year, self.max_year)

ParsePosition = namedtuple("ParsePosition", "file, sheet, row")
CategoryInput = namedtuple("CategoryInput",
                           "file_name sheet_name, first_row, name_col, name2_col, team_col, birth_year_col, pos_col, is_alternative")
Config = namedtuple("Config", "year, categories")

parsePosition = ParsePosition


def main():
    config = load_config('config2020.xlsx')

    read_results(config)
    category_sum_results = extract_summary_results(config)
    complete_summary_results(category_sum_results)

    writer = ResultWriter(config, category_sum_results)
    writer.write()
    pass


def read_results(config):
    for cat in config.categories:
        for i in cat.inputs:
            res = read_result_sheet(i.file_name, i.sheet_name, i.first_row, i.name_col, i.name2_col, i.team_col,
                                    i.birth_year_col, i.pos_col, i.is_alternative, cat)
            cat.results.append(res)


def extract_summary_results(config):
    category_sum_results = []
    for cat in config.categories:
        cat_res = CategorySummaryResults(cat)
        res_map = {}
        category_sum_results.append(cat_res)
        race_idx = -1
        for i in range(0, len(cat.results)):
            race_results = cat.results[i]

            if not cat.inputs[i].is_alternative:
                race_idx = race_idx + 1

            for res_line in race_results:
                key = res_line.person.get_key()
                pr = res_map.get(key)
                if pr is None:
                    res_map[key] = pr = PersonalResults(res_line.person, cat.get_race_count())
                elif pr.person.team is None and pr.person.team is not None:
                    pr.person = Person(pr.person.name, pr.person.team, pr.person.birth_year)

                rr = pr.race_results[race_idx]
                rr.position = res_line.position

                rr.points = compute_points(len(race_results), res_line.position, res_line.is_alternative,
                                           cat.count_positions)
                rr.half_points = res_line.is_alternative

        cat_res.personal_results = res_map.values()
    return category_sum_results


point_table = [50, 45, 40, 37, 34, 31, 28, 26, 24, 22, 20, 19, 18, 17, 16, 15, 14, 13, 12, 11, 10, 9, 8, 7, 6, 5, 4, 3,
               2, 1]


def compute_points(people_count, position, half_points, count_positions):
    if position == 'DNF':
        return 0

    # nejmensi kategorie jen bod za ucast
    if not count_positions:
        return 1

    if half_points:
        p = max(0, 26 - position)
    else:
        bonus = min(people_count, 30)
        p = max(bonus - (position - 1), 0)
        if position - 1 < len(point_table):
            p = p + point_table[position - 1]

    return p


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

        race_count = cat_results.category.get_race_count()
        # sort by first column if there is only one race
        if race_count == 1:
            cat_results.personal_results.sort(key=lambda pr: pr.race_results[0].points, reverse=True)

        # complete sum_position
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
                            ws.cell(row=row, column=6).value == 1,
                            load_category_input_config(wb, cat_name),
                            current_year)
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
        name2_col = get_column_index(ws.cell(row=row, column=5).value)
        team_col = get_column_index(ws.cell(row=row, column=6).value)
        birth_year_col = get_column_index(ws.cell(row=row, column=7).value)
        pos_col = get_column_index(ws.cell(row=row, column=8).value)
        is_alternative = ws.cell(row=row, column=8).value == 1

        input_configs.append(CategoryInput(
            file_name, sheet_name, first_row, name_col, name2_col, team_col, birth_year_col, pos_col, is_alternative
        ))
        row = row + 1
    return input_configs


def get_column_index(col_name):
    return ord(col_name[0]) - ord('A') + 1


def read_result_sheet(file_name, sheet_name, first_row, name_col, name2_col, team_col, birth_year_col, pos_col, is_alternative,
                      category):
    wb = load_workbook(file_name)
    try:
        ws = wb.get_sheet_by_name(sheet_name)
    except Exception:
        error("Failed to get sheet '%s' from %s" % (sheet_name, file_name))
        raise

    parsePosition.file = file_name
    parsePosition.sheet = sheet_name

    lines = []
    row = first_row
    while True:
        parsePosition.row = row
        name_val = ws.cell(row=row, column=name_col).value
        if not name_val or name_val.isspace():
            break

        if name2_col:
            name_val = name_val + ' ' + ws.cell(row=row, column=name2_col).value

        birth_year_cell = ws.cell(row=row, column=birth_year_col)
        line = create_normalized_result_line(
            name_val,
            ws.cell(row=row, column=team_col).value,
            birth_year_cell.value, has_approved_value(birth_year_cell),
            ws.cell(row=row, column=pos_col).value,
            is_alternative, category)

        if line is not None:
            lines.append(line)
        row = row + 1
    return lines


def has_approved_value(cell):
    return cell.font.b and cell.font.i and cell.font.u == 'single'


def create_normalized_result_line(name, team, birth_year, approved_birth_year, pos, is_alternative, category):
    n_name = normalize_name(name)

    birth_year = to_long(birth_year)
    if not isinstance(birth_year, long):
        warning("Birth year '%s' of %s is not a number" % (birth_year, name))
    elif birth_year < category.min_year or birth_year > category.max_year:
        if not approved_birth_year:
            warning("Birth year %d looks is out of range category %s (%d-%d)"
                    % (birth_year, category.name, category.min_year, category.max_year))
        # else:
        #     info("Birth year %d looks is out of range category %s (%d-%d)"
        #          % (birth_year, category.name, category.min_year, category.max_year))

    n_pos = to_long(pos)
    if not isinstance(n_pos, long):
        if n_pos != 'DNF':
            info("Position '%s' is not a number! DNF '%s'!" % (n_pos, n_name))
        n_pos = 'DNF'

    return ResultLine(Person(n_name, team, birth_year), n_pos, is_alternative)


def to_long(n):
    if isinstance(n, long):
        return n
    if isinstance(n, unicode):
        if n.isdecimal():
            return long(n)
        if len(n) > 1 and n[-1] == '.' and n[:-1].isdecimal():
            return long(n[:-1])
    return n


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
    if len(parts) == 3:
        n1 = parts[0].capitalize()
        n2 = parts[1].capitalize()
        n3 = parts[2].capitalize()
        if n1 in first_names and n2 in first_names and n3 not in first_names:
            return "%s %s %s" % (n1, n2, n3)
        if n1 not in first_names and n2 in first_names and n3 in first_names:
            return "%s %s %s" % (n2, n3, n1)
    else:
        error("Unexpected name format, 2 parts expected: '%s'" % src)
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
