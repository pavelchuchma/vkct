import os

from openpyxl import load_workbook


class ResultWriter:
    def __init__(self, config, category_sum_results):
        self.config = config
        self.category_sum_results = category_sum_results
        self.outputFileName = "vysledky%s.xlsx" % self.config.year
        self.wb = None
        self.template_sheet = None

    def write(self):
        if os.path.exists(self.outputFileName):
            os.remove(self.outputFileName)

        self.wb = load_workbook('outputTemplate.xlsx')
        self.template_sheet = self.wb.get_sheet_by_name('Template')

        for cat_results in self.category_sum_results:
            self.prepare_sheet(cat_results.category.name, self.config.year, cat_results.category.get_title(), len(cat_results.personal_results))

            ws = self.wb.get_sheet_by_name(cat_results.category.name)
            row = 7
            for pr in cat_results.personal_results:
                ws.cell(row=row, column=2).value = pr.person.name
                ws.cell(row=row, column=3).value = pr.person.team
                ws.cell(row=row, column=4).value = pr.person.birth_year

                r = pr.race_results[0]
                self.write_position_and_points(ws.cell(row=row, column=5), ws.cell(row=row, column=6), r)

                for i in range(1, len(pr.race_results)):
                    r = pr.race_results[i]
                    self.write_position_and_points(ws.cell(row=row, column=7 + (i - 1) * 4),
                                                   ws.cell(row=row, column=8 + (i - 1) * 4), r)
                    if r.sum_points is not None:
                        ws.cell(row=row, column=9 + (i - 1) * 4).value = r.sum_position
                        ws.cell(row=row, column=10 + (i - 1) * 4).value = r.sum_points

                row = row + 1

        self.wb.remove_sheet(self.template_sheet)
        self.wb.save(self.outputFileName)
        self.wb.close()

    @staticmethod
    def write_position_and_points(pos_cell, points_cell, race_result):
        if race_result.position is not None:
            pos_cell.value = race_result.position if not race_result.half_points else "%s*" % race_result.position
            points_cell.value = race_result.points

    def prepare_sheet(self, cat_name, year, cat_title, line_count):
        ws = self.wb.copy_worksheet(self.template_sheet)
        ws.title = cat_name
        ws.cell(row=2, column=2).value = "%s %d" % (ws.cell(row=2, column=2).value, year)
        ws.cell(row=3, column=2).value = cat_title
        for col in range(2, 27):
            src = ws.cell(row=7, column=col)
            for row in range(8, 7 + line_count):
                n = ws.cell(row=row, column=col)
                n.border = src.border.copy()
                n.fill = src.fill.copy()
                n.font = src.font.copy()
                n.alignment = src.alignment.copy()

        ws.freeze_panes = ws.cell(row=7, column=5)
        ws.print_area = 'B2:Z%d' % (line_count + 6)
