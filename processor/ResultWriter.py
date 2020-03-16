from openpyxl import load_workbook
import os


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
            self.prepare_sheet(cat_results.category.name, len(cat_results.personal_results))

            ws = self.wb.get_sheet_by_name(cat_results.category.name)
            row = 7
            for pr in cat_results.personal_results:
                ws.cell(row=row, column=2).value = pr.person.name
                ws.cell(row=row, column=3).value = pr.person.team
                ws.cell(row=row, column=4).value = pr.person.birth_year

                r = pr.race_results[0]
                if r.position is not None:
                    ws.cell(row=row, column=5).value = r.position
                    ws.cell(row=row, column=6).value = r.points

                for i in range(1, len(pr.race_results)):
                    r = pr.race_results[i]
                    if r.position is not None:
                        ws.cell(row=row, column=7 + (i - 1) * 4).value = r.position
                        ws.cell(row=row, column=8 + (i - 1) * 4).value = r.points
                    if r.sum_points is not None:
                        ws.cell(row=row, column=9 + (i - 1) * 4).value = r.sum_position
                        ws.cell(row=row, column=10 + (i - 1) * 4).value = r.sum_points

                row = row + 1

        self.wb.remove_sheet(self.template_sheet)
        self.wb.save(self.outputFileName)
        self.wb.close()

    def prepare_sheet(self, cat_name, line_count):
        ws = self.wb.copy_worksheet(self.template_sheet)
        ws.title = cat_name

        for col in range(2, 30):
            src = ws.cell(row=7, column=col)
            for row in range(8, 7 + line_count):
                n = ws.cell(row=row, column=col)
                n.border = src.border.copy()
                n.fill = src.fill.copy()
                n.font = src.font.copy()
                n.alignment = src.alignment.copy()

        ws.freeze_panes = ws.cell(row=7, column=5)

