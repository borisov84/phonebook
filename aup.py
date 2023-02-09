import openpyxl as op
from lxml import etree

filename = 'ауп.xlsx'
wb = op.load_workbook(filename)
sheet_ranges = wb['Телефоны АУП']
max_row = sheet_ranges.max_row

departments = ['Руководство ',
               'Первичная профсоюзная организация                                                                                             "Газпром трансгаз Югорск профсоюз"',
               'Служба по связям с общественностью и СМИ', 'Юридический отдел', 'Отдел кадров и трудовых отношений',
               'Отдел внутреннего аудита', 'Бухгалтерия', 'Отдел  налогов', 'Финансовый отдел',
               'Планово-экономический отдел', 'Отдел страхования', 'Отдел организации труда и заработной платы',
               'Отдел управления имуществом', 'Отдел подготовки и проведения закупок',
               ' Служба организации реконструкции и строительства основных фондов', ' Специальный отдел',
               'Служба корпоративной защиты']
department_rows = []
department_ranges = []


def get_deps_row():
    department = ""
    for num in sheet_ranges.iter_rows(min_row=1, max_col=1, max_row=max_row):
        for department in departments:
            if num[0].value == department:
                department_rows.append(num[0].row)
                # print(department)
    department_rows.append(max_row)
    for index, rng in enumerate(department_rows):
        if index + 1 < len(department_rows):
            department_ranges.append([rng, department_rows[index + 1] - 1])
    print(f'Найдено {len(department_ranges)} диапазонов для {len(departments)} филиалов')
    # print(department_rows)


if __name__ == '__main__':
    get_deps_row()
