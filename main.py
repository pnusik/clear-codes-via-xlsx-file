import csv
import os
import sys
import traceback
from collections.abc import Iterator

import openpyxl
import openpyxl.cell
from lxml import etree
from plyer import filechooser  # pyright: ignore[reportMissingTypeStubs]


def treecodes(filename: str) -> set[str]:
    """
        Распарсивает xml дерево и добовляет плохие коды в множество. Возваршяет его.
    """
    tree = etree.parse(filename)
    cis_elements = tree.xpath("//pack_content/cis/text()")
    return {str(x) for x in cis_elements}



def fullcodeschecer(filename: str, badcodes: set[str]) -> Iterator[str]:
    """
    Генератор принимающий плохие коды и путь до основдного файла.
    Возвращяет коды которые уже подходят
    """
    with open(filename, "r", encoding="UTF-8") as f:
        reader = csv.DictReader(f, delimiter="	")
        for x in reader:
            if x["КМ"] in badcodes or x["Статус"] == "20":
                continue

            yield x["КМ"]



def select_file(multiple: bool) -> list[str]:
    """
    Вызывает выбор файла на виндовс
    """

    if getattr(sys, "frozen", False):
        current_dir = os.path.dirname(sys.executable)
    else:
        current_dir = os.path.dirname(os.path.abspath(__file__))

    
    path = filechooser.open_file(
            title="Выберите файл",
            path=current_dir,
            filters=[("Csv файл", "*.csv")],
            multiple=multiple
        ) # type: ignore
    if not path:
        raise TypeError("select_file() Путь отсутствует...")
    
    return path


def exel_creator(good_codes: Iterator[str]):
    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet()

    if not ws:
        raise TypeError("Exel error")

    for x in good_codes:
        cell = openpyxl.cell.WriteOnlyCell(ws, x)
        cell.number_format = "@"

        ws.append([cell])

    wb.save("result.xlsx")
    print("result.xlsx был создан или перезаписан")




def main():
    print("Выберете файл/файлы(через ctrl) в котором находятся кода исклчения(которые надо исключить из общего списка)")
    badcodes: set[str] = set()
    badcodes_file_url = select_file(True)
    for uri in badcodes_file_url:
        codes = treecodes(uri)
        badcodes.update(codes)


    print("Выберете файл в котором находятся все коды(полный список)")
    good_file_url = select_file(False)
    generate_good_codes = fullcodeschecer(filename=good_file_url[0], badcodes=badcodes)

    exel_creator(generate_good_codes)

if __name__ == "__main__":
    try:
        main()
        print("Успешно!")
    except Exception:  # noqa: BLE001
        print("ОШИБКА!!!!")
        print("-" * 40)
        traceback.print_exc()
        print("-" * 40)

input("Программа завершена...")