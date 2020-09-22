from openpyxl import load_workbook
from typing import Iterable, Tuple, List, Set, IO
from data.loader import CardBatch
from excel.populate import populate_worksheet

PRINT_AREA = "A1:U21"

def generate_1032(
        template_path: str,
        output_path: str,
        card_data: List[CardBatch],
        batch_number: int,
        template_sheetname: str = "Sheet1"
):
    template_wb = load_workbook(filename=template_path)
    source = template_wb[template_sheetname]
    for set_of_five in card_data:
        # make a new worksheet, cloned from the first one
        targ = template_wb.copy_worksheet(source)

        # fill it in
        populate_worksheet(targ, set_of_five, batch_number)

        # set the print area (this doesn't seem to make it through copy_worksheet):
        targ.print_area = PRINT_AREA

    # now get rid of Sheet1, we don't need it:
    s1 = template_wb[template_sheetname]
    template_wb.remove(s1)

    template_wb.save(output_path)
