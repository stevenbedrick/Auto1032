from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.worksheet.worksheet import Worksheet
from PIL import Image as PilImage
from itertools import zip_longest
from typing import Iterable, Tuple, List, Set
from io import BytesIO
from tqdm import tqdm


def populate_worksheet(ws_to_fill: Worksheet, vals: List[Tuple[int, int]], batch_num: int):
    """
    Actually populate a worksheet with a batch of values from the DR568

    Note that this happens in-place to ws_to_fill
    
    :param ws_to_fill:
    :param vals:
    :param batch_num:
    :return:
    """
    seq_num_col = "F"
    proxy_num_col = "G"
    start_row = 5

    for idx, (card_num, proxy_num) in enumerate(vals):
        row_offset = start_row + idx
        seq_cell = f"{seq_num_col}{row_offset}"
        proxy_cell = f"{proxy_num_col}{row_offset}"
        ws_to_fill[seq_cell].value = card_num
        ws_to_fill[proxy_cell].value = proxy_num

    # now set the title appropriately:
    first_card_num = vals[0][0]
    last_card_num = vals[-1][0]
    ws_to_fill.title = f"Env {first_card_num} to {last_card_num}"

    ws_to_fill["G2"].value = batch_num
