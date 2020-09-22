from typing import Callable

from excel.generate import generate_1032
from data.loader import load_values_from_dr568
from excel.drawing import add_drawings
from tempfile import mkstemp
import os

# def generate_1032(
#         template_path: str,
#         output_path: str,
#         card_data: List[CardBatch],
#         batch_number: int,
#         template_sheetname: str = "Sheet1"
# ):


def run_complete_process(
        data_input_path: str,
        data_input_sheet: str,
        batch_number: int,
        template_path: str,
        template_sheet: str,
        output_path: str,
        drawing_input_path: str,
        drawing_rel_input_path: str,
        printer_settings_input_path: str,
        sheet_rel_template_path: str,
        logo_input_path: str,
        progress_callback: Callable

):
    """
    Runs the entire process, soup to nuts
    progress_callback will be given two arguments: stage number and total num. stages
    """
    total_stages = 3

    # Step 0: Load data from input file


    vals = load_values_from_dr568(path_to_spreadsheet=data_input_path,
                                  from_sheet=data_input_sheet,
                                  target_batch=batch_number)

    if progress_callback:
        progress_callback(1, total_stages)
    # Step 1: Generate the blank 1032, writing to a temp path
    _, scratch_file_path = mkstemp()


    generate_1032(template_path=template_path,
                  output_path=scratch_file_path,
                  card_data=vals,
                  batch_number=batch_number,
                  template_sheetname=template_sheet)

    if progress_callback:
        progress_callback(2, total_stages)

    # Step 2: Do the secondary cleanup step (populating the drawing to each sheet, etc.)
    add_drawings(orig_input_file=scratch_file_path,
                 output_fname=output_path,
                 drawing_input_path=drawing_input_path,
                 drawing_rel_input_path=drawing_rel_input_path,
                 printer_settings_input_path=printer_settings_input_path,
                 sheet_rel_template_path=sheet_rel_template_path,
                 logo_input_path=logo_input_path
                 )

    # step 3: clean up scratch file
    os.remove(scratch_file_path)

    if progress_callback:
        progress_callback(3, total_stages)