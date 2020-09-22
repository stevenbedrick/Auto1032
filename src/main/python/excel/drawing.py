import zipfile
import os
import tempfile
import glob
import shutil
import re
from typing import Tuple, Mapping

from lxml import etree
from lxml.etree import Element, QName

orig_input_file = "test_sdb_output.xlsx"
output_fname = "test_sdb_output.with_drawing.xlsx"

# paths for the various bits and pieces we need to add:
DRAWING_INPUT_RESOURCE_PATH = "drawing1.xml"
DRAWING_REL_INPUT_RESOURCE_PATH = "drawing1.xml.rels"
PRINTER_SETTINGS_RESOURCE_INPUT_PATH = "printerSettings1.bin"
SHEET_REL_TEMPLATE_RESOURCE_PATH = "sheet1.xml.rels"
LOGO_RESOURCE_PATH="logo.jpeg"

# where to find various things in the unzipped spreadsheet:
drawing_output_path = "xl/drawings/drawing2.xml"
drawing_rel_output_path = "xl/drawings/_rels/drawing2.xml.rels"
manifest_file_path = "[Content_Types].xml"
printer_settings_dir_path = "xl/printerSettings/"
printer_settings_output_path = "xl/printerSettings/printerSettings1.bin"
workbook_path = "xl/workbook.xml"
workbook_rel_path = "xl/_rels/workbook.xml.rels"
logo_output_path="xl/media/image1.jpeg"

SHEET_NAME_REGEX = r"sheet(\d+).xml"


def load_worksheet_map(workbook_xml_path, workbook_rels_path) -> Tuple[Mapping, Mapping, Mapping]:
    """
    Parse the workbook, and map out the various worksheets, where to find them, and so on.
    :param workbook_xml_path:
    :param workbook_rels_path:
    :return:
    """
    workbook_doc = etree.parse(open(workbook_xml_path))
    workbook_rels_doc = etree.parse(open(workbook_rels_path))
    sheet_names_to_rel_id = {}

    for sheet in workbook_doc.xpath("/m:workbook/m:sheets/m:sheet",
                                    namespaces={'m': "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}):
        name = sheet.get('name')
        rId = sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        sheet_names_to_rel_id[name] = rId

    # invert the dictionary:
    rel_id_to_sheet_name = {r_id: s_name for s_name, r_id in sheet_names_to_rel_id.items()}

    rel_id_to_filename = {}

    for rel in workbook_rels_doc.xpath("/r:Relationships/r:Relationship", namespaces={
        'r': 'http://schemas.openxmlformats.org/package/2006/relationships'
    }):
        # we only want worksheets:
        if not rel.get('Type') == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet":
            continue

        target = rel.get('Target')
        rId = rel.get('Id')

        rel_id_to_filename[rId] = target

    fname_to_sheet_name = {os.path.basename(fname): rel_id_to_sheet_name[r_id] for r_id, fname in
                           rel_id_to_filename.items()}
    return sheet_names_to_rel_id, rel_id_to_filename, fname_to_sheet_name


def add_drawings(orig_input_file: str,
                 output_fname: str,
                 drawing_input_path: str,
                 drawing_rel_input_path: str,
                 printer_settings_input_path: str,
                 sheet_rel_template_path: str,
                 logo_input_path: str
                 ):
    z = zipfile.ZipFile(orig_input_file)

    # unzip into a temp working dir:
    with tempfile.TemporaryDirectory() as tmpdir:
        z.extractall(tmpdir)

        # step 0: load our map of sheet names to sheet files:
        _, _, xml_to_sheet_name = load_worksheet_map(os.path.join(tmpdir, workbook_path),
                                                     os.path.join(tmpdir, workbook_rel_path))

        # make sure that certain folders exist: TODO: don't hard-code
        os.makedirs(os.path.join(tmpdir, "xl/drawings"))
        os.makedirs(os.path.join(tmpdir, "xl/drawings/_rels"))
        os.makedirs(os.path.join(tmpdir, "xl/worksheets/_rels"))
        os.makedirs(os.path.join(tmpdir, "xl/media"))

        # step 1a: add the drawing (and rel) and logo to the appropriate folder:

        shutil.copyfile(drawing_input_path, os.path.join(tmpdir, drawing_output_path))
        shutil.copyfile(drawing_rel_input_path, os.path.join(tmpdir, drawing_rel_output_path))
        shutil.copyfile(logo_input_path, os.path.join(tmpdir, logo_output_path))

        # step 1b: add the printersettings:
        os.makedirs(os.path.join(tmpdir, printer_settings_dir_path))
        shutil.copyfile(printer_settings_input_path, os.path.join(tmpdir, printer_settings_output_path))

        # step 2: add entries to manifest file

        #   <Override PartName="/xl/drawings/drawing2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml" />
        #     <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"/>

        # 2a: read manifest:
        manifest_path = os.path.join(tmpdir, manifest_file_path)
        manifest_doc = etree.parse(open(manifest_path))

        # 2b: add entry for drawing2 and the printer setting
        manifest_root = manifest_doc.getroot()
        manifest_root.append(etree.Element("Override", PartName="/xl/drawings/drawing2.xml",
                                           ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"))
        manifest_root.append(etree.Element("Default", Extension="bin",
                                           ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"))

        # and for the image:
        manifest_root.append(etree.Element("Default", Extension="jpeg", ContentType="image/jpeg"))

        # 2c: write it back out
        manifest_doc.write(manifest_path)

        # 3: add the drawing to the sheets.
        # 3a: for each sheet of interest:
        all_sheets = glob.glob(os.path.join(tmpdir, "xl/worksheets/*.xml"))
        for s in all_sheets:
            # we want to ignore sheets like Sheet1, Sheet2, etc.
            try:
                if xml_to_sheet_name[os.path.basename(s)].startswith("Sheet"):
                    continue
            except KeyError:
                raise Exception("Bad sheet xml to sheet name map")

            # which sheet are we on?
            m = re.match(SHEET_NAME_REGEX, os.path.basename(s))
            if len(m.groups()) != 1:
                raise Exception(f"Bad sheet name: {os.path.basename(s)}")
            sheet_num = m.groups()[0]

            # 3b: add a _rel entry for this sheet:
            shutil.copyfile(sheet_rel_template_path,
                            os.path.join(tmpdir, f"xl/worksheets/_rels/sheet{sheet_num}.xml.rels"))

            # Now, get ready to do some surgery on this worksheet:
            worksheet_doc = etree.parse(open(s))
            worksheet_root = worksheet_doc.getroot()

            # 3c: add a <Drawing/> node to this worksheet

            # <drawing r:id="rId2" />
            # r's namespace: xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            id_attr = QName('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id')
            drawing_element = Element("drawing", nsmap={
                'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
            })

            drawing_element.set(id_attr, 'rId2')
            worksheet_root.append(drawing_element)

            # 3d: write it out:
            worksheet_doc.write(s)

        # put things together again:
        z2 = zipfile.ZipFile(output_fname, "w", compression=zipfile.ZIP_DEFLATED)

        for root, subdirs, files in os.walk(tmpdir):
            for f in files:
                full_path = os.path.join(root, f)
                arcname = full_path.replace(f"{tmpdir}/", '')
                z2.write(full_path, arcname=arcname)

        z2.close()
