from pathlib import Path

import docx2pdf
import openpyxl
import wx
from docx import Document
from docx.shared import Pt


NAME_PARAGRAPH_INDEX = 0
NAME_RUN_INDEX = 0
FONT_NAME = "Calibri"
FONT_SIZE_PT = 16

HEADER_ROW_INDEX = 0
ID_COLUMN_INDEX = 0
NAME_COLUMN_INDEX = 1


def sanitize_certificate_id(certificate_id):
    return str(certificate_id).replace("+", "").strip()


def generate_pdf(template_path, output_dir, certificate_id, participant_name):
    doc = Document(template_path)
    run = doc.paragraphs[NAME_PARAGRAPH_INDEX].runs[NAME_RUN_INDEX]

    run.font.name = FONT_NAME
    run.font.size = Pt(FONT_SIZE_PT)

    display_name = str(participant_name or "").strip()
    run.add_text(display_name.title() if display_name else " ")

    clean_id = sanitize_certificate_id(certificate_id)
    doc_file = output_dir / f"{clean_id}.docx"
    pdf_file = output_dir / f"{clean_id}.pdf"

    doc.save(doc_file)
    docx2pdf.convert(str(doc_file), str(pdf_file))
    doc_file.unlink(missing_ok=True)


def process_certificates(template_path, excel_path, output_dir, log_callback=None):
    workbook = openpyxl.open(excel_path)
    count = 0

    for row_index, row in enumerate(workbook.active.rows):
        if row_index == HEADER_ROW_INDEX:
            continue

        certificate_id = row[ID_COLUMN_INDEX].value
        if not certificate_id:
            continue

        participant_name = row[NAME_COLUMN_INDEX].value
        if log_callback:
            log_callback(f"Processing: {certificate_id} -> {participant_name}")

        generate_pdf(template_path, output_dir, certificate_id, participant_name)
        count += 1

    return count


class CertificateFrame(wx.Frame):
    def __init__(self):
        super().__init__(parent=None, title="Certificate Maker", size=(760, 460))

        panel = wx.Panel(self)
        main_sizer = wx.BoxSizer(wx.VERTICAL)

        self.template_ctrl = self._build_path_row(panel, main_sizer, "Template (.docx)", self.on_browse_template)
        self.excel_ctrl = self._build_path_row(panel, main_sizer, "Excel (.xlsx)", self.on_browse_excel)
        self.output_ctrl = self._build_path_row(panel, main_sizer, "Output Folder", self.on_browse_output)

        self.generate_btn = wx.Button(panel, label="Generate PDFs")
        self.generate_btn.Bind(wx.EVT_BUTTON, self.on_generate)
        main_sizer.Add(self.generate_btn, 0, wx.ALL | wx.ALIGN_LEFT, 12)

        self.log_ctrl = wx.TextCtrl(panel, style=wx.TE_MULTILINE | wx.TE_READONLY)
        main_sizer.Add(self.log_ctrl, 1, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 12)

        panel.SetSizer(main_sizer)
        self.Centre()

    def _build_path_row(self, panel, parent_sizer, label, handler):
        row = wx.BoxSizer(wx.HORIZONTAL)
        row.Add(wx.StaticText(panel, label=label), 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 10)

        ctrl = wx.TextCtrl(panel)
        row.Add(ctrl, 1, wx.RIGHT, 8)

        browse_btn = wx.Button(panel, label="Browse")
        browse_btn.Bind(wx.EVT_BUTTON, handler)
        row.Add(browse_btn, 0)

        parent_sizer.Add(row, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP, 12)
        return ctrl

    def log(self, message):
        self.log_ctrl.AppendText(f"{message}\n")

    def on_browse_template(self, _event):
        with wx.FileDialog(self, "Select template", wildcard="Word file (*.docx)|*.docx", style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as dialog:
            if dialog.ShowModal() == wx.ID_OK:
                self.template_ctrl.SetValue(dialog.GetPath())

    def on_browse_excel(self, _event):
        with wx.FileDialog(self, "Select excel", wildcard="Excel file (*.xlsx)|*.xlsx", style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as dialog:
            if dialog.ShowModal() == wx.ID_OK:
                self.excel_ctrl.SetValue(dialog.GetPath())

    def on_browse_output(self, _event):
        with wx.DirDialog(self, "Select output folder", style=wx.DD_DEFAULT_STYLE | wx.DD_DIR_MUST_EXIST) as dialog:
            if dialog.ShowModal() == wx.ID_OK:
                self.output_ctrl.SetValue(dialog.GetPath())

    def on_generate(self, _event):
        template = Path(self.template_ctrl.GetValue().strip())
        excel = Path(self.excel_ctrl.GetValue().strip())
        output = Path(self.output_ctrl.GetValue().strip())

        if not template.is_file():
            wx.MessageBox("Please select a valid .docx template file.", "Validation", wx.ICON_WARNING)
            return
        if not excel.is_file():
            wx.MessageBox("Please select a valid .xlsx file.", "Validation", wx.ICON_WARNING)
            return
        if not output.is_dir():
            wx.MessageBox("Please select a valid output folder.", "Validation", wx.ICON_WARNING)
            return

        self.generate_btn.Disable()
        self.log("Starting generation...")

        try:
            total = process_certificates(template, excel, output, self.log)
            self.log(f"Completed. Generated {total} certificates.")
            wx.MessageBox(f"Done. Generated {total} certificates.", "Success", wx.ICON_INFORMATION)
        except Exception as exc:
            self.log(f"Error: {exc}")
            wx.MessageBox(str(exc), "Error", wx.ICON_ERROR)
        finally:
            self.generate_btn.Enable()


def main():
    app = wx.App(False)
    frame = CertificateFrame()
    frame.Show()
    app.MainLoop()


if __name__ == "__main__":
    main()
