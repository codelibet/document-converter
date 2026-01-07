#!/usr/bin/env python3
import gi
gi.require_version("Gtk", "3.0")
from gi.repository import Gtk, GLib

import os
import threading

from PIL import Image
from pypdf import PdfReader, PdfWriter
from pdf2image import convert_from_path
from pdf2docx import Converter as PDF2DOCX

from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas


class Converter:

    def __init__(self):
        self.files = []

        builder = Gtk.Builder()
        builder.add_from_file("main.ui")

        # widgets
        self.window = builder.get_object("main_window")
        self.combo_src = builder.get_object("combo_original")
        self.combo_dst = builder.get_object("combo_convert")
        self.entry_files = builder.get_object("entry_files")
        self.progress = builder.get_object("progress_bar")

        self.btn_select = builder.get_object("button_select")
        self.btn_convert = builder.get_object("button_convert")

        # populate formats
        for fmt in ("Image", "PDF", "DOCX"):
            self.combo_src.append_text(fmt)
            self.combo_dst.append_text(fmt)

        self.combo_src.set_active(0)
        self.combo_dst.set_active(1)

        # signals
        self.btn_select.connect("clicked", self.on_select_files)
        self.btn_convert.connect("clicked", self.on_convert)

        self.window.connect("destroy", Gtk.main_quit)
        self.window.show_all()

    # FILE SELECTION

    def on_select_files(self, _):
        dialog = Gtk.FileChooserDialog(
            title="Select files",
            parent=self.window,
            action=Gtk.FileChooserAction.OPEN,
            buttons=(
                Gtk.STOCK_CANCEL, Gtk.ResponseType.CANCEL,
                Gtk.STOCK_OPEN, Gtk.ResponseType.OK
            )
        )
        dialog.set_select_multiple(True)

        if dialog.run() == Gtk.ResponseType.OK:
            self.files = dialog.get_filenames()
            self.entry_files.set_text(f"{len(self.files)} file(s) selected")
        else:
            self.files = []
            self.entry_files.set_text("")

        dialog.destroy()


    # CONVERT


    def on_convert(self, _):
        if not self.files:
            self.error("No files selected")
            return

        src = self.combo_src.get_active_text()
        dst = self.combo_dst.get_active_text()

        if not src or not dst:
            self.error("Select source and target formats")
            return

        self.progress.set_fraction(0.0)
        self.progress.set_text("Working…")

        threading.Thread(
            target=self.worker,
            args=(src, dst),
            daemon=True
        ).start()


    # WORKER


    def worker(self, src, dst):
        try:
            if src == "Image" and dst == "PDF":
                self.images_to_pdf()

            elif src == "PDF" and dst == "Image":
                self.pdf_to_images()

            elif src == "PDF" and dst == "PDF":
                if len(self.files) > 1:
                    self.merge_pdfs()
                else:
                    self.split_pdf()

            elif src == "PDF" and dst == "DOCX":
                self.pdf_to_docx()

            elif src == "DOCX" and dst == "PDF":
                self.docx_to_pdf()

            else:
                raise RuntimeError("Unsupported conversion")

            GLib.idle_add(self.done)

        except Exception as e:
            GLib.idle_add(self.error, str(e))

    # CONVERSIONS


    def images_to_pdf(self):
        imgs = [Image.open(f).convert("RGB") for f in self.files]
        imgs[0].save(self.out("_images.pdf"), save_all=True, append_images=imgs[1:])

    def pdf_to_images(self):
        for f in self.files:
            pages = convert_from_path(f)
            base = os.path.splitext(f)[0]
            for i, img in enumerate(pages, 1):
                img.save(f"{base}_page_{i}.png")

    def merge_pdfs(self):
        w = PdfWriter()
        for f in self.files:
            r = PdfReader(f)
            for p in r.pages:
                w.add_page(p)
        with open(self.out("_merged.pdf"), "wb") as fp:
            w.write(fp)

    def split_pdf(self):
        reader = PdfReader(self.files[0])

        dialog = Gtk.Dialog(
            title="Page range",
            parent=self.window,
            buttons=(
                Gtk.STOCK_CANCEL, Gtk.ResponseType.CANCEL,
                Gtk.STOCK_OK, Gtk.ResponseType.OK
            )
        )

        box = dialog.get_content_area()
        entry = Gtk.Entry()
        entry.set_placeholder_text("e.g. 1-3,5,7")

        box.add(Gtk.Label(label=f"Pages (1–{len(reader.pages)})"))
        box.add(entry)
        dialog.show_all()

        if dialog.run() != Gtk.ResponseType.OK:
            dialog.destroy()
            return

        text = entry.get_text()
        dialog.destroy()

        pages = self.parse_ranges(text, len(reader.pages))

        w = PdfWriter()
        for i in pages:
            w.add_page(reader.pages[i])

        with open(self.out("_range.pdf"), "wb") as fp:
            w.write(fp)

    def pdf_to_docx(self):
        for f in self.files:
            out = os.path.splitext(f)[0] + ".docx"
            cv = PDF2DOCX(f)
            cv.convert(out)
            cv.close()

    def docx_to_pdf(self):
        for f in self.files:
            doc = Document(f)
            out = os.path.splitext(f)[0] + ".pdf"

            c = canvas.Canvas(out, pagesize=A4)
            width, height = A4
            y = height - 40

            for p in doc.paragraphs:
                if y < 40:
                    c.showPage()
                    y = height - 40
                c.drawString(40, y, p.text)
                y -= 14

            c.save()


    # HELPERS


    def parse_ranges(self, text, total):
        pages = set()
        for part in text.split(","):
            if "-" in part:
                a, b = map(int, part.split("-"))
                pages.update(range(a - 1, b))
            else:
                pages.add(int(part) - 1)
        return sorted(pages)

    def out(self, suffix):
        return os.path.splitext(self.files[0])[0] + suffix


    # UI FEEDBACK


    def done(self):
        self.progress.set_fraction(1.0)
        self.progress.set_text("Done")

    def error(self, msg):
        d = Gtk.MessageDialog(
            parent=self.window,
            message_type=Gtk.MessageType.ERROR,
            buttons=Gtk.ButtonsType.OK,
            text=str(msg)
        )
        d.run()
        d.destroy()


if __name__ == "__main__":
    Converter()
    Gtk.main()

