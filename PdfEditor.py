from PyPDF2 import PdfFileWriter, PdfFileReader
import sys
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, letter
from reportlab.lib.styles import getSampleStyleSheet
import configparser
import PdfDownload
from Constants import *
from tableauserverclient.server.endpoint.exceptions import ServerResponseError


config = configparser.ConfigParser()


# create a text object to print
def create_text_object(can, x, y, text, style):

    font = (style.fontName, style.fontSize, style.leading)
    lines = text.split("\n")
    offsets = [0] * len(lines)
    tx = can.beginText(x, y)
    tx.setFont(*font)
    for offset, line in zip(offsets, lines):
        tx.setXPos(offset)
        tx.textLine(line)
        tx.setXPos(-offset)
    return tx


# draw header, footer and logo in a white canvas
def create_layer(header, footer, logo, page_nr, orientation, mode):

    packet = io.BytesIO()
    if mode == LOCAL:
        can = canvas.Canvas(packet, pagesize=A4)
        width, height = A4
    else:
        can = canvas.Canvas(packet, pagesize=letter)
        width, height = letter

    # swap dimensions if orientation is landscape
    if orientation == LANDSCAPE:
        temp = height
        height = width
        width = temp

    # font style
    style = getSampleStyleSheet()['Normal']
    style.fontName = "Helvetica"
    style.fontSize = 15
    style.leading = 14.4

    # if footer is empty draw page number, otherwise draw the footer text
    if footer == "":
        footer = "Page "+str(page_nr)
    can.drawText(create_text_object(can, 30, 20, footer, style))

    # draw header if present
    if header != "":
        can.drawText(create_text_object(can, 30, height - 23, header, style))

    # draw the logo image if present
    if logo != "":
        can.drawImage(logo, width - 140, height - 45, 120, 40, 'auto')

    can.save()
    packet.seek(0)
    return PdfFileReader(packet)


def main():
    # get settings

    try:
        config.read(SETTINGS_CONFIGURATION_FILE)
        output_file = config.get(SETTINGS, OUTPUT)
        header = config.get(SETTINGS, HEADER)
        footer = config.get(SETTINGS, FOOTER)
        logo = config.get(SETTINGS, LOGO)
        orientation = config.get(SETTINGS, ORIENTATION)
        mode = config.get(SETTINGS, MODE)
    except configparser.NoOptionError as e:
        print(str(e)+". Check configuration in settings.ini")
        return -1
    if mode == LOCAL:
        try:
            input_file = config.get(LOCAL, FILE)
        except configparser.NoOptionError as e:
            print(str(e) + ". Check configuration in settings.ini")
            return -1
    else:
        try:
            workbook = config.get(SERVER, WORKBOOK)
            view = config.get(SERVER, VIEW)
            PdfDownload.download(workbook, view)
        except configparser.NoOptionError as e:
            print(str(e) + ". Check configuration in settings.ini")
            return -1
        except ServerResponseError:
            print("Sign in error. Check your auth parameters in settings.ini")
            return -1
        except RuntimeError as e:
            print(e)
            return -1
        input_file = temp_filepath + view + ".pdf"

    # input and output file must be present
    if input_file == "" or output_file == "":
        print("Specify source and destination pdf files in settings.ini")
        return -1

    try:
        original = PdfFileReader(open(input_file, "rb"))
    except FileNotFoundError:
        print("File not found: "+input_file)
        return -1
    file_writer = PdfFileWriter()

    page_number = original.getNumPages()

    # for each page
    for i in range(page_number):
        # create the canvas with the info
        upper_layer = create_layer(header, footer, logo, i + 1, orientation, mode)
        page = original.getPage(i)

        # merge the info with the original page
        page.mergePage(upper_layer.getPage(0))
        file_writer.addPage(page)
        print("Edited " + input_file + " - page nr: " + str(i + 1))

    print("save file -> " + output_file)

    try:
        output_stream = open(output_file, "wb")
    except FileNotFoundError:
        print("File not found: "+input_file)
        return -1

    # print to output file
    file_writer.write(output_stream)
    output_stream.close()
    return 0

if __name__ == '__main__':
    sys.exit(main())
