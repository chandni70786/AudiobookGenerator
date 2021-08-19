# Importing tkinter to make gui in python 


from typing import Text
from PySimpleGUI.PySimpleGUI import Output, Print
import pygame
from PIL import Image
import pytesseract
import os
import glob
import PySimpleGUI as sg
import fitz
import pyttsx3
# Importing tkPDFViewer to place pdf file in gui. 
# In tkPDFViewer library there is an tkPDFViewer module, that I have imported as pdf 
from tkPDFViewer import tkPDFViewer as pdf 
import PyPDF4
import sys
import os
import win32com.client


# Initialize the Pytesseract OCR software
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


def check_file_type_convert(current_directory,file):
    # If it is .pdf 
    if file.endswith('.pdf') or file.endswith(".png") or file.endswith(".jpg") or file.endswith("jpeg") or file.endswith(".tiff") or file.endswith(".bmp"):
        return file

    # If it is .txt file it coverts it to pdf
    if file.endswith('.txt'):
        
        # Python program to convert text file to PDF using FPDF
        from fpdf import FPDF 
        pdf = FPDF(orientation = 'P', unit = 'mm', format='A4')
        # Portrait, millimeter units, A4 page size     
        pdf.add_page()  
        f = open(file, "r") 
        for x in f: 
            pdf.cell(200, 10, txt = x, ln = 1)
            
        pdf.output(current_directory + "\output.pdf") 
        print(".txt to PDF conversion sucessful and Saved")
        return current_directory+"\output.pdf"

    # If it is .xls or .xlsx or .csv file it coverts it to pdf
    # XLS, XLSX, XLSM, XLTX and XLTM
    if file.endswith('.xls') or file.endswith('.xlsx') or file.endswith('.xlsm') or file.endswith('.xltx')or file.endswith('.csv') or file.endswith('.xml'):
        xlApp = win32com.client.Dispatch("Excel.Application")
        books = xlApp.Workbooks.Open(file)
        ws = books.Worksheets[0]
        ws.Visible = 1
        out_file= current_directory+"\output.pdf"
        ws.ExportAsFixedFormat(0,out_file )
        print("Exel to PDF conversion sucessful and Saved")
        return current_directory+"\output.pdf"

    # If it is .doc or .docx file it coverts it to pdf
    if file.endswith('.docx') or file.endswith('.doc'):
             
        wdFormatPDF = 17
        
        word = win32com.client.Dispatch('Word.Application')
        doc = word.Documents.Open(file)
        doc.SaveAs(current_directory+"\output.pdf", FileFormat=wdFormatPDF)
        doc.Close()
        word.Quit()
        print(".DOCX to PDF conversion sucessful and Saved")
        return current_directory+"\output.pdf"

    # If it is .ppt or .pptx file it coverts it to pdf
    if file.endswith('.pptx') or file.endswith('.ppt') or file.endswith('.pptm') or file.endswith('.ppsx') or file.endswith('.ppsm') or file.endswith('.pps') or file.endswith('.potx') or file.endswith('.ppa'):

        powerpoint=win32com.client.Dispatch("Powerpoint.Application")
        pdf=powerpoint.Presentations.Open(file,WithWindow=False)
        pdf.SaveAs(current_directory+"\output.pdf",32)
        pdf.Close()
        powerpoint.Quit()
        print(".PPTX to .PDF conversion sucessful and Saved")
        return current_directory+"\output.pdf"


global ply_count,pdf_to_read

global e

# In this function we get first and last page, which we want the software to read
def get_text(value, filepath):

    string = value
    string = string.strip()
    
    if "-" in string:
        first_page_number = int(string.split("-")[0])
        last_page_number = int(string.split("-")[1])
    elif string =='Full'or string =='full' or string =='FULL':
        first_page_number = 1
        last_page_number = PyPDF4.PdfFileReader(filepath).getNumPages()
    else:
        first_page_number = int(string)
        last_page_number = 0

    return first_page_number,last_page_number

def main():
    pygame.mixer.init()
    engine = pyttsx3.init()
    voices = engine.getProperty('voices')
    
   
    ply_count = -1

    ##### Create directory for Text to speech software
    current_directory = os.getcwd()
    final_directory = os.path.join(current_directory,r'Text_to_speech_software')
    if not os.path.exists(final_directory):
        os.makedirs(final_directory)
    print("current_directory is ",current_directory)
    print("final_directory is ",final_directory )

    ############################################################ GUI Part #############################################################

    # All the stuff inside your window.
    sg.theme('DarkBlue14')
    layout = [  [sg.Text('Choose File to read -'),sg.Input(do_not_clear=True),sg.FileBrowse()],
                [sg.Text('Total No. of Pages are - '),sg.Text(size=(5,1), key='-OUTPUT-'),sg.Button('CHECK')],################### [sg.Text('Total No. of Pages are - ',key='-NO-')],
                [sg.Text('Enter Page number or range separated by - ',), sg.InputText(default_text=1,key='range_s')],
                [[sg.Text('Voice -',), sg.Combo(list(voice.name for voice in voices), enable_events=True, key='Speech_lang')],[sg.Text('Language -',), sg.Combo(list(lang for lang in pytesseract.get_languages(config='')), enable_events=True, key='txt_lang',default_value='eng')]],
                [sg.Button('Ok'), sg.Button('Cancel'), sg.Button('About')]
            ]

    # Create the GUI Window Prompt
    window = sg.Window('KINDLY ENTER YOUR PREFERENCES HERE', layout)
    sg.theme('DarkBlue14')
    valid = False
    win2_active = False
    # Event Loop to process "events" and get the "values" of the inputs
    while True:
        event, values = window.read(timeout=100)
        # Here we read the path of the pdf file
        pdf_to_read = values[0]
        if event == 'CHECK':
            pdf_to_read = check_file_type_convert(current_directory,pdf_to_read)
            print("This folder path:  ",pdf_to_read)
            if pdf_to_read.endswith(".png") or pdf_to_read.endswith(".jpg") or pdf_to_read.endswith("jpeg") or pdf_to_read.endswith(".tiff") or pdf_to_read.endswith(".bmp"):
                window['-OUTPUT-'].update(1)
            else:
                window['-OUTPUT-'].update(PyPDF4.PdfFileReader(pdf_to_read).getNumPages())
        
    
        if not win2_active and event == 'About':
            win2_active = True
            sg.theme('LightBlue3')
            layout2 = [[sg.Text('An AUDIOBOOK is a recording or voiceover of a book or other work \nread aloud.It is a customizable narrator–specific sentence tokenizer \nthat allows for unlimited lengths of text to be read, all while keeping\nproper intonation, abbreviations, decimals and more into normal\nhuman like speech.\n\nEven though the technologies have been improved to quite a greater\nextent,the way we are learning is still in a old fashioned way. So to \nbreak this gap and move up with the current technological advance-\nment,we have come up with an idea of creating an audiobook \ngenerator where-in it takes input file of any sort and converts into\nspeech of human language.\n\n\n\n\n\n\nBy,\nChaitanya Sachidanand\nChandni Kumari\nAnusha B')],
                    [sg.Button('Exit')]]

            win2 = sg.Window('ABOUT AUDIOBOOK', layout2)

        if win2_active:
            ev2, vals2 = win2.read(timeout=100)
            if ev2 == sg.WIN_CLOSED or ev2 == 'Exit':
                win2_active  = False
                win2.close()
                

        # We will get the total page count and it gets updated in the window
        
        if values == True:
            print('How many pages',len(fitz.open(values[0])))
            window['-NO-'].Update(value=len(fitz.open(values[0])))

        if event in (None, 'Cancel'):	# if user closes window or clicks cancel
            print("Exitting")
            window.close()
            exit()
        
        if event == "Ok":

            if values[0] == "":
                sg.Popup("Enter value", "Enter PDF file to be transcribed ")
            if values['range_s'] == "":
                sg.Popup("Enter value", "Enter page number(s) to be transcribed")

            if values[0]!="" and values['range_s']!="":
                for char in values['range_s']:
                    if values['range_s']=='Full'or 'full' or 'FULL':
                        valid=True
                        break
                    if char.isdigit()==False:
                        sg.Popup("Invalid value","Enter valid number or numbers separated by -")
                        break
                    else:
                        valid=True
                        break
        # Break while loop if valid first and last page numbers received
        if valid==True:
            print('You entered ', values['range_s'])
            break
    
    pdf_to_read = check_file_type_convert(current_directory,values[0])
    Speech_lang = values['Speech_lang']  # use the combo key
    Txt_lang = values['txt_lang'] 
    window.close()
    
    
    if pdf_to_read.endswith(".png") or pdf_to_read.endswith(".jpg") or pdf_to_read.endswith("jpeg") or pdf_to_read.endswith(".tiff") or pdf_to_read.endswith(".bmp"):
        No_of_pg=1
    else:

        No_of_pg=int(PyPDF4.PdfFileReader(pdf_to_read).getNumPages())
    
    
    
    first_page_number,last_page_number = get_text(values['range_s'],values[0])
    print("\n\n\n\n\n\nFist page :   ",first_page_number,"\n\nLast page no:  \n\n\n\n\n",last_page_number)
    
    
    # In this bunch of code, we get permission to delete the folder if it already exists, where we intend to save our PDF images and audio
    image_directory = glob.glob(final_directory)
    for file in os.listdir(final_directory):
        filepath = os.path.join(final_directory,file)
        print(filepath)
        os.chmod(filepath, 0o777)
        os.remove(filepath)

    print("This is the pdf which should open with path:",pdf_to_read)
    # Here we read desired PDF pages and store them as images in a folder
    doc = fitz.open(pdf_to_read)
    k=1

    print("\n\n\n\n\n\nThis is speech lang\n\n",Speech_lang)
    engine.setProperty('voice', Speech_lang)



    # If user wants to read a single page
    if last_page_number == 0:
        page = doc.loadPage(first_page_number-1) #number of page
        zoom_x = 2.0
        zoom_y = 2.0
        mat = fitz.Matrix(zoom_x,zoom_y)
        pix = page.getPixmap(matrix=mat)
        output = os.path.join(final_directory, r"image_1_to_read.png")
        pix.writePNG(output)

    # If user wants to read range of pages
    else:
        for i in range(first_page_number-1,last_page_number):
            page = doc.loadPage(i) #number of page
            zoom_x = 2.0
            zoom_y = 2.0
            mat = fitz.Matrix(zoom_x,zoom_y)
            pix = page.getPixmap(matrix=mat)
            output = os.path.join(final_directory, r"image_"+str(k)+"_to_read.png")
            pix.writePNG(output)
            k+=1

    print("Done")

    

    mytext = []

    
    #this is for the Layout Design of the Window
    layout = [[sg.Text('Please wait!!\nWhile the desired pages are getting loaded....')],
        [sg.Text('Pages loaded are :'),sg.Text(size=(5,1), key='-OUTPUT-'),sg.Text('of    %d' % last_page_number)],
              [sg.ProgressBar(1, orientation='h',size=(30, 30), key='progress')],
            ]
    #This Creates the Physical Window
    sg.theme('DarkBlue12')
    window = sg.Window('PROGRESS WINDOW', layout).Finalize()
    progress_bar = window.FindElement('progress')
   
    # Here we load the image(s) created in Text_to_speech folder and read the text in image via pytesseract Optical Character Recognition (OCR) software
    # thus reading text in images and giving us a string
    
    i=0
    pg_no_cou = first_page_number
    event, values = window.read(timeout=10)
    for i in range(len(os.listdir(final_directory))):
        print('_'*50,"\nCount of no of intems present in final dir: ",len(os.listdir(final_directory)),"\n",'_'*50)
        data = pytesseract.image_to_string(Image.open(os.path.join(final_directory, r"image_"+str(i+1)+"_to_read.png")),lang=Txt_lang)#)
        data = data.replace("|","I") # For some reason the image to text translation would put | instead of the letter I. So we replace | with I
        data = data.split('\n')
        mytext.append(data)
        if not progress_bar.UpdateBar(i, len(os.listdir(final_directory))):
         if not sg.OneLineProgressMeter('My 1-line progress meter', i+1, 2000, 'single'):
            break
        pg_no_cou+=1
        window['-OUTPUT-'].update(pg_no_cou)
        
    window.Close()


    print(mytext)

    # Here we make sure that the text is read correctly and we read it line by line. Because sometimes, text would end abruptly

    newtext= ""
    for text in mytext:
        for line in text:
            line = line.strip()
            # If line is small, ignore it
            if len(line.split(" ")) < 10 and len(line.split(" "))>0:
                newtext= newtext + " " + str(line) + "\n"

            elif len(line.split(" "))<2:
                pass
            else:
                if line[-1]!=".":
                    newtext = newtext + " " + str(line)
                else:
                    newtext = newtext + " " + line + "\n"

    print(newtext)

 
        
    doc = fitz.open(pdf_to_read)

    page_count = len(doc)
  
    # storage for page display lists
    dlist_tab = [None] * page_count

    title = "PyMuPDF display of '%s', pages: %i" % (pdf_to_read, page_count)


    def get_page(pno, zoom=0):
        """Return a PNG image for a document page number. If zoom is other than 0, one of the 4 page quadrants are zoomed-in instead and the corresponding clip returned.
        """
        dlist = dlist_tab[pno]  # get display list
        if not dlist:  # create if not yet there
            dlist_tab[pno] = doc[pno].getDisplayList()
            dlist = dlist_tab[pno]
        r = dlist.rect  # page rectangle
        mp = r.tl + (r.br - r.tl) * 0.5  # rect middle point
        mt = r.tl + (r.tr - r.tl) * 0.5  # middle of top edge
        ml = r.tl + (r.bl - r.tl) * 0.5  # middle of left edge
        mr = r.tr + (r.br - r.tr) * 0.5  # middle of right egde
        mb = r.bl + (r.br - r.bl) * 0.5  # middle of bottom edge
        mat = fitz.Matrix(2, 2)  # zoom matrix
        if zoom == 1:  # top-left quadrant
            clip = fitz.Rect(r.tl, mp)
        elif zoom == 4:  # bot-right quadrant
            clip = fitz.Rect(mp, r.br)
        elif zoom == 2:  # top-right
            clip = fitz.Rect(mt, mr)
        elif zoom == 3:  # bot-left
            clip = fitz.Rect(ml, mb)
        if zoom == 0:  # total page
            pix = dlist.getPixmap(alpha=False)
        else:
            pix = dlist.getPixmap(alpha=False, matrix=mat, clip=clip)
        return pix.getPNGData()  # return the PNG image

    # The Loop sets the voice which was selected in the list available before
    for voice in voices:
        if Speech_lang == voice.name:
            engine.setProperty("voice", voice.id)

    cur_page = first_page_number-1
    data = get_page(cur_page)  # show page 1 for start
    image_elem = sg.Image(data=data,)#size=(500,650))
    goto = sg.InputText(str(cur_page + 1), size=(5, 1))

    # Audio Control sg.Frame
    col2 = sg.Column([[sg.Frame('Audio Controls',[[
                sg.Button('Read loud'),
                sg.Button('Pause'),
                sg.Button('Stop'),
            ],
            [
                sg.FileSaveAs(
            key='fig_save',
            file_types=(('MP3', '.mp3'),('WAV', '.wav')),  # TODO: better names
            ),
                sg.Combo([0.1,0.175,0.25,0.5,0.75,1,1.25,1.5,1.75,2] , default_value = 1,enable_events=True, key='Set_Speed',),
            ]])]],pad=(0,0))

    col1 = sg.Column([
        # PDF Control sg.Frame
        [sg.Frame('Pdf Controls',[[
                sg.Button('Prev'),
                sg.Button('Next'),
                sg.Text('Page:'),
                goto,
                
                sg.Text('/'+str(No_of_pg))
            ],
            [
                sg.Text("Zoom:"),
                sg.Button('Top-L'),
                sg.Button('Top-R'),
                sg.Button('Bot-L'),
                sg.Button('Bot-R'),
            ]],)],
        ], pad=(0,0))

    layout = [[col1, col2],
            [image_elem]]
    

    my_keys = ("Next", "Next:34", "Prev", "Prior:33", "Top-L", "Top-R",
            "Bot-L", "Bot-R", "MouseWheel:Down", "MouseWheel:Up")
    zoom_buttons = ("Top-L", "Top-R", "Bot-L", "Bot-R")



    window = sg.Window(title, layout,
                    return_keyboard_events=True, use_default_focus=True,auto_size_text=True,keep_on_top=True)
                   
    old_page = 0
    old_zoom = 0  # used for zoom on/off
    # the zoom buttons work in on/off mode.
    
    while True:
        event, values = window.read()#timeout=100)
        zoom = 0
        force_page = False
        if event == sg.WIN_CLOSED:    
            if os.path.exists("temp.wav"):
                os.remove("temp.wav")
            else:
                print("The file does not exist")
            break

        if event in ("Escape:27",):  # this spares me a 'Quit' button!
            
            break
        if event[0] == chr(13):  # surprise: this is 'Enter'!
            try:
                cur_page = int(values[0]) - 1  # check if valid
                while cur_page < 0:
                    cur_page += page_count
            except:
                cur_page = 0  # this guy's trying to fool me
            goto.update(str(cur_page + 1))
            

        elif event in ("Next", "Next:34", "MouseWheel:Down"):
            cur_page += 1
        elif event in ("Prev", "Prior:33", "MouseWheel:Up"):
            cur_page -= 1
        elif event == "Top-L":
            zoom = 1
        elif event == "Top-R":
            zoom = 2
        elif event == "Bot-L":
            zoom = 3
        elif event == "Bot-R":
            zoom = 4


#--------------------------$$$$$$$$$$$$$$$$$$$$$$$$$-------------------------#
        #Start telling audio of the page entered
        elif event in ("Pause") and ply_count == 1:
            #reads audio and plays it
            pygame.mixer.music.pause()
            ply_count = 0

        #Reads the audio and starts narrating
        elif event in ("Read loud") and ply_count == 0 :
            pygame.mixer.music.unpause()
            ply_count = 1

        elif event in ("Read loud") and ply_count == -1:
            outfile = "temp.wav"
            engine = pyttsx3.init()
            engine.setProperty("voice", Speech_lang)
            # getting details of current speaking rate
            #printing current voice rate
            engine.setProperty('rate', values['Set_Speed']*200)     # setting up new voice rate 
            rate = engine.getProperty('rate')
            print("The speech in the start of the reading ",rate)
            engine.save_to_file(newtext, outfile)#.get('1.0', END)
            engine.runAndWait()
            pygame.mixer.music.load(outfile)
            pygame.mixer.music.play()
            ply_count = 1

        # Stops the audio playing 
        elif event in ("Stop"):
            pygame.mixer.music.stop()
            pygame.mixer.music.unload()
            ply_count = -1


        # Sets the speed and starts speaking
        elif event in ('Set_Speed'):
            engine.setProperty('rate', values['Set_Speed']*200)
            print('\n',values['Set_Speed'])
            print(values['Set_Speed']*200)
            pygame.mixer.music.stop()
            pygame.mixer.music.unload()
            ply_count = -1
            

        elif values['fig_save']!=None:
            file_save = values['fig_save']
            print(file_save)
            engine.save_to_file(newtext, file_save)
            engine.runAndWait()


        # sanitize page number    
        if cur_page >= page_count:   # wrap around
            cur_page = 0
        while cur_page < 0:  # we show conventional page numbers
            cur_page += page_count

        # prevent creating same data again
        if cur_page != old_page:
            zoom = old_zoom = 0
            force_page = True

        if event in zoom_buttons:
            if 0 < zoom == old_zoom:
                zoom = 0
                force_page = True

            if zoom != old_zoom:
                force_page = True

        if force_page:
            data = get_page(cur_page, zoom)
            image_elem.update(data=data)
            old_page = cur_page
        old_zoom = zoom

        # update page number field
        if event in my_keys or not values[0]:
            goto.update(str(cur_page + 1)) 
        
  ############################################## GUI END ############################################

if __name__ == '__main__':
    main()