import win32com.client
import pywintypes
try:
    word = win32com.client.Dispatch("Word.Application")
    word_file = r'C:\Users\lost\Downloads\test.docx'
    wb = word.Documents.Open(word_file, False, True, None, "AAAa")
    print("ok")
    wb.Close(False)
    word.Quit()
except pywintypes.com_error as e:
    print("error")