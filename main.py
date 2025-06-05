import threading
import itertools
import sys
import time

def spinner(event):
    for c in itertools.cycle('|/-\\'):
        if event.is_set():
            break
        sys.stdout.write(f'\r🔄 Brute-forcing password... {c}')
        sys.stdout.flush()
        time.sleep(0.1)
    sys.stdout.write('\r✅ Done.                          \n')



import multiprocessing
import win32com.client
import os
excel_file = r'D:\Lost\test1.xlsx'
# excel_file = r'C:\Users\lost\Desktop\test_u.xlsx'

def try_passwords(start, end, found_event):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False

    for p in range(start, end):
        if found_event.is_set():
            break  # Stop if another process succeeded

        password = f"{p:04d}"  # Convert 7 → "0007", 45 → "0045"
        try:
            wb = excel.Workbooks.Open(excel_file, False, True, None, password)
            wb.Unprotect(password)
            print(f"\n[PID {os.getpid()}] ✅ Success: {password}")
            wb.Close(False)
            excel.Quit()
            found_event.set()  # Notify others
            return
        except:
            continue

    excel.Quit()

if __name__ == '__main__':
    cpu_count = multiprocessing.cpu_count()
    ranges = []
    total_passwords = 10000  # "0000" to "9999"
    # total_passwords = 10

    chunk_size = total_passwords // cpu_count
    start = 0

    found_event = multiprocessing.Event()

    for i in range(cpu_count):
        end = start + chunk_size
        if i == cpu_count - 1:
            end = 10000 # Make sure we include "9999"
        ranges.append((start, end))
        start = end
    spinner_done = threading.Event()
    t = threading.Thread(target=spinner, args=(spinner_done,))
    t.start()

    processes = []
    for start, end in ranges:
        p = multiprocessing.Process(target=try_passwords, args=(start, end, found_event))
        p.start()
        processes.append(p)

    for p in processes:
        p.join()
    spinner_done.set()
